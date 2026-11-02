import os
import io
import csv
import socket
from functools import wraps
from datetime import datetime, date
from flask import Flask, render_template, request, jsonify, redirect, url_for, send_file, flash, session
from werkzeug.security import check_password_hash, generate_password_hash
from database import get_connection, init_db, get_next_invoice_number

app = Flask(__name__)
# Secure session key
app.secret_key = os.environ.get('SECRET_KEY', 'apex-wholesale-retail-secure-session-key-2026')

# Initialize DB with WAL mode and multi-user tables
init_db()

def get_local_ip():
    """Detect local Wi-Fi / LAN IP so staff phones/tablets can connect."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'

def get_settings():
    conn = get_connection()
    row = conn.execute("SELECT * FROM settings ORDER BY id ASC LIMIT 1").fetchone()
    conn.close()
    static_logo_file = os.path.join(os.path.dirname(__file__), 'static', 'img', 'logo.png')
    default_logo_url = '/static/img/logo.png' if os.path.exists(static_logo_file) else ''

    if row:
        d = dict(row)
        if not d.get('invoice_prefix'):
            d['invoice_prefix'] = 'SSIPL'
        if not d.get('logo_path') or len(d.get('logo_path', '').strip()) == 0:
            d['logo_path'] = default_logo_url
        return d
    return {
        "shop_name": "SSIPL Bags & Footwear Hub",
        "tagline": "Wholesale & Retail Dealers in Bags & Footwear",
        "phone": "+91 98765 43210",
        "address": "Market Complex",
        "tax_number": "",
        "currency_symbol": "₹",
        "invoice_prefix": "SSIPL",
        "logo_path": default_logo_url,
        "invoice_footer": "Thank you for shopping with us!"
    }

def get_current_user():
    user_id = session.get('user_id')
    if not user_id:
        return None
    conn = get_connection()
    user = conn.execute("SELECT id, username, full_name, role, phone FROM users WHERE id = ? AND is_active = 1", (user_id,)).fetchone()
    conn.close()
    return dict(user) if user else None

@app.context_processor
def inject_global_data():
    return {
        'settings': get_settings(),
        'current_user': get_current_user(),
        'local_ip': get_local_ip(),
        'now': datetime.now()
    }

# ----------------- AUTHENTICATION DECORATORS -----------------
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login', next=request.url))
        return f(*args, **kwargs)
    return decorated_function

def role_required(allowed_roles):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'user_id' not in session:
                return redirect(url_for('login', next=request.url))
            user = get_current_user()
            if not user:
                session.clear()
                return redirect(url_for('login'))
            if user['role'] not in allowed_roles and user['role'] != 'ADMIN':
                flash('Access Restricted: Your user account does not have permission for this section.', 'danger')
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return decorated_function
    return decorator

# ----------------- LOGIN / LOGOUT -----------------
@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user_id' in session and get_current_user():
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip().lower()
        password = request.form.get('password', '').strip()

        conn = get_connection()
        user = conn.execute("SELECT * FROM users WHERE username = ? AND is_active = 1", (username,)).fetchone()
        conn.close()

        if user and check_password_hash(user['password_hash'], password):
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['role'] = user['role']
            session['full_name'] = user['full_name']

            flash(f"Welcome, {user['full_name']}! Signed in as {user['role']}.", "success")
            next_url = request.args.get('next')
            if next_url and next_url.startswith('/'):
                return redirect(next_url)
            
            # Role based default landings
            if user['role'] == 'SALESMAN':
                return redirect(url_for('pos'))
            elif user['role'] == 'WAREHOUSE':
                return redirect(url_for('inventory'))
            return redirect(url_for('dashboard'))
        else:
            flash("Invalid username or password. Please verify your credentials.", "danger")

    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    flash("You have successfully signed out.", "success")
    return redirect(url_for('login'))

# ----------------- HOME / DASHBOARD -----------------
@app.route('/')
@login_required
def dashboard():
    user = get_current_user()
    conn = get_connection()
    cur = conn.cursor()

    # Product count & stock totals
    cur.execute("SELECT COUNT(*), COALESCE(SUM(stock_qty), 0) FROM products")
    total_products, total_stock_items = cur.fetchone()

    # Inventory Valuation (Restricted for Salesmen)
    stock_cost_val, stock_ws_val, stock_retail_val = 0, 0, 0
    if user['role'] in ('ADMIN', 'ACCOUNTING'):
        cur.execute("SELECT COALESCE(SUM(stock_qty * cost_price), 0), COALESCE(SUM(stock_qty * wholesale_price), 0), COALESCE(SUM(stock_qty * retail_price), 0) FROM products")
        stock_cost_val, stock_ws_val, stock_retail_val = cur.fetchone()

    # Low stock items
    cur.execute("SELECT * FROM products WHERE stock_qty <= min_alert_qty ORDER BY stock_qty ASC LIMIT 10")
    low_stock_items = [dict(r) for r in cur.fetchall()]
    low_stock_count = len(low_stock_items)

    # Today's Sales Summary
    cur.execute("""
        SELECT 
            COALESCE(SUM(grand_total), 0),
            COUNT(*),
            COALESCE(SUM(CASE WHEN sale_type = 'WHOLESALE' THEN grand_total ELSE 0 END), 0),
            COALESCE(SUM(CASE WHEN sale_type = 'RETAIL' THEN grand_total ELSE 0 END), 0)
        FROM sales 
        WHERE DATE(created_at) = DATE('now', 'localtime')
    """)
    today_sales_total, today_orders_count, today_wholesale_total, today_retail_total = cur.fetchone()

    # Today's estimated gross profit (Admin / Accounting only)
    today_profit = 0
    if user['role'] in ('ADMIN', 'ACCOUNTING'):
        cur.execute("""
            SELECT COALESCE(SUM(si.quantity * (si.unit_price - si.cost_price)), 0)
            FROM sale_items si
            JOIN sales s ON si.sale_id = s.id
            WHERE DATE(s.created_at) = DATE('now', 'localtime')
        """)
        today_profit = cur.fetchone()[0]

    # Total Outstanding Credit (Receivables from wholesale buyers)
    cur.execute("SELECT COALESCE(SUM(current_balance), 0) FROM parties WHERE current_balance > 0")
    total_receivable = cur.fetchone()[0]

    # Recent 6 Sales
    cur.execute("SELECT * FROM sales ORDER BY id DESC LIMIT 6")
    recent_sales = [dict(r) for r in cur.fetchall()]

    # Categories breakdown
    cur.execute("SELECT category, COUNT(*), SUM(stock_qty) FROM products GROUP BY category")
    category_breakdown = [dict(r) for r in cur.fetchall()]

    conn.close()

    return render_template('dashboard.html',
        total_products=total_products,
        total_stock_items=total_stock_items,
        stock_cost_val=round(stock_cost_val, 2),
        stock_ws_val=round(stock_ws_val, 2),
        stock_retail_val=round(stock_retail_val, 2),
        low_stock_items=low_stock_items,
        low_stock_count=low_stock_count,
        today_sales_total=round(today_sales_total, 2),
        today_orders_count=today_orders_count,
        today_wholesale_total=round(today_wholesale_total, 2),
        today_retail_total=round(today_retail_total, 2),
        today_profit=round(today_profit, 2),
        total_receivable=round(total_receivable, 2),
        recent_sales=recent_sales,
        category_breakdown=category_breakdown
    )

# ----------------- INVENTORY & PRODUCTS -----------------
@app.route('/inventory')
@login_required
@role_required(['ADMIN', 'SALESMAN', 'WAREHOUSE', 'ACCOUNTING'])
def inventory():
    category = request.args.get('category', '')
    status = request.args.get('status', '')
    search = request.args.get('search', '').strip()
    user = get_current_user()

    conn = get_connection()
    cur = conn.cursor()

    query = "SELECT * FROM products WHERE 1=1"
    params = []

    if category:
        query += " AND category = ?"
        params.append(category)

    if status == 'low':
        query += " AND stock_qty <= min_alert_qty"
    elif status == 'out':
        query += " AND stock_qty <= 0"
    elif status == 'in':
        query += " AND stock_qty > min_alert_qty"

    if search:
        query += " AND (name LIKE ? OR sku LIKE ? OR brand LIKE ? OR subcategory LIKE ? OR size LIKE ?)"
        term = f"%{search}%"
        params.extend([term, term, term, term, term])

    query += " ORDER BY category ASC, name ASC, size ASC"
    cur.execute(query, params)
    raw_products = [dict(r) for r in cur.fetchall()]

    # Hide cost price from salesmen so they cannot see supplier margins
    products = []
    for p in raw_products:
        p_dict = dict(p)
        if user['role'] == 'SALESMAN':
            p_dict['cost_price'] = None
        products.append(p_dict)

    conn.close()
    return render_template('inventory.html', products=products, category=category, status=status, search=search)

@app.route('/api/products/add', methods=['POST'])
@login_required
@role_required(['ADMIN', 'WAREHOUSE'])
def add_product():
    conn = get_connection()
    cur = conn.cursor()
    try:
        sku = request.form.get('sku', '').strip()
        name = request.form.get('name', '').strip()
        category = request.form.get('category', 'Bags')
        subcategory = request.form.get('subcategory', '').strip()
        brand = request.form.get('brand', '').strip()
        size = request.form.get('size', '').strip()
        color = request.form.get('color', '').strip()
        unit = request.form.get('unit', 'Pieces')
        cost_price = float(request.form.get('cost_price', 0))
        wholesale_price = float(request.form.get('wholesale_price', 0))
        retail_price = float(request.form.get('retail_price', 0))
        stock_qty = int(request.form.get('stock_qty', 0))
        min_alert_qty = int(request.form.get('min_alert_qty', 5))
        location = request.form.get('location', '')
        notes = request.form.get('notes', '')

        # Auto-generate SKU if empty
        if not sku:
            prefix = "BAG" if category == "Bags" else "SH"
            cur.execute("SELECT COUNT(*) FROM products WHERE category = ?", (category,))
            count = cur.fetchone()[0] + 1
            sku = f"{prefix}-{count:04d}"

        cur.execute("""
            INSERT INTO products (sku, name, category, subcategory, brand, size, color, unit, 
                                  cost_price, wholesale_price, retail_price, stock_qty, min_alert_qty, location, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (sku, name, category, subcategory, brand, size, color, unit,
              cost_price, wholesale_price, retail_price, stock_qty, min_alert_qty, location, notes))

        prod_id = cur.lastrowid
        if stock_qty > 0:
            cur.execute("""
                INSERT INTO stock_logs (product_id, change_type, quantity_change, stock_after, reference_no, reason)
                VALUES (?, 'MANUAL_ADJUSTMENT', ?, ?, 'OPENING', 'Initial Stock Entry by Staff')
            """, (prod_id, stock_qty, stock_qty))

        conn.commit()
        flash(f'Product "{name}" added successfully with SKU: {sku}!', 'success')
    except sqlite3.IntegrityError:
        flash(f'Error: SKU "{sku}" already exists! Please use a unique SKU.', 'danger')
    except Exception as e:
        flash(f'Error adding product: {str(e)}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('inventory'))

@app.route('/api/products/edit/<int:product_id>', methods=['POST'])
@login_required
@role_required(['ADMIN', 'WAREHOUSE'])
def edit_product(product_id):
    conn = get_connection()
    cur = conn.cursor()
    try:
        name = request.form.get('name', '').strip()
        category = request.form.get('category', 'Bags')
        subcategory = request.form.get('subcategory', '').strip()
        brand = request.form.get('brand', '').strip()
        size = request.form.get('size', '').strip()
        color = request.form.get('color', '').strip()
        unit = request.form.get('unit', 'Pieces')
        cost_price = float(request.form.get('cost_price', 0))
        wholesale_price = float(request.form.get('wholesale_price', 0))
        retail_price = float(request.form.get('retail_price', 0))
        stock_qty = int(request.form.get('stock_qty', 0))
        min_alert_qty = int(request.form.get('min_alert_qty', 5))
        location = request.form.get('location', '')
        notes = request.form.get('notes', '')

        cur.execute("SELECT stock_qty FROM products WHERE id = ?", (product_id,))
        old_stock = cur.fetchone()[0]

        cur.execute("""
            UPDATE products 
            SET name=?, category=?, subcategory=?, brand=?, size=?, color=?, unit=?,
                cost_price=?, wholesale_price=?, retail_price=?, stock_qty=?, min_alert_qty=?, location=?, notes=?,
                updated_at=CURRENT_TIMESTAMP
            WHERE id=?
        """, (name, category, subcategory, brand, size, color, unit,
              cost_price, wholesale_price, retail_price, stock_qty, min_alert_qty, location, notes, product_id))

        if stock_qty != old_stock:
            diff = stock_qty - old_stock
            cur.execute("""
                INSERT INTO stock_logs (product_id, change_type, quantity_change, stock_after, reference_no, reason)
                VALUES (?, 'MANUAL_ADJUSTMENT', ?, ?, 'EDIT', 'Manual Stock Edit')
            """, (product_id, diff, stock_qty))

        conn.commit()
        flash('Product updated successfully!', 'success')
    except Exception as e:
        flash(f'Error updating product: {str(e)}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('inventory'))

@app.route('/api/products/quick-adjust/<int:product_id>', methods=['POST'])
@login_required
@role_required(['ADMIN', 'WAREHOUSE'])
def quick_adjust_stock(product_id):
    conn = get_connection()
    cur = conn.cursor()
    try:
        qty_change = int(request.form.get('quantity_change', 0))
        reason = request.form.get('reason', 'Quick Stock Adjustment')

        cur.execute("SELECT stock_qty, name FROM products WHERE id = ?", (product_id,))
        row = cur.fetchone()
        if not row:
            return jsonify({'success': False, 'message': 'Product not found'}), 404

        new_stock = max(0, row[0] + qty_change)
        cur.execute("UPDATE products SET stock_qty = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (new_stock, product_id))
        cur.execute("""
            INSERT INTO stock_logs (product_id, change_type, quantity_change, stock_after, reference_no, reason)
            VALUES (?, 'MANUAL_ADJUSTMENT', ?, ?, 'ADJUST', ?)
        """, (product_id, qty_change, new_stock, reason))

        conn.commit()
        return jsonify({'success': True, 'new_stock': new_stock, 'message': f'Stock updated to {new_stock}'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500
    finally:
        conn.close()

@app.route('/api/products/delete/<int:product_id>', methods=['POST'])
@login_required
@role_required(['ADMIN'])
def delete_product(product_id):
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute("DELETE FROM products WHERE id = ?", (product_id,))
        conn.commit()
        flash('Product deleted successfully.', 'success')
    except Exception as e:
        flash(f'Cannot delete product (it may be linked to previous sales/purchases): {str(e)}', 'danger')
    finally:
        conn.close()
    return redirect(url_for('inventory'))

# ----------------- POS / BILLING (WHOLESALE & RETAIL) -----------------
@app.route('/pos')
@login_required
@role_required(['ADMIN', 'SALESMAN', 'ACCOUNTING'])
def pos():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT * FROM products ORDER BY name ASC")
    products = [dict(r) for r in cur.fetchall()]

    cur.execute("SELECT * FROM parties WHERE party_type IN ('WHOLESALE_BUYER', 'RETAIL_CUSTOMER') ORDER BY name ASC")
    parties = [dict(r) for r in cur.fetchall()]

    conn.close()
    return render_template('pos.html', products=products, parties=parties)

@app.route('/api/sales/create', methods=['POST'])
@login_required
@role_required(['ADMIN', 'SALESMAN', 'ACCOUNTING'])
def create_sale():
    data = request.get_json()
    if not data or not data.get('items'):
        return jsonify({'success': False, 'message': 'No items in cart'}), 400

    conn = get_connection()
    cur = conn.cursor()
    try:
        sale_type = data.get('sale_type', 'RETAIL').upper()
        customer_name = data.get('customer_name', 'Walk-in Customer').strip() or 'Walk-in Customer'
        customer_phone = data.get('customer_phone', '').strip()
        party_id = data.get('party_id') or None
        payment_mode = data.get('payment_mode', 'CASH').upper()
        notes = data.get('notes', '')

        subtotal = float(data.get('subtotal', 0))
        discount = float(data.get('discount', 0))
        grand_total = float(data.get('grand_total', 0))
        paid_amount = float(data.get('paid_amount', grand_total))
        balance_amount = max(0.0, grand_total - paid_amount)

        if balance_amount > 0:
            payment_status = 'PARTIAL' if paid_amount > 0 else 'CREDIT'
        else:
            payment_status = 'PAID'

        year_str = datetime.now().strftime("%Y")
        settings_data = get_settings()
        prefix = settings_data.get('invoice_prefix', 'SSIPL').strip() or 'SSIPL'
        
        # Concurrency-safe atomic invoice numbering (safe for 10+ simultaneous salesmen)
        invoice_no = get_next_invoice_number(conn, prefix=prefix, year_str=year_str)

        # Salesman attribution
        salesman_id = session.get('user_id')
        salesman_name = session.get('full_name') or session.get('username') or 'Staff'

        # Insert Sale record
        cur.execute("""
            INSERT INTO sales (invoice_no, sale_type, party_id, customer_name, customer_phone, 
                               subtotal, discount, grand_total, payment_mode, payment_status, 
                               paid_amount, balance_amount, notes, salesman_id, salesman_name)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (invoice_no, sale_type, party_id, customer_name, customer_phone,
              subtotal, discount, grand_total, payment_mode, payment_status,
              paid_amount, balance_amount, notes, salesman_id, salesman_name))
        sale_id = cur.lastrowid

        # Insert Sale Items & Deduct Stock
        for item in data['items']:
            prod_id = item['id']
            qty = int(item['quantity'])
            price = float(item['price'])
            item_total = qty * price

            cur.execute("SELECT stock_qty, cost_price, name, sku, size, color, unit FROM products WHERE id = ?", (prod_id,))
            prod = cur.fetchone()
            if not prod:
                raise Exception(f"Product ID {prod_id} not found!")

            current_stock = prod['stock_qty']
            cost_price = prod['cost_price']
            new_stock = current_stock - qty

            # Deduct stock
            cur.execute("UPDATE products SET stock_qty = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (new_stock, prod_id))

            cur.execute("""
                INSERT INTO sale_items (sale_id, product_id, product_name, sku, size, color, unit, quantity, unit_price, cost_price, total_price)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (sale_id, prod_id, prod['name'], prod['sku'], prod['size'], prod['color'], prod['unit'], qty, price, cost_price, item_total))

            cur.execute("""
                INSERT INTO stock_logs (product_id, change_type, quantity_change, stock_after, reference_no, reason)
                VALUES (?, 'SALE_OUT', ?, ?, ?, ?)
            """, (prod_id, -qty, new_stock, invoice_no, f"Sale to {customer_name} ({sale_type})"))

        if party_id and balance_amount > 0:
            cur.execute("UPDATE parties SET current_balance = current_balance + ? WHERE id = ?", (balance_amount, party_id))
        elif not party_id and balance_amount > 0 and customer_name != 'Walk-in Customer':
            cur.execute("""
                INSERT INTO parties (name, phone, party_type, current_balance)
                VALUES (?, ?, 'WHOLESALE_BUYER', ?)
            """, (customer_name, customer_phone, balance_amount))
            new_party_id = cur.lastrowid
            cur.execute("UPDATE sales SET party_id = ? WHERE id = ?", (new_party_id, sale_id))

        conn.commit()
        return jsonify({
            'success': True,
            'sale_id': sale_id,
            'invoice_no': invoice_no,
            'message': 'Sale completed successfully!'
        })
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 500
    finally:
        conn.close()

# ----------------- STOCK IN / BULK PURCHASES -----------------
@app.route('/purchases')
@login_required
@role_required(['ADMIN', 'WAREHOUSE'])
def purchases():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT * FROM purchases ORDER BY id DESC")
    purchase_list = [dict(r) for r in cur.fetchall()]

    cur.execute("SELECT * FROM products ORDER BY name ASC")
    products = [dict(r) for r in cur.fetchall()]

    cur.execute("SELECT * FROM parties WHERE party_type = 'SUPPLIER' ORDER BY name ASC")
    suppliers = [dict(r) for r in cur.fetchall()]

    conn.close()
    return render_template('purchases.html', purchases=purchase_list, products=products, suppliers=suppliers)

@app.route('/api/purchases/create', methods=['POST'])
@login_required
@role_required(['ADMIN', 'WAREHOUSE'])
def create_purchase():
    data = request.get_json()
    if not data or not data.get('items'):
        return jsonify({'success': False, 'message': 'No items specified in purchase'}), 400

    conn = get_connection()
    cur = conn.cursor()
    try:
        supplier_name = data.get('supplier_name', '').strip() or 'Direct Manufacturer'
        invoice_no = data.get('invoice_no', '').strip()
        purchase_date = data.get('purchase_date', date.today().isoformat())
        notes = data.get('notes', '')

        if not invoice_no:
            count = cur.execute("SELECT COUNT(*) FROM purchases").fetchone()[0] + 1
            invoice_no = f"PUR-{datetime.now().strftime('%Y%m')}-{count:04d}"

        total_amount = sum(float(item['quantity']) * float(item['unit_cost']) for item in data['items'])

        cur.execute("SELECT id FROM parties WHERE name = ? AND party_type = 'SUPPLIER'", (supplier_name,))
        sup_row = cur.fetchone()
        if sup_row:
            supplier_id = sup_row[0]
        else:
            cur.execute("INSERT INTO parties (name, party_type) VALUES (?, 'SUPPLIER')", (supplier_name,))
            supplier_id = cur.lastrowid

        cur.execute("""
            INSERT INTO purchases (invoice_no, supplier_id, supplier_name, purchase_date, total_amount, paid_amount, payment_status, notes)
            VALUES (?, ?, ?, ?, ?, ?, 'PAID', ?)
        """, (invoice_no, supplier_id, supplier_name, purchase_date, total_amount, total_amount, notes))
        purchase_id = cur.lastrowid

        for item in data['items']:
            prod_id = item['id']
            qty = int(item['quantity'])
            unit_cost = float(item['unit_cost'])
            tot_cost = qty * unit_cost

            cur.execute("SELECT stock_qty FROM products WHERE id = ?", (prod_id,))
            prod = cur.fetchone()
            if not prod:
                raise Exception(f"Product {prod_id} not found")

            old_stock = prod['stock_qty']
            new_stock = old_stock + qty

            cur.execute("""
                UPDATE products 
                SET stock_qty = ?, cost_price = ?, updated_at = CURRENT_TIMESTAMP 
                WHERE id = ?
            """, (new_stock, unit_cost, prod_id))

            cur.execute("""
                INSERT INTO purchase_items (purchase_id, product_id, quantity, unit_cost, total_cost)
                VALUES (?, ?, ?, ?, ?)
            """, (purchase_id, prod_id, qty, unit_cost, tot_cost))

            cur.execute("""
                INSERT INTO stock_logs (product_id, change_type, quantity_change, stock_after, reference_no, reason)
                VALUES (?, 'PURCHASE_IN', ?, ?, ?, ?)
            """, (prod_id, qty, new_stock, invoice_no, f"Bulk Purchase from {supplier_name}"))

        conn.commit()
        return jsonify({'success': True, 'message': 'Bulk stock received and inventory updated!'})
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 500
    finally:
        conn.close()

# ----------------- SALES HISTORY & INVOICES -----------------
@app.route('/sales')
@login_required
@role_required(['ADMIN', 'SALESMAN', 'ACCOUNTING'])
def sales():
    sale_type = request.args.get('sale_type', '')
    start_date = request.args.get('start_date', '')
    end_date = request.args.get('end_date', '')
    search = request.args.get('search', '').strip()

    conn = get_connection()
    cur = conn.cursor()

    query = "SELECT * FROM sales WHERE 1=1"
    params = []

    if sale_type:
        query += " AND sale_type = ?"
        params.append(sale_type)
    if start_date:
        query += " AND DATE(created_at) >= ?"
        params.append(start_date)
    if end_date:
        query += " AND DATE(created_at) <= ?"
        params.append(end_date)
    if search:
        query += " AND (invoice_no LIKE ? OR customer_name LIKE ? OR customer_phone LIKE ?)"
        term = f"%{search}%"
        params.extend([term, term, term])

    query += " ORDER BY id DESC"
    cur.execute(query, params)
    sales_list = [dict(r) for r in cur.fetchall()]

    conn.close()
    return render_template('sales.html', sales=sales_list, sale_type=sale_type, start_date=start_date, end_date=end_date, search=search)

@app.route('/invoice/<int:sale_id>')
@login_required
def print_invoice(sale_id):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT * FROM sales WHERE id = ?", (sale_id,))
    sale = cur.fetchone()
    if not sale:
        conn.close()
        return "Invoice not found", 404

    cur.execute("SELECT * FROM sale_items WHERE sale_id = ?", (sale_id,))
    items = [dict(r) for r in cur.fetchall()]

    party = None
    if sale['party_id']:
        cur.execute("SELECT * FROM parties WHERE id = ?", (sale['party_id'],))
        party_row = cur.fetchone()
        if party_row:
            party = dict(party_row)

    conn.close()
    return render_template('invoice_print.html', sale=dict(sale), items=items, party=party)

# ----------------- WHOLESALE PARTY LEDGER / KHATA -----------------
@app.route('/ledger')
@login_required
@role_required(['ADMIN', 'ACCOUNTING'])
def ledger():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT * FROM parties WHERE party_type = 'WHOLESALE_BUYER' ORDER BY current_balance DESC, name ASC")
    parties = [dict(r) for r in cur.fetchall()]

    cur.execute("SELECT COALESCE(SUM(current_balance), 0) FROM parties WHERE party_type = 'WHOLESALE_BUYER' AND current_balance > 0")
    total_due = cur.fetchone()[0]

    conn.close()
    return render_template('ledger.html', parties=parties, total_due=round(total_due, 2))

@app.route('/api/ledger/record-payment', methods=['POST'])
@login_required
@role_required(['ADMIN', 'ACCOUNTING'])
def record_party_payment():
    conn = get_connection()
    cur = conn.cursor()
    try:
        party_id = int(request.form.get('party_id'))
        amount = float(request.form.get('amount', 0))
        payment_mode = request.form.get('payment_mode', 'CASH')
        notes = request.form.get('notes', 'Payment Received')

        if amount <= 0:
            flash('Payment amount must be greater than 0', 'danger')
            return redirect(url_for('ledger'))

        cur.execute("SELECT name, current_balance FROM parties WHERE id = ?", (party_id,))
        party = cur.fetchone()
        if not party:
            flash('Party not found', 'danger')
            return redirect(url_for('ledger'))

        new_balance = max(0.0, party['current_balance'] - amount)
        cur.execute("UPDATE parties SET current_balance = ? WHERE id = ?", (new_balance, party_id))
        conn.commit()

        flash(f'Received {amount} from {party["name"]}. New pending balance: {new_balance}', 'success')
    except Exception as e:
        flash(f'Error recording payment: {str(e)}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('ledger'))

@app.route('/api/parties/add', methods=['POST'])
@login_required
@role_required(['ADMIN', 'ACCOUNTING'])
def add_party():
    conn = get_connection()
    cur = conn.cursor()
    try:
        name = request.form.get('name', '').strip()
        phone = request.form.get('phone', '').strip()
        party_type = request.form.get('party_type', 'WHOLESALE_BUYER')
        address = request.form.get('address', '').strip()
        gst_no = request.form.get('gst_no', '').strip()
        opening_balance = float(request.form.get('opening_balance', 0))

        cur.execute("""
            INSERT INTO parties (name, phone, party_type, address, gst_no, current_balance)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (name, phone, party_type, address, gst_no, opening_balance))
        conn.commit()
        flash(f'Party "{name}" added successfully!', 'success')
    except Exception as e:
        flash(f'Error adding party: {str(e)}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('ledger'))

# ----------------- STAFF / USER MANAGEMENT -----------------
@app.route('/users')
@login_required
@role_required(['ADMIN'])
def users_list():
    conn = get_connection()
    users = [dict(r) for r in conn.execute("SELECT id, username, full_name, role, phone, is_active, created_at FROM users ORDER BY role ASC, full_name ASC").fetchall()]
    conn.close()
    return render_template('users.html', users=users)

@app.route('/api/users/add', methods=['POST'])
@login_required
@role_required(['ADMIN'])
def add_user():
    username = request.form.get('username', '').strip().lower()
    password = request.form.get('password', '').strip()
    full_name = request.form.get('full_name', '').strip()
    role = request.form.get('role', 'SALESMAN').upper()
    phone = request.form.get('phone', '').strip()

    if not username or not password or not full_name:
        flash("Username, password, and full name are required.", "danger")
        return redirect(url_for('users_list'))

    conn = get_connection()
    try:
        conn.execute("""
            INSERT INTO users (username, password_hash, full_name, role, phone)
            VALUES (?, ?, ?, ?, ?)
        """, (username, generate_password_hash(password), full_name, role, phone))
        conn.commit()
        flash(f'Staff member "{full_name}" ({role}) created successfully!', 'success')
    except sqlite3.IntegrityError:
        flash(f'Username "{username}" is already taken. Please choose another.', 'danger')
    except Exception as e:
        flash(f'Error: {str(e)}', 'danger')
    finally:
        conn.close()
    return redirect(url_for('users_list'))

@app.route('/api/users/toggle/<int:user_id>', methods=['POST'])
@login_required
@role_required(['ADMIN'])
def toggle_user(user_id):
    if user_id == session.get('user_id'):
        flash("You cannot deactivate your own logged-in admin account!", "danger")
        return redirect(url_for('users_list'))
    conn = get_connection()
    conn.execute("UPDATE users SET is_active = CASE WHEN is_active = 1 THEN 0 ELSE 1 END WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()
    flash("Staff account status updated successfully.", "success")
    return redirect(url_for('users_list'))

# ----------------- EXPORT TO EXCEL / CSV -----------------
@app.route('/export/inventory/csv')
@login_required
@role_required(['ADMIN', 'ACCOUNTING'])
def export_inventory_csv():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT sku, name, category, subcategory, brand, size, color, unit, 
               cost_price, wholesale_price, retail_price, stock_qty, min_alert_qty, location, notes
        FROM products ORDER BY category, name, size
    """)
    rows = cur.fetchall()
    conn.close()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        'SKU', 'Product Name', 'Category', 'Subcategory', 'Brand', 'Size', 'Color', 'Unit',
        'Cost Price', 'Wholesale Price', 'Retail Price', 'Stock Qty', 'Min Alert Qty', 'Location', 'Notes'
    ])

    for row in rows:
        writer.writerow(list(row))

    output.seek(0)
    return send_file(
        io.BytesIO(output.getvalue().encode('utf-8-sig')),
        mimetype='text/csv',
        as_attachment=True,
        download_name=f'inventory_stock_{date.today().isoformat()}.csv'
    )

@app.route('/export/sales/csv')
@login_required
@role_required(['ADMIN', 'ACCOUNTING'])
def export_sales_csv():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT invoice_no, sale_type, salesman_name, customer_name, customer_phone, subtotal, discount, grand_total, 
               payment_mode, payment_status, paid_amount, balance_amount, created_at
        FROM sales ORDER BY id DESC
    """)
    rows = cur.fetchall()
    conn.close()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        'Invoice No', 'Sale Type', 'Salesman / Staff', 'Customer Name', 'Phone', 'Subtotal', 'Discount', 'Grand Total',
        'Payment Mode', 'Payment Status', 'Paid Amount', 'Balance Due', 'Date & Time'
    ])

    for row in rows:
        writer.writerow(list(row))

    output.seek(0)
    return send_file(
        io.BytesIO(output.getvalue().encode('utf-8-sig')),
        mimetype='text/csv',
        as_attachment=True,
        download_name=f'sales_report_{date.today().isoformat()}.csv'
    )

# ----------------- SETTINGS -----------------
@app.route('/settings', methods=['GET', 'POST'])
@login_required
@role_required(['ADMIN'])
def settings():
    conn = get_connection()
    cur = conn.cursor()

    if request.method == 'POST':
        import base64
        shop_name = request.form.get('shop_name', '').strip()
        tagline = request.form.get('tagline', '').strip()
        phone = request.form.get('phone', '').strip()
        email = request.form.get('email', '').strip()
        address = request.form.get('address', '').strip()
        tax_number = request.form.get('tax_number', '').strip()
        currency_symbol = request.form.get('currency_symbol', '₹').strip()
        invoice_prefix = request.form.get('invoice_prefix', 'SSIPL').strip() or 'SSIPL'
        invoice_footer = request.form.get('invoice_footer', '').strip()
        remove_logo = request.form.get('remove_logo') == '1'

        logo_file = request.files.get('logo')
        if logo_file and logo_file.filename:
            raw_bytes = logo_file.read()
            mimetype = logo_file.content_type or 'image/png'
            b64_str = base64.b64encode(raw_bytes).decode('utf-8')
            logo_data_uri = f"data:{mimetype};base64,{b64_str}"

            cur.execute("""
                UPDATE settings 
                SET shop_name=?, tagline=?, phone=?, email=?, address=?, tax_number=?, currency_symbol=?, invoice_prefix=?, logo_path=?, invoice_footer=?, updated_at=CURRENT_TIMESTAMP
                WHERE id = (SELECT id FROM settings ORDER BY id ASC LIMIT 1)
            """, (shop_name, tagline, phone, email, address, tax_number, currency_symbol, invoice_prefix, logo_data_uri, invoice_footer))
        elif remove_logo:
            cur.execute("""
                UPDATE settings 
                SET shop_name=?, tagline=?, phone=?, email=?, address=?, tax_number=?, currency_symbol=?, invoice_prefix=?, logo_path='', invoice_footer=?, updated_at=CURRENT_TIMESTAMP
                WHERE id = (SELECT id FROM settings ORDER BY id ASC LIMIT 1)
            """, (shop_name, tagline, phone, email, address, tax_number, currency_symbol, invoice_prefix, invoice_footer))
        else:
            cur.execute("""
                UPDATE settings 
                SET shop_name=?, tagline=?, phone=?, email=?, address=?, tax_number=?, currency_symbol=?, invoice_prefix=?, invoice_footer=?, updated_at=CURRENT_TIMESTAMP
                WHERE id = (SELECT id FROM settings ORDER BY id ASC LIMIT 1)
            """, (shop_name, tagline, phone, email, address, tax_number, currency_symbol, invoice_prefix, invoice_footer))

        conn.commit()
        conn.close()
        flash('Shop profile, logo, and invoice settings updated successfully!', 'success')
        return redirect(url_for('settings'))

    cur.execute("SELECT * FROM settings ORDER BY id ASC LIMIT 1")
    current_settings = dict(cur.fetchone())
    conn.close()
    return render_template('settings.html', s=current_settings)

# ----------------- FRIENDLY ERROR HANDLERS -----------------
@app.errorhandler(404)
def not_found(e):
    return render_template('base.html'), 404

@app.errorhandler(500)
def server_error(e):
    return "<h3>Internal Server Error</h3><p>Please check the logs or restart the application.</p>", 500

if __name__ == '__main__':
    local_ip = get_local_ip()
    print("=" * 65)
    print("  APEX BAGS & FOOTWEAR MULTI-USER INVENTORY SYSTEM")
    print("=" * 65)
    print(f"  Access on THIS computer:        http://127.0.0.1:5000")
    print(f"  Access from Phones/Tablets:     http://{local_ip}:5000")
    print("=" * 65)
    print("  Default Staff Logins:")
    print("    Admin / Owner:        admin    / admin123")
    print("    Floor Salesman:       sales1   / sales123")
    print("    Accounting / Cashier: accounts / accounts123")
    print("    Warehouse Manager:    godown   / godown123")
    print("=" * 65)
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
