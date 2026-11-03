import os
import io
import csv
import socket
from functools import wraps
from datetime import datetime, date
from flask import Flask, render_template, request, jsonify, redirect, url_for, send_file, flash, session
from werkzeug.security import check_password_hash, generate_password_hash
from database import get_connection, init_db, get_next_invoice_number
import excel_helpers

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
    pending_est = 0
    try:
        conn = get_connection()
        pending_est = conn.execute("SELECT COUNT(*) FROM estimates WHERE status = 'BOOKED'").fetchone()[0]
        conn.close()
    except Exception:
        pass

    return {
        'settings': get_settings(),
        'current_user': get_current_user(),
        'local_ip': get_local_ip(),
        'now': datetime.now(),
        'pending_estimates_count': pending_est
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
        hsn_code = request.form.get('hsn_code', '').strip()
        if not hsn_code:
            hsn_code = '4202' if category == 'Bags' else '6404'
        try:
            gst_rate = float(request.form.get('gst_rate', 18.0))
        except (ValueError, TypeError):
            gst_rate = 18.0

        # Auto-generate SKU if empty
        if not sku:
            prefix = "BAG" if category == "Bags" else "SH"
            cur.execute("SELECT COUNT(*) FROM products WHERE category = ?", (category,))
            count = cur.fetchone()[0] + 1
            sku = f"{prefix}-{count:04d}"

        cur.execute("""
            INSERT INTO products (sku, name, category, subcategory, brand, size, color, unit, 
                                  cost_price, wholesale_price, retail_price, stock_qty, min_alert_qty, location, notes, hsn_code, gst_rate)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (sku, name, category, subcategory, brand, size, color, unit,
              cost_price, wholesale_price, retail_price, stock_qty, min_alert_qty, location, notes, hsn_code, gst_rate))

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
        hsn_code = request.form.get('hsn_code', '').strip()
        if not hsn_code:
            hsn_code = '4202' if category == 'Bags' else '6404'
        try:
            gst_rate = float(request.form.get('gst_rate', 18.0))
        except (ValueError, TypeError):
            gst_rate = 18.0

        cur.execute("SELECT stock_qty FROM products WHERE id = ?", (product_id,))
        old_stock = cur.fetchone()[0]

        cur.execute("""
            UPDATE products 
            SET name=?, category=?, subcategory=?, brand=?, size=?, color=?, unit=?,
                cost_price=?, wholesale_price=?, retail_price=?, stock_qty=?, min_alert_qty=?, location=?, notes=?,
                hsn_code=?, gst_rate=?, updated_at=CURRENT_TIMESTAMP
            WHERE id=?
        """, (name, category, subcategory, brand, size, color, unit,
              cost_price, wholesale_price, retail_price, stock_qty, min_alert_qty, location, notes,
              hsn_code, gst_rate, product_id))

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

        # Determine Interstate (IGST) vs Intra-State (CGST + SGST)
        is_interstate = int(data.get('is_interstate', 0))
        cur.execute("SELECT state_code FROM settings LIMIT 1")
        s_row = cur.fetchone()
        shop_sc = (s_row['state_code'] if s_row and s_row['state_code'] else '27').strip()

        if party_id:
            cur.execute("SELECT state_code, gst_no FROM parties WHERE id = ?", (party_id,))
            p_row = cur.fetchone()
            if p_row:
                party_sc = (p_row['state_code'] or '').strip()
                if not party_sc and p_row['gst_no'] and len(p_row['gst_no'].strip()) >= 2:
                    party_sc = p_row['gst_no'].strip()[:2]
                if party_sc and party_sc != shop_sc:
                    is_interstate = 1

        # Enrich line items with catalog HSN and GST rate
        enriched_items = []
        for item in data['items']:
            prod_id = item['id']
            qty = int(item['quantity'])
            price = float(item['price'])

            cur.execute("SELECT stock_qty, cost_price, name, sku, size, color, unit, category, hsn_code, gst_rate FROM products WHERE id = ?", (prod_id,))
            prod = cur.fetchone()
            if not prod:
                raise Exception(f"Product ID {prod_id} not found!")

            cat = prod['category'] or 'Bags'
            default_hsn = '4202' if 'bag' in cat.lower() else '6404'
            hsn = str(prod['hsn_code'] or default_hsn).strip() or default_hsn
            gst_rate = float(prod['gst_rate'] if prod['gst_rate'] is not None else 18.0)

            enriched_items.append({
                'id': prod_id,
                'name': prod['name'],
                'sku': prod['sku'],
                'size': prod['size'],
                'color': prod['color'],
                'unit': prod['unit'] or ('Pieces' if cat == 'Bags' else 'Pairs'),
                'quantity': qty,
                'unit_price': price,
                'price': price,
                'cost_price': prod['cost_price'],
                'stock_qty': prod['stock_qty'],
                'hsn_code': hsn,
                'gst_rate': gst_rate
            })

        # Calculate GST Breakdown compliant with Indian GST laws
        gst_calc = excel_helpers.calculate_gst_breakdown(enriched_items, is_interstate=bool(is_interstate), prices_include_gst=True)

        subtotal = float(data.get('subtotal', gst_calc['grand_total']))
        discount = float(data.get('discount', 0))
        grand_total = float(data.get('grand_total', max(0.0, subtotal - discount)))
        paid_amount = float(data.get('paid_amount', grand_total))
        balance_amount = max(0.0, grand_total - paid_amount)

        taxable_amount = gst_calc['taxable_total']
        cgst_amount = gst_calc['cgst_total']
        sgst_amount = gst_calc['sgst_total']
        igst_amount = gst_calc['igst_total']
        total_tax = gst_calc['total_tax']

        # Adjust tax proportionally if bill discount was applied
        if discount > 0 and grand_total > 0 and subtotal > 0:
            ratio = grand_total / subtotal
            taxable_amount = round(taxable_amount * ratio, 2)
            cgst_amount = round(cgst_amount * ratio, 2)
            sgst_amount = round(sgst_amount * ratio, 2)
            igst_amount = round(igst_amount * ratio, 2)
            total_tax = round(cgst_amount + sgst_amount + igst_amount, 2)

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

        # Insert Sale record with full GST compliance breakdown
        cur.execute("""
            INSERT INTO sales (invoice_no, sale_type, party_id, customer_name, customer_phone, 
                               subtotal, discount, grand_total, payment_mode, payment_status, 
                               paid_amount, balance_amount, notes, salesman_id, salesman_name,
                               taxable_amount, cgst_amount, sgst_amount, igst_amount, total_tax, is_interstate)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (invoice_no, sale_type, party_id, customer_name, customer_phone,
              subtotal, discount, grand_total, payment_mode, payment_status,
              paid_amount, balance_amount, notes, salesman_id, salesman_name,
              taxable_amount, cgst_amount, sgst_amount, igst_amount, total_tax, is_interstate))
        sale_id = cur.lastrowid

        # Insert Sale Items with HSN & GST Breakdown & Deduct Physical Stock
        for item in gst_calc['items']:
            prod_id = item['id']
            qty = item['quantity']
            new_stock = item['stock_qty'] - qty

            # Deduct stock
            cur.execute("UPDATE products SET stock_qty = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (new_stock, prod_id))

            cur.execute("""
                INSERT INTO sale_items (sale_id, product_id, product_name, sku, size, color, unit, quantity, unit_price, cost_price, total_price,
                                        hsn_code, gst_rate, taxable_amount, cgst_amount, sgst_amount, igst_amount, tax_amount)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (sale_id, prod_id, item['name'], item['sku'], item['size'], item['color'], item['unit'],
                  qty, item['unit_price'], item['cost_price'], item['total_price'],
                  item['hsn_code'], item['gst_rate'], item['taxable_amount'],
                  item['cgst_amount'], item['sgst_amount'], item['igst_amount'], item['tax_amount']))

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

        cur.execute("SELECT state_code FROM settings LIMIT 1")
        s_row = cur.fetchone()
        shop_sc = (s_row['state_code'] if s_row and s_row['state_code'] else '27').strip()

        cur.execute("SELECT id, state_code, gst_no FROM parties WHERE name = ? AND party_type = 'SUPPLIER'", (supplier_name,))
        sup_row = cur.fetchone()
        is_interstate = int(data.get('is_interstate', 0))
        if sup_row:
            supplier_id = sup_row['id']
            sup_sc = (sup_row['state_code'] or '').strip()
            if not sup_sc and sup_row['gst_no'] and len(sup_row['gst_no'].strip()) >= 2:
                sup_sc = sup_row['gst_no'].strip()[:2]
            if sup_sc and sup_sc != shop_sc:
                is_interstate = 1
        else:
            cur.execute("INSERT INTO parties (name, party_type) VALUES (?, 'SUPPLIER')", (supplier_name,))
            supplier_id = cur.lastrowid

        # Enrich line items with catalog HSN and GST rate
        enriched_items = []
        for item in data['items']:
            prod_id = item['id']
            qty = int(item['quantity'])
            unit_cost = float(item['unit_cost'])

            cur.execute("SELECT stock_qty, name, sku, category, hsn_code, gst_rate, unit FROM products WHERE id = ?", (prod_id,))
            prod = cur.fetchone()
            if not prod:
                raise Exception(f"Product {prod_id} not found")

            cat = prod['category'] or 'Bags'
            default_hsn = '4202' if 'bag' in cat.lower() else '6404'
            hsn = str(prod['hsn_code'] or default_hsn).strip() or default_hsn
            gst_rate = float(prod['gst_rate'] if prod['gst_rate'] is not None else 18.0)

            enriched_items.append({
                'id': prod_id,
                'name': prod['name'],
                'sku': prod['sku'],
                'unit': prod['unit'] or ('Pieces' if cat == 'Bags' else 'Pairs'),
                'quantity': qty,
                'unit_cost': unit_cost,
                'price': unit_cost,
                'stock_qty': prod['stock_qty'],
                'hsn_code': hsn,
                'gst_rate': gst_rate
            })

        gst_calc = excel_helpers.calculate_gst_breakdown(enriched_items, is_interstate=bool(is_interstate), prices_include_gst=True)
        total_amount = gst_calc['grand_total']

        cur.execute("""
            INSERT INTO purchases (invoice_no, supplier_id, supplier_name, purchase_date, total_amount, paid_amount, payment_status, notes,
                                   taxable_amount, cgst_amount, sgst_amount, igst_amount, total_tax, is_interstate)
            VALUES (?, ?, ?, ?, ?, ?, 'PAID', ?, ?, ?, ?, ?, ?, ?)
        """, (invoice_no, supplier_id, supplier_name, purchase_date, total_amount, total_amount, notes,
              gst_calc['taxable_total'], gst_calc['cgst_total'], gst_calc['sgst_total'], gst_calc['igst_total'], gst_calc['total_tax'], is_interstate))
        purchase_id = cur.lastrowid

        for item in gst_calc['items']:
            prod_id = item['id']
            qty = item['quantity']
            unit_cost = item['unit_cost']
            tot_cost = item['total_price']
            new_stock = item['stock_qty'] + qty

            cur.execute("""
                UPDATE products 
                SET stock_qty = ?, cost_price = ?, updated_at = CURRENT_TIMESTAMP 
                WHERE id = ?
            """, (new_stock, unit_cost, prod_id))

            cur.execute("""
                INSERT INTO purchase_items (purchase_id, product_id, quantity, unit_cost, total_cost,
                                            hsn_code, gst_rate, taxable_amount, tax_amount)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (purchase_id, prod_id, qty, unit_cost, tot_cost,
                  item['hsn_code'], item['gst_rate'], item['taxable_amount'], item['tax_amount']))

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

@app.route('/export/template/bulk-purchase')
@login_required
@role_required(['ADMIN', 'WAREHOUSE'])
def download_bulk_purchase_template():
    file_format = request.args.get('format', 'xlsx').lower()
    file_bytes, filename, mimetype = excel_helpers.generate_bulk_purchase_template(file_format)
    return send_file(io.BytesIO(file_bytes), mimetype=mimetype, as_attachment=True, download_name=filename)

@app.route('/api/purchases/upload-excel', methods=['POST'])
@login_required
@role_required(['ADMIN', 'WAREHOUSE'])
def upload_bulk_purchases_excel():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'No file uploaded.'}), 400
    file = request.files['file']
    if not file or not file.filename:
        return jsonify({'success': False, 'message': 'Please select an Excel (.xlsx) or CSV file.'}), 400

    try:
        raw_rows = excel_helpers.parse_uploaded_file(file)
        if not raw_rows:
            return jsonify({'success': False, 'message': 'Uploaded file has no data rows.'}), 400
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error reading file: {str(e)}'}), 400

    conn = get_connection()
    cur = conn.cursor()
    try:
        grouped_purchases = {}
        for r in raw_rows:
            item_data = excel_helpers.extract_purchase_row_data(r)
            if not item_data['sku']:
                continue
            sup_name = item_data['supplier_name'] or 'Direct Factory Supplier'
            inv_no = item_data['invoice_no']
            group_key = (sup_name, inv_no, item_data['purchase_date'])
            if group_key not in grouped_purchases:
                grouped_purchases[group_key] = []
            grouped_purchases[group_key].append(item_data)

        if not grouped_purchases:
            conn.close()
            return jsonify({'success': False, 'message': 'No valid product rows with SKU found in file.'}), 400

        total_invoices_created = 0
        total_items_processed = 0
        total_purchase_val = 0.0

        for (sup_name, inv_no, p_date), items in grouped_purchases.items():
            cur.execute("SELECT id, state_code, gst_no FROM parties WHERE name = ? AND party_type = 'SUPPLIER'", (sup_name,))
            s_row = cur.fetchone()
            is_interstate = 0
            if s_row:
                supplier_id = s_row['id']
                if s_row['gst_no'] and len(s_row['gst_no'].strip()) >= 2 and s_row['gst_no'].strip()[:2] != '27':
                    is_interstate = 1
            else:
                cur.execute("INSERT INTO parties (name, party_type) VALUES (?, 'SUPPLIER')", (sup_name,))
                supplier_id = cur.lastrowid

            final_inv_no = inv_no
            if not final_inv_no:
                count = cur.execute("SELECT COUNT(*) FROM purchases").fetchone()[0] + 1
                final_inv_no = f"PUR-{datetime.now().strftime('%Y%m')}-{count:04d}"

            # Compute GST breakdown for this purchase batch
            gst_calc = excel_helpers.calculate_gst_breakdown(items, is_interstate=bool(is_interstate), prices_include_gst=True)
            shipment_total = gst_calc['grand_total']
            total_purchase_val += shipment_total

            cur.execute("SELECT id FROM purchases WHERE invoice_no = ?", (final_inv_no,))
            existing_pur = cur.fetchone()
            if existing_pur:
                purchase_id = existing_pur[0]
                cur.execute("""
                    UPDATE purchases 
                    SET total_amount = total_amount + ?, paid_amount = paid_amount + ?,
                        taxable_amount = taxable_amount + ?, cgst_amount = cgst_amount + ?,
                        sgst_amount = sgst_amount + ?, igst_amount = igst_amount + ?,
                        total_tax = total_tax + ?
                    WHERE id = ?
                """, (shipment_total, shipment_total, gst_calc['taxable_total'],
                      gst_calc['cgst_total'], gst_calc['sgst_total'], gst_calc['igst_total'],
                      gst_calc['total_tax'], purchase_id))
            else:
                cur.execute("""
                    INSERT INTO purchases (invoice_no, supplier_id, supplier_name, purchase_date, total_amount, paid_amount, payment_status, notes,
                                           taxable_amount, cgst_amount, sgst_amount, igst_amount, total_tax, is_interstate)
                    VALUES (?, ?, ?, ?, ?, ?, 'PAID', ?, ?, ?, ?, ?, ?, ?)
                """, (final_inv_no, supplier_id, sup_name, p_date, shipment_total, shipment_total, "Imported via Excel Bulk Upload",
                      gst_calc['taxable_total'], gst_calc['cgst_total'], gst_calc['sgst_total'], gst_calc['igst_total'], gst_calc['total_tax'], is_interstate))
                purchase_id = cur.lastrowid
                total_invoices_created += 1

            for it in gst_calc['items']:
                sku = it['sku']
                qty = it['quantity']
                cost = it['cost_price']
                item_cost_total = it['total_price']

                cur.execute("SELECT id, stock_qty FROM products WHERE sku = ?", (sku,))
                prod_row = cur.fetchone()
                if prod_row:
                    prod_id = prod_row['id']
                    new_stock = prod_row['stock_qty'] + qty
                    cur.execute("""
                        UPDATE products 
                        SET stock_qty = ?, cost_price = ?, wholesale_price = COALESCE(NULLIF(?, 0), wholesale_price),
                            retail_price = COALESCE(NULLIF(?, 0), retail_price),
                            hsn_code = COALESCE(NULLIF(?, ''), hsn_code),
                            gst_rate = COALESCE(NULLIF(?, 0), gst_rate),
                            updated_at = CURRENT_TIMESTAMP
                        WHERE id = ?
                    """, (new_stock, cost, it['wholesale_price'], it['retail_price'], it['hsn_code'], it['gst_rate'], prod_id))
                else:
                    cur.execute("""
                        INSERT INTO products (sku, name, category, subcategory, brand, size, color, unit, cost_price, wholesale_price, retail_price, stock_qty, min_alert_qty, location, notes, hsn_code, gst_rate)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 5, ?, ?, ?, ?)
                    """, (sku, it['name'], it['category'], it['subcategory'], it['brand'], it['size'], it['color'], it['unit'], cost, it['wholesale_price'], it['retail_price'], qty, it['location'], it['notes'], it['hsn_code'], it['gst_rate']))
                    prod_id = cur.lastrowid
                    new_stock = qty

                cur.execute("""
                    INSERT INTO purchase_items (purchase_id, product_id, quantity, unit_cost, total_cost,
                                                hsn_code, gst_rate, taxable_amount, tax_amount)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (purchase_id, prod_id, qty, cost, item_cost_total,
                      it['hsn_code'], it['gst_rate'], it['taxable_amount'], it['tax_amount']))

                cur.execute("""
                    INSERT INTO stock_logs (product_id, change_type, quantity_change, stock_after, reference_no, reason)
                    VALUES (?, 'PURCHASE_IN', ?, ?, ?, ?)
                """, (prod_id, qty, new_stock, final_inv_no, f"Bulk Excel Purchase from {sup_name}"))
                total_items_processed += 1

        conn.commit()
        conn.close()
        return jsonify({
            'success': True,
            'message': f'Bulk Purchase Upload Successful! Processed {total_items_processed} items across {total_invoices_created} supplier invoices. Total: ₹{total_purchase_val:,.2f}'
        })
    except Exception as e:
        conn.rollback()
        conn.close()
        return jsonify({'success': False, 'message': f'Error during bulk purchase upload: {str(e)}'}), 500

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
    raw_items = [dict(r) for r in cur.fetchall()]

    party = None
    if sale['party_id']:
        cur.execute("SELECT * FROM parties WHERE id = ?", (sale['party_id'],))
        party_row = cur.fetchone()
        if party_row:
            party = dict(party_row)

    conn.close()

    # Build HSN Summary for Indian GST Rule 46 invoice compliance
    hsn_dict = {}
    items = []
    is_interstate = bool(sale['is_interstate'])

    for it in raw_items:
        hsn = str(it.get('hsn_code') or '').strip()
        if not hsn:
            cat = ('bag' if ('bag' in (it.get('product_name') or '').lower() or 'wallet' in (it.get('product_name') or '').lower()) else 'shoe')
            hsn = '4202' if cat == 'bag' else '6404'
        it['hsn_code'] = hsn

        rate = float(it.get('gst_rate') or 18.0)
        it['gst_rate'] = rate

        tot = float(it.get('total_price') or 0.0)
        taxable = float(it.get('taxable_amount') or 0.0)
        if taxable <= 0 and tot > 0:
            taxable = round(tot / (1.0 + (rate / 100.0)), 2)
            tax = round(tot - taxable, 2)
            it['taxable_amount'] = taxable
            it['tax_amount'] = tax
            if is_interstate:
                it['igst_amount'] = tax
                it['cgst_amount'] = 0.0
                it['sgst_amount'] = 0.0
            else:
                it['cgst_amount'] = round(tax / 2.0, 2)
                it['sgst_amount'] = round(tax - it['cgst_amount'], 2)
                it['igst_amount'] = 0.0

        items.append(it)

        # Aggregate HSN
        key = (hsn, rate)
        if key not in hsn_dict:
            desc = "Trunks, suitcases, bags, handbags, wallets" if hsn == '4202' else "Footwear, shoes, sneakers, sandals"
            hsn_dict[key] = {
                'hsn_code': hsn,
                'description': desc,
                'uqc': it.get('unit') or ('Pieces' if hsn == '4202' else 'Pairs'),
                'total_qty': 0,
                'gst_rate': rate,
                'taxable_amount': 0.0,
                'cgst_amount': 0.0,
                'sgst_amount': 0.0,
                'igst_amount': 0.0,
                'total_tax': 0.0,
                'total_amount': 0.0
            }
        hsn_dict[key]['total_qty'] += float(it.get('quantity') or 0)
        hsn_dict[key]['taxable_amount'] += float(it.get('taxable_amount') or 0.0)
        hsn_dict[key]['cgst_amount'] += float(it.get('cgst_amount') or 0.0)
        hsn_dict[key]['sgst_amount'] += float(it.get('sgst_amount') or 0.0)
        hsn_dict[key]['igst_amount'] += float(it.get('igst_amount') or 0.0)
        hsn_dict[key]['total_tax'] += float(it.get('tax_amount') or (float(it.get('cgst_amount') or 0) + float(it.get('sgst_amount') or 0) + float(it.get('igst_amount') or 0)))
        hsn_dict[key]['total_amount'] += tot

    hsn_summary = list(hsn_dict.values())
    for h in hsn_summary:
        h['taxable_amount'] = round(h['taxable_amount'], 2)
        h['cgst_amount'] = round(h['cgst_amount'], 2)
        h['sgst_amount'] = round(h['sgst_amount'], 2)
        h['igst_amount'] = round(h['igst_amount'], 2)
        h['total_tax'] = round(h['total_tax'], 2)
        h['total_amount'] = round(h['total_amount'], 2)

    return render_template('invoice_print.html', sale=dict(sale), items=items, party=party, hsn_summary=hsn_summary)

@app.route('/export/template/bulk-sale')
@login_required
@role_required(['ADMIN', 'SALESMAN', 'ACCOUNTING'])
def download_bulk_sale_template():
    file_format = request.args.get('format', 'xlsx').lower()
    file_bytes, filename, mimetype = excel_helpers.generate_bulk_sale_template(file_format)
    return send_file(io.BytesIO(file_bytes), mimetype=mimetype, as_attachment=True, download_name=filename)

@app.route('/api/sales/upload-excel', methods=['POST'])
@login_required
@role_required(['ADMIN', 'SALESMAN', 'ACCOUNTING'])
def upload_bulk_sales_excel():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'No file uploaded.'}), 400
    file = request.files['file']
    if not file or not file.filename:
        return jsonify({'success': False, 'message': 'Please select an Excel (.xlsx) or CSV file.'}), 400

    try:
        raw_rows = excel_helpers.parse_uploaded_file(file)
        if not raw_rows:
            return jsonify({'success': False, 'message': 'Uploaded file contains no data rows.'}), 400
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error reading file: {str(e)}'}), 400

    conn = get_connection()
    cur = conn.cursor()
    try:
        grouped_sales = {}
        for idx, r in enumerate(raw_rows):
            item_data = excel_helpers.extract_sale_row_data(r)
            if not item_data['sku']:
                continue

            cur.execute("SELECT id, name, sku, size, color, unit, stock_qty, wholesale_price, retail_price, cost_price FROM products WHERE sku = ?", (item_data['sku'],))
            prod = cur.fetchone()
            if not prod:
                conn.close()
                return jsonify({'success': False, 'message': f'Row {idx + 2}: SKU "{item_data["sku"]}" does not exist in inventory catalog! Please add it first or check spelling.'}), 400

            item_data['product'] = dict(prod)
            if item_data['unit_price'] <= 0:
                item_data['unit_price'] = prod['wholesale_price'] if item_data['sale_type'] == 'WHOLESALE' else prod['retail_price']

            ref = item_data['order_ref'] or f"ROW_{idx}"
            grp_key = (item_data['customer_name'], item_data['customer_phone'], item_data['sale_type'], item_data['payment_mode'], ref)
            if grp_key not in grouped_sales:
                grouped_sales[grp_key] = []
            grouped_sales[grp_key].append(item_data)

        if not grouped_sales:
            conn.close()
            return jsonify({'success': False, 'message': 'No valid sale items found in file.'}), 400

        # Check stock availability for all items
        for grp_key, items in grouped_sales.items():
            for it in items:
                prod = it['product']
                if it['quantity'] > prod['stock_qty']:
                    conn.close()
                    return jsonify({
                        'success': False,
                        'message': f'Insufficient stock for SKU {prod["sku"]} ({prod["name"]})! Requested: {it["quantity"]}, Available: {prod["stock_qty"]}.'
                    }), 400

        settings_data = get_settings()
        prefix = settings_data.get('invoice_prefix', 'SSIPL').strip() or 'SSIPL'
        salesman_id = session.get('user_id')
        salesman_name = session.get('full_name') or session.get('username') or 'Sales Staff'

        total_invoices_created = 0
        total_sales_value = 0.0

        for (cust_name, cust_phone, sale_type, pay_mode, ref), items in grouped_sales.items():
            party_id = None
            is_interstate = 0
            if sale_type == 'WHOLESALE' or cust_name != 'Walk-in Customer':
                cur.execute("SELECT id, state_code, gst_no FROM parties WHERE name = ? AND party_type = 'WHOLESALE_BUYER'", (cust_name,))
                p_row = cur.fetchone()
                if p_row:
                    party_id = p_row['id']
                    if p_row['gst_no'] and len(p_row['gst_no'].strip()) >= 2 and p_row['gst_no'].strip()[:2] != '27':
                        is_interstate = 1
                else:
                    cur.execute("INSERT INTO parties (name, phone, party_type, current_balance) VALUES (?, ?, 'WHOLESALE_BUYER', 0.0)", (cust_name, cust_phone))
                    party_id = cur.lastrowid

            # Enrich items with proper HSN and GST rates
            calc_input_items = []
            for it in items:
                prod = it['product']
                cat = prod.get('category', 'Bags')
                hsn = it.get('hsn_code') or prod.get('hsn_code') or ('4202' if 'bag' in cat.lower() else '6404')
                gst_rate = it.get('gst_rate') or prod.get('gst_rate') or 18.0
                calc_input_items.append({
                    'id': prod['id'],
                    'name': prod['name'],
                    'sku': prod['sku'],
                    'size': prod['size'],
                    'color': prod['color'],
                    'unit': prod['unit'],
                    'quantity': it['quantity'],
                    'price': it['unit_price'],
                    'unit_price': it['unit_price'],
                    'cost_price': prod['cost_price'],
                    'stock_qty': prod['stock_qty'],
                    'hsn_code': str(hsn).strip(),
                    'gst_rate': float(gst_rate)
                })

            gst_calc = excel_helpers.calculate_gst_breakdown(calc_input_items, is_interstate=bool(is_interstate), prices_include_gst=True)
            subtotal = gst_calc['grand_total']
            total_discount = sum(it.get('discount', 0) for it in items)
            grand_total = max(0.0, subtotal - total_discount)
            total_sales_value += grand_total

            taxable_amount = gst_calc['taxable_total']
            cgst_amount = gst_calc['cgst_total']
            sgst_amount = gst_calc['sgst_total']
            igst_amount = gst_calc['igst_total']
            total_tax = gst_calc['total_tax']

            if total_discount > 0 and grand_total > 0 and subtotal > 0:
                ratio = grand_total / subtotal
                taxable_amount = round(taxable_amount * ratio, 2)
                cgst_amount = round(cgst_amount * ratio, 2)
                sgst_amount = round(sgst_amount * ratio, 2)
                igst_amount = round(igst_amount * ratio, 2)
                total_tax = round(cgst_amount + sgst_amount + igst_amount, 2)

            specified_paid = [it['paid_amount'] for it in items if it['paid_amount'] is not None]
            if specified_paid:
                paid_amount = sum(specified_paid)
            else:
                paid_amount = 0.0 if pay_mode == 'CREDIT' else grand_total

            balance_amount = max(0.0, grand_total - paid_amount)
            payment_status = 'PAID' if balance_amount == 0 else ('PARTIAL' if paid_amount > 0 else 'CREDIT')

            invoice_no = get_next_invoice_number(conn, prefix=prefix)
            notes = f"Bulk Excel Import (Ref: {ref})"

            cur.execute("""
                INSERT INTO sales (invoice_no, sale_type, party_id, customer_name, customer_phone, 
                                   subtotal, discount, grand_total, payment_mode, payment_status, 
                                   paid_amount, balance_amount, notes, salesman_id, salesman_name,
                                   taxable_amount, cgst_amount, sgst_amount, igst_amount, total_tax, is_interstate)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (invoice_no, sale_type, party_id, cust_name, cust_phone, subtotal, total_discount, grand_total,
                  pay_mode, payment_status, paid_amount, balance_amount, notes, salesman_id, salesman_name,
                  taxable_amount, cgst_amount, sgst_amount, igst_amount, total_tax, is_interstate))
            sale_id = cur.lastrowid
            total_invoices_created += 1

            for it in gst_calc['items']:
                prod_id = it['id']
                qty = it['quantity']
                price = it['unit_price']
                cost = it['cost_price']
                item_total = it['total_price']

                cur.execute("SELECT stock_qty FROM products WHERE id = ?", (prod_id,))
                cur_prod = cur.fetchone()
                new_stock = cur_prod['stock_qty'] - qty
                cur.execute("UPDATE products SET stock_qty = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (new_stock, prod_id))

                cur.execute("""
                    INSERT INTO sale_items (sale_id, product_id, product_name, sku, size, color, unit, quantity, unit_price, cost_price, total_price,
                                            hsn_code, gst_rate, taxable_amount, cgst_amount, sgst_amount, igst_amount, tax_amount)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (sale_id, prod_id, it['name'], it['sku'], it['size'], it['color'], it['unit'],
                      qty, price, cost, item_total,
                      it['hsn_code'], it['gst_rate'], it['taxable_amount'],
                      it['cgst_amount'], it['sgst_amount'], it['igst_amount'], it['tax_amount']))

                cur.execute("""
                    INSERT INTO stock_logs (product_id, change_type, quantity_change, stock_after, reference_no, reason)
                    VALUES (?, 'SALE_OUT', ?, ?, ?, ?)
                """, (prod_id, -qty, new_stock, invoice_no, f"Bulk Excel Sale to {cust_name}"))

            if party_id and balance_amount > 0:
                cur.execute("UPDATE parties SET current_balance = current_balance + ? WHERE id = ?", (balance_amount, party_id))

        conn.commit()
        conn.close()
        return jsonify({
            'success': True,
            'message': f'Bulk Sales Upload Successful! Created {total_invoices_created} invoices for total ₹{total_sales_value:,.2f}. Stock deducted.'
        })
    except Exception as e:
        conn.rollback()
        conn.close()
        return jsonify({'success': False, 'message': f'Error during bulk sales upload: {str(e)}'}), 500

# ----------------- ESTIMATES / SALES ORDERS (ADVANCE BOOKING & LATER DISPATCH) -----------------
@app.route('/estimates')
@login_required
@role_required(['ADMIN', 'SALESMAN', 'ACCOUNTING'])
def estimates():
    status_filter = request.args.get('status', '').strip().upper()
    search = request.args.get('search', '').strip()

    conn = get_connection()
    cur = conn.cursor()

    query = "SELECT * FROM estimates WHERE 1=1"
    params = []
    if status_filter:
        query += " AND status = ?"
        params.append(status_filter)
    if search:
        query += " AND (estimate_no LIKE ? OR customer_name LIKE ? OR customer_phone LIKE ?)"
        term = f"%{search}%"
        params.extend([term, term, term])

    query += " ORDER BY id DESC"
    cur.execute(query, params)
    estimates_list = [dict(r) for r in cur.fetchall()]

    # Summary Stats
    cur.execute("SELECT COUNT(*), COALESCE(SUM(advance_amount), 0), COALESCE(SUM(balance_amount), 0) FROM estimates WHERE status = 'BOOKED'")
    stats_row = cur.fetchone()
    active_count = stats_row[0] or 0
    total_advance = stats_row[1] or 0.0
    pending_balance = stats_row[2] or 0.0

    cur.execute("SELECT * FROM products WHERE stock_qty > 0 ORDER BY name ASC")
    products = [dict(r) for r in cur.fetchall()]

    cur.execute("SELECT * FROM parties WHERE party_type = 'WHOLESALE_BUYER' ORDER BY name ASC")
    parties = [dict(r) for r in cur.fetchall()]

    conn.close()
    return render_template('estimates.html',
                           estimates=estimates_list,
                           active_count=active_count,
                           total_advance=total_advance,
                           pending_balance=pending_balance,
                           products=products,
                           parties=parties,
                           status_filter=status_filter,
                           search=search)

@app.route('/api/estimates/create', methods=['POST'])
@login_required
@role_required(['ADMIN', 'SALESMAN', 'ACCOUNTING'])
def create_estimate():
    data = request.get_json()
    if not data or not data.get('items'):
        return jsonify({'success': False, 'message': 'No items specified in booking order'}), 400

    conn = get_connection()
    cur = conn.cursor()
    try:
        customer_name = data.get('customer_name', '').strip() or 'Walk-in Customer'
        customer_phone = data.get('customer_phone', '').strip()
        party_id = data.get('party_id') or None
        order_type = data.get('order_type', 'WHOLESALE').upper()
        # Determine interstate for estimate
        is_interstate = int(data.get('is_interstate', 0))
        cur.execute("SELECT state_code FROM settings LIMIT 1")
        s_row = cur.fetchone()
        shop_sc = (s_row['state_code'] if s_row and s_row['state_code'] else '27').strip()
        if party_id:
            cur.execute("SELECT state_code, gst_no FROM parties WHERE id = ?", (party_id,))
            p_row = cur.fetchone()
            if p_row:
                p_sc = (p_row['state_code'] or '').strip()
                if not p_sc and p_row['gst_no'] and len(p_row['gst_no'].strip()) >= 2:
                    p_sc = p_row['gst_no'].strip()[:2]
                if p_sc and p_sc != shop_sc:
                    is_interstate = 1

        enriched_items = []
        for item in data['items']:
            prod_id = item['id']
            qty = int(item['quantity'])
            price = float(item['price'])
            cur.execute("SELECT name, sku, size, color, unit, category, hsn_code, gst_rate FROM products WHERE id = ?", (prod_id,))
            prod = cur.fetchone()
            if not prod:
                raise Exception(f"Product ID {prod_id} not found!")

            cat = prod['category'] or 'Bags'
            hsn = str(prod['hsn_code'] or ('4202' if 'bag' in cat.lower() else '6404')).strip()
            gst_rate = float(prod['gst_rate'] if prod['gst_rate'] is not None else 18.0)

            enriched_items.append({
                'id': prod_id,
                'name': prod['name'],
                'sku': prod['sku'],
                'size': prod['size'],
                'color': prod['color'],
                'unit': prod['unit'] or ('Pieces' if cat == 'Bags' else 'Pairs'),
                'quantity': qty,
                'unit_price': price,
                'price': price,
                'hsn_code': hsn,
                'gst_rate': gst_rate
            })

        gst_calc = excel_helpers.calculate_gst_breakdown(enriched_items, is_interstate=bool(is_interstate), prices_include_gst=True)
        subtotal = float(data.get('subtotal', gst_calc['grand_total']))
        discount = float(data.get('discount', 0.0))
        grand_total = float(data.get('grand_total', max(0.0, subtotal - discount)))
        advance_amount = float(data.get('advance_amount', 0.0))
        advance_payment_mode = data.get('advance_payment_mode', 'CASH').upper()
        expected_dispatch_date = data.get('expected_dispatch_date') or None
        notes = data.get('notes', '').strip()

        taxable_amount = gst_calc['taxable_total']
        cgst_amount = gst_calc['cgst_total']
        sgst_amount = gst_calc['sgst_total']
        igst_amount = gst_calc['igst_total']
        total_tax = gst_calc['total_tax']

        if discount > 0 and grand_total > 0 and subtotal > 0:
            ratio = grand_total / subtotal
            taxable_amount = round(taxable_amount * ratio, 2)
            cgst_amount = round(cgst_amount * ratio, 2)
            sgst_amount = round(sgst_amount * ratio, 2)
            igst_amount = round(igst_amount * ratio, 2)
            total_tax = round(cgst_amount + sgst_amount + igst_amount, 2)

        balance_amount = max(0.0, grand_total - advance_amount)
        salesman_id = session.get('user_id')
        salesman_name = session.get('full_name') or session.get('username') or 'Staff'

        estimate_no = get_next_invoice_number(conn, prefix='EST')

        cur.execute("""
            INSERT INTO estimates (estimate_no, party_id, customer_name, customer_phone, order_type,
                                   subtotal, discount, grand_total, advance_amount, advance_payment_mode,
                                   balance_amount, status, expected_dispatch_date, salesman_id, salesman_name, notes,
                                   taxable_amount, cgst_amount, sgst_amount, igst_amount, total_tax, is_interstate)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'BOOKED', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (estimate_no, party_id, customer_name, customer_phone, order_type,
              subtotal, discount, grand_total, advance_amount, advance_payment_mode,
              balance_amount, expected_dispatch_date, salesman_id, salesman_name, notes,
              taxable_amount, cgst_amount, sgst_amount, igst_amount, total_tax, is_interstate))
        estimate_id = cur.lastrowid

        for it in gst_calc['items']:
            cur.execute("""
                INSERT INTO estimate_items (estimate_id, product_id, product_name, sku, size, color, unit, quantity, unit_price, total_price,
                                            hsn_code, gst_rate, taxable_amount, tax_amount)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (estimate_id, it['id'], it['name'], it['sku'], it['size'], it['color'], it['unit'],
                  it['quantity'], it['unit_price'], it['total_price'],
                  it['hsn_code'], it['gst_rate'], it['taxable_amount'], it['tax_amount']))

        conn.commit()
        return jsonify({
            'success': True,
            'estimate_id': estimate_id,
            'estimate_no': estimate_no,
            'message': f'Advance Booking {estimate_no} created successfully!'
        })
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 500
    finally:
        conn.close()

@app.route('/estimate/<int:estimate_id>/print')
@login_required
def print_estimate(estimate_id):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM estimates WHERE id = ?", (estimate_id,))
    est = cur.fetchone()
    if not est:
        conn.close()
        return "Estimate / Booking Order not found", 404

    cur.execute("SELECT * FROM estimate_items WHERE estimate_id = ?", (estimate_id,))
    items = [dict(r) for r in cur.fetchall()]

    party = None
    if est['party_id']:
        cur.execute("SELECT * FROM parties WHERE id = ?", (est['party_id'],))
        p_row = cur.fetchone()
        if p_row:
            party = dict(p_row)

    conn.close()
    return render_template('estimate_print.html', estimate=dict(est), items=items, party=party)

@app.route('/api/estimates/<int:estimate_id>/dispatch', methods=['POST'])
@login_required
@role_required(['ADMIN', 'SALESMAN', 'ACCOUNTING'])
def dispatch_estimate(estimate_id):
    data = request.get_json() or {}
    additional_paid = float(data.get('additional_paid_amount', 0.0))
    dispatch_payment_mode = data.get('dispatch_payment_mode', 'CASH').upper()

    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute("SELECT * FROM estimates WHERE id = ?", (estimate_id,))
        est = cur.fetchone()
        if not est:
            conn.close()
            return jsonify({'success': False, 'message': 'Estimate not found'}), 404

        if est['status'] != 'BOOKED':
            conn.close()
            return jsonify({'success': False, 'message': f'Order is already {est["status"]}. Cannot dispatch.'}), 400

        cur.execute("SELECT * FROM estimate_items WHERE estimate_id = ?", (estimate_id,))
        items = [dict(r) for r in cur.fetchall()]

        # Verify physical stock
        for it in items:
            cur.execute("SELECT stock_qty, name, sku FROM products WHERE id = ?", (it['product_id'],))
            p = cur.fetchone()
            if not p:
                conn.close()
                return jsonify({'success': False, 'message': f'Product {it["product_name"]} no longer exists!'}), 400
            if p['stock_qty'] < it['quantity']:
                conn.close()
                return jsonify({'success': False, 'message': f'Not enough stock to dispatch {p["name"]}! Required: {it["quantity"]}, Available: {p["stock_qty"]}.'}), 400

        settings_data = get_settings()
        prefix = settings_data.get('invoice_prefix', 'SSIPL').strip() or 'SSIPL'
        invoice_no = get_next_invoice_number(conn, prefix=prefix)

        advance_amount = float(est['advance_amount'])
        grand_total = float(est['grand_total'])
        total_paid = min(grand_total, advance_amount + additional_paid)
        final_balance = max(0.0, grand_total - total_paid)
        payment_status = 'PAID' if final_balance == 0 else ('PARTIAL' if total_paid > 0 else 'CREDIT')

        salesman_id = session.get('user_id')
        salesman_name = session.get('full_name') or session.get('username') or est['salesman_name']
        notes = f"Dispatched against Advance Booking {est['estimate_no']}"
        if est['notes']:
            notes += f" | {est['notes']}"

        # Insert Sale with GST fields
        is_interstate = bool(est['is_interstate']) if 'is_interstate' in est.keys() else False
        taxable_amount = est['taxable_amount'] if 'taxable_amount' in est.keys() else 0.0
        cgst_amount = est['cgst_amount'] if 'cgst_amount' in est.keys() else 0.0
        sgst_amount = est['sgst_amount'] if 'sgst_amount' in est.keys() else 0.0
        igst_amount = est['igst_amount'] if 'igst_amount' in est.keys() else 0.0
        total_tax = est['total_tax'] if 'total_tax' in est.keys() else 0.0

        cur.execute("""
            INSERT INTO sales (invoice_no, sale_type, party_id, customer_name, customer_phone, 
                               subtotal, discount, grand_total, payment_mode, payment_status, 
                               paid_amount, balance_amount, notes, salesman_id, salesman_name,
                               estimate_id, advance_adjusted,
                               taxable_amount, cgst_amount, sgst_amount, igst_amount, total_tax, is_interstate)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (invoice_no, est['order_type'], est['party_id'], est['customer_name'], est['customer_phone'],
              est['subtotal'], est['discount'], grand_total, dispatch_payment_mode, payment_status,
              total_paid, final_balance, notes, salesman_id, salesman_name,
              estimate_id, advance_amount,
              taxable_amount, cgst_amount, sgst_amount, igst_amount, total_tax, 1 if is_interstate else 0))
        sale_id = cur.lastrowid

        # Deduct physical stock and insert sale items with GST fields
        for it in items:
            prod_id = it['product_id']
            qty = it['quantity']
            price = it['unit_price']
            tot_price = it['total_price']
            hsn = it['hsn_code'] if 'hsn_code' in it.keys() and it['hsn_code'] else ('4202' if 'bag' in it['product_name'].lower() else '6404')
            gst_rate = float(it['gst_rate']) if 'gst_rate' in it.keys() and it['gst_rate'] is not None else 18.0
            
            taxable_val = float(it['taxable_amount']) if 'taxable_amount' in it.keys() and it['taxable_amount'] else round(tot_price / (1.0 + (gst_rate / 100.0)), 2)
            tax_val = float(it['tax_amount']) if 'tax_amount' in it.keys() and it['tax_amount'] else round(tot_price - taxable_val, 2)
            if is_interstate:
                cgst = 0.0
                sgst = 0.0
                igst = tax_val
            else:
                cgst = round(tax_val / 2.0, 2)
                sgst = round(tax_val - cgst, 2)
                igst = 0.0

            cur.execute("SELECT stock_qty, cost_price FROM products WHERE id = ?", (prod_id,))
            prod = cur.fetchone()
            cost_price = prod['cost_price']
            new_stock = prod['stock_qty'] - qty

            cur.execute("UPDATE products SET stock_qty = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (new_stock, prod_id))

            cur.execute("""
                INSERT INTO sale_items (sale_id, product_id, product_name, sku, size, color, unit, quantity, unit_price, cost_price, total_price,
                                        hsn_code, gst_rate, taxable_amount, cgst_amount, sgst_amount, igst_amount, tax_amount)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (sale_id, prod_id, it['product_name'], it['sku'], it['size'], it['color'], it['unit'],
                  qty, price, cost_price, tot_price,
                  hsn, gst_rate, taxable_val, cgst, sgst, igst, tax_val))

            cur.execute("""
                INSERT INTO stock_logs (product_id, change_type, quantity_change, stock_after, reference_no, reason)
                VALUES (?, 'SALE_OUT', ?, ?, ?, ?)
            """, (prod_id, -qty, new_stock, invoice_no, f"Dispatched from Order {est['estimate_no']}"))

        # Wholesale ledger balance update
        if est['party_id'] and final_balance > 0:
            cur.execute("UPDATE parties SET current_balance = current_balance + ? WHERE id = ?", (final_balance, est['party_id']))

        # Update estimate status
        cur.execute("""
            UPDATE estimates 
            SET status = 'DISPATCHED', converted_invoice_id = ?, converted_invoice_no = ?, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        """, (sale_id, invoice_no, estimate_id))

        conn.commit()
        return jsonify({
            'success': True,
            'invoice_id': sale_id,
            'invoice_no': invoice_no,
            'message': f"Order {est['estimate_no']} successfully dispatched! Final Tax Invoice #{invoice_no} generated."
        })
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 500
    finally:
        conn.close()

@app.route('/api/estimates/<int:estimate_id>/cancel', methods=['POST'])
@login_required
@role_required(['ADMIN', 'SALESMAN', 'ACCOUNTING'])
def cancel_estimate(estimate_id):
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute("SELECT status FROM estimates WHERE id = ?", (estimate_id,))
        row = cur.fetchone()
        if not row:
            return jsonify({'success': False, 'message': 'Estimate not found'}), 404
        if row['status'] != 'BOOKED':
            return jsonify({'success': False, 'message': f'Cannot cancel order with status {row["status"]}'}), 400

        cur.execute("UPDATE estimates SET status = 'CANCELLED', updated_at = CURRENT_TIMESTAMP WHERE id = ?", (estimate_id,))
        conn.commit()
        return jsonify({'success': True, 'message': 'Advance Booking / Estimate cancelled.'})
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 500
    finally:
        conn.close()

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
        state_code = request.form.get('state_code', '27').strip() or '27'
        state_name = request.form.get('state_name', 'Maharashtra').strip() or 'Maharashtra'
        remove_logo = request.form.get('remove_logo') == '1'

        logo_file = request.files.get('logo')
        if logo_file and logo_file.filename:
            raw_bytes = logo_file.read()
            mimetype = logo_file.content_type or 'image/png'
            b64_str = base64.b64encode(raw_bytes).decode('utf-8')
            logo_data_uri = f"data:{mimetype};base64,{b64_str}"

            cur.execute("""
                UPDATE settings 
                SET shop_name=?, tagline=?, phone=?, email=?, address=?, tax_number=?, currency_symbol=?, invoice_prefix=?, logo_path=?, invoice_footer=?, state_code=?, state_name=?, updated_at=CURRENT_TIMESTAMP
                WHERE id = (SELECT id FROM settings ORDER BY id ASC LIMIT 1)
            """, (shop_name, tagline, phone, email, address, tax_number, currency_symbol, invoice_prefix, logo_data_uri, invoice_footer, state_code, state_name))
        elif remove_logo:
            cur.execute("""
                UPDATE settings 
                SET shop_name=?, tagline=?, phone=?, email=?, address=?, tax_number=?, currency_symbol=?, invoice_prefix=?, logo_path='', invoice_footer=?, state_code=?, state_name=?, updated_at=CURRENT_TIMESTAMP
                WHERE id = (SELECT id FROM settings ORDER BY id ASC LIMIT 1)
            """, (shop_name, tagline, phone, email, address, tax_number, currency_symbol, invoice_prefix, invoice_footer, state_code, state_name))
        else:
            cur.execute("""
                UPDATE settings 
                SET shop_name=?, tagline=?, phone=?, email=?, address=?, tax_number=?, currency_symbol=?, invoice_prefix=?, invoice_footer=?, state_code=?, state_name=?, updated_at=CURRENT_TIMESTAMP
                WHERE id = (SELECT id FROM settings ORDER BY id ASC LIMIT 1)
            """, (shop_name, tagline, phone, email, address, tax_number, currency_symbol, invoice_prefix, invoice_footer, state_code, state_name))

        conn.commit()
        conn.close()
        flash('Shop profile, logo, state code, and invoice settings updated successfully!', 'success')
        return redirect(url_for('settings'))

    cur.execute("SELECT * FROM settings ORDER BY id ASC LIMIT 1")
    current_settings = dict(cur.fetchone())
    conn.close()
    return render_template('settings.html', s=current_settings)

# ----------------- INDIAN GST REPORTS & RETURNS (GSTR-1 & GSTR-3B) -----------------
@app.route('/gst-reports')
@login_required
@role_required(['ADMIN', 'ACCOUNTING'])
def gst_reports():
    start_date = request.args.get('start_date', '')
    end_date = request.args.get('end_date', '')

    # Default to current month
    if not start_date and not end_date:
        today = date.today()
        start_date = today.replace(day=1).isoformat()
        end_date = today.isoformat()

    conn = get_connection()
    summary = excel_helpers.generate_gstr3b_summary_data(conn, start_date=start_date if start_date else None, end_date=end_date if end_date else None)
    settings_data = get_settings()
    conn.close()

    return render_template('gst_reports.html', summary=summary, start_date=start_date, end_date=end_date, settings=settings_data)


@app.route('/export/gstr1')
@login_required
@role_required(['ADMIN', 'ACCOUNTING'])
def export_gstr1():
    start_date = request.args.get('start_date', '') or None
    end_date = request.args.get('end_date', '') or None

    conn = get_connection()
    try:
        file_bytes, filename, mimetype = excel_helpers.generate_gstr1_excel_report(conn, start_date=start_date, end_date=end_date)
    finally:
        conn.close()

    return send_file(io.BytesIO(file_bytes), mimetype=mimetype, as_attachment=True, download_name=filename)

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
