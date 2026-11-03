"""
Database module for Bags & Shoes Wholesale/Retail Inventory Management.
Uses SQLite for robust, zero-configuration, single-file local storage.
"""

import sqlite3
import os
from datetime import datetime

from werkzeug.security import generate_password_hash, check_password_hash

DB_PATH = os.path.join(os.path.dirname(__file__), 'inventory.db')

def get_connection():
    conn = sqlite3.connect(DB_PATH, timeout=10.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    # WAL Mode allows multiple readers and writers simultaneously without database locks
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA busy_timeout = 10000;")
    conn.execute("PRAGMA synchronous = NORMAL;")
    return conn

def init_db():
    conn = get_connection()
    cursor = conn.cursor()

    # 0. Users Table for Multi-User Shop Access (Sales, Warehouse, Accounts, Admin)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        full_name TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'SALESMAN', -- 'ADMIN', 'SALESMAN', 'ACCOUNTING', 'WAREHOUSE'
        phone TEXT,
        is_active INTEGER DEFAULT 1,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # 1. Settings Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS settings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        shop_name TEXT NOT NULL DEFAULT 'SSIPL Bags & Footwear Hub',
        tagline TEXT DEFAULT 'Wholesale & Retail Dealers in Bags & Footwear',
        phone TEXT DEFAULT '+91 98765 43210',
        email TEXT DEFAULT 'contact@stepstyle.com',
        address TEXT DEFAULT 'Shop #12, Market Complex, Wholesale Market',
        tax_number TEXT DEFAULT 'GSTIN: 27AAAAA0000A1Z5',
        currency_symbol TEXT DEFAULT '₹',
        invoice_prefix TEXT DEFAULT 'SSIPL',
        logo_path TEXT DEFAULT '',
        invoice_footer TEXT DEFAULT 'Thank you for your business! Goods once sold can be exchanged within 7 days with original invoice.',
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # Safe auto-migration for existing databases
    try:
        cursor.execute("ALTER TABLE settings ADD COLUMN logo_path TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE settings ADD COLUMN invoice_prefix TEXT DEFAULT 'SSIPL'")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE sales ADD COLUMN salesman_id INTEGER")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE sales ADD COLUMN salesman_name TEXT DEFAULT 'Staff'")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE sales ADD COLUMN estimate_id INTEGER")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE sales ADD COLUMN advance_adjusted REAL DEFAULT 0.0")
    except sqlite3.OperationalError:
        pass

    # GST Compliance Columns Migration
    gst_migrations = [
        ("ALTER TABLE settings ADD COLUMN state_code TEXT DEFAULT '27'", "settings.state_code"),
        ("ALTER TABLE settings ADD COLUMN state_name TEXT DEFAULT 'Maharashtra'", "settings.state_name"),
        ("ALTER TABLE products ADD COLUMN hsn_code TEXT DEFAULT '4202'", "products.hsn_code"),
        ("ALTER TABLE products ADD COLUMN gst_rate REAL DEFAULT 18.0", "products.gst_rate"),
        
        ("ALTER TABLE sales ADD COLUMN taxable_amount REAL DEFAULT 0.0", "sales.taxable_amount"),
        ("ALTER TABLE sales ADD COLUMN cgst_amount REAL DEFAULT 0.0", "sales.cgst_amount"),
        ("ALTER TABLE sales ADD COLUMN sgst_amount REAL DEFAULT 0.0", "sales.sgst_amount"),
        ("ALTER TABLE sales ADD COLUMN igst_amount REAL DEFAULT 0.0", "sales.igst_amount"),
        ("ALTER TABLE sales ADD COLUMN total_tax REAL DEFAULT 0.0", "sales.total_tax"),
        ("ALTER TABLE sales ADD COLUMN is_interstate INTEGER DEFAULT 0", "sales.is_interstate"),

        ("ALTER TABLE sale_items ADD COLUMN hsn_code TEXT DEFAULT ''", "sale_items.hsn_code"),
        ("ALTER TABLE sale_items ADD COLUMN gst_rate REAL DEFAULT 18.0", "sale_items.gst_rate"),
        ("ALTER TABLE sale_items ADD COLUMN taxable_amount REAL DEFAULT 0.0", "sale_items.taxable_amount"),
        ("ALTER TABLE sale_items ADD COLUMN cgst_amount REAL DEFAULT 0.0", "sale_items.cgst_amount"),
        ("ALTER TABLE sale_items ADD COLUMN sgst_amount REAL DEFAULT 0.0", "sale_items.sgst_amount"),
        ("ALTER TABLE sale_items ADD COLUMN igst_amount REAL DEFAULT 0.0", "sale_items.igst_amount"),
        ("ALTER TABLE sale_items ADD COLUMN tax_amount REAL DEFAULT 0.0", "sale_items.tax_amount"),

        ("ALTER TABLE purchases ADD COLUMN taxable_amount REAL DEFAULT 0.0", "purchases.taxable_amount"),
        ("ALTER TABLE purchases ADD COLUMN cgst_amount REAL DEFAULT 0.0", "purchases.cgst_amount"),
        ("ALTER TABLE purchases ADD COLUMN sgst_amount REAL DEFAULT 0.0", "purchases.sgst_amount"),
        ("ALTER TABLE purchases ADD COLUMN igst_amount REAL DEFAULT 0.0", "purchases.igst_amount"),
        ("ALTER TABLE purchases ADD COLUMN total_tax REAL DEFAULT 0.0", "purchases.total_tax"),
        ("ALTER TABLE purchases ADD COLUMN is_interstate INTEGER DEFAULT 0", "purchases.is_interstate"),

        ("ALTER TABLE purchase_items ADD COLUMN hsn_code TEXT DEFAULT ''", "purchase_items.hsn_code"),
        ("ALTER TABLE purchase_items ADD COLUMN gst_rate REAL DEFAULT 18.0", "purchase_items.gst_rate"),
        ("ALTER TABLE purchase_items ADD COLUMN taxable_amount REAL DEFAULT 0.0", "purchase_items.taxable_amount"),
        ("ALTER TABLE purchase_items ADD COLUMN tax_amount REAL DEFAULT 0.0", "purchase_items.tax_amount"),

        ("ALTER TABLE estimates ADD COLUMN taxable_amount REAL DEFAULT 0.0", "estimates.taxable_amount"),
        ("ALTER TABLE estimates ADD COLUMN cgst_amount REAL DEFAULT 0.0", "estimates.cgst_amount"),
        ("ALTER TABLE estimates ADD COLUMN sgst_amount REAL DEFAULT 0.0", "estimates.sgst_amount"),
        ("ALTER TABLE estimates ADD COLUMN igst_amount REAL DEFAULT 0.0", "estimates.igst_amount"),
        ("ALTER TABLE estimates ADD COLUMN total_tax REAL DEFAULT 0.0", "estimates.total_tax"),
        ("ALTER TABLE estimates ADD COLUMN is_interstate INTEGER DEFAULT 0", "estimates.is_interstate"),

        ("ALTER TABLE estimate_items ADD COLUMN hsn_code TEXT DEFAULT ''", "estimate_items.hsn_code"),
        ("ALTER TABLE estimate_items ADD COLUMN gst_rate REAL DEFAULT 18.0", "estimate_items.gst_rate"),
        ("ALTER TABLE estimate_items ADD COLUMN taxable_amount REAL DEFAULT 0.0", "estimate_items.taxable_amount"),
        ("ALTER TABLE estimate_items ADD COLUMN tax_amount REAL DEFAULT 0.0", "estimate_items.tax_amount"),

        ("ALTER TABLE parties ADD COLUMN state_code TEXT DEFAULT '27'", "parties.state_code"),
        ("ALTER TABLE parties ADD COLUMN state_name TEXT DEFAULT 'Maharashtra'", "parties.state_name"),
    ]
    for sql_query, col in gst_migrations:
        try:
            cursor.execute(sql_query)
        except sqlite3.OperationalError:
            pass

    # Ensure standard Indian HSN codes & GST rates are assigned to catalog
    cursor.execute("UPDATE products SET hsn_code = '4202', gst_rate = 18.0 WHERE category = 'Bags' AND (hsn_code IS NULL OR hsn_code = '' OR hsn_code = '4202')")
    cursor.execute("UPDATE products SET hsn_code = '6404', gst_rate = 18.0 WHERE category = 'Shoes' AND (hsn_code IS NULL OR hsn_code = '' OR hsn_code = '4202')")

    # Ensure shop name is updated to SSIPL
    cursor.execute("UPDATE settings SET shop_name = 'SSIPL Bags & Footwear Hub' WHERE shop_name LIKE '%Apex%'")
    cursor.execute("UPDATE users SET full_name = 'Sales Staff 1' WHERE username = 'sales1' AND full_name = 'Floor Salesman'")

    # Set default logo from static/img/logo.png if logo_path is empty or missing
    static_logo_file = os.path.join(os.path.dirname(__file__), 'static', 'img', 'logo.png')
    if os.path.exists(static_logo_file):
        import base64
        cursor.execute("SELECT logo_path FROM settings ORDER BY id ASC LIMIT 1")
        row = cursor.fetchone()
        if row and (not row[0] or len(str(row[0]).strip()) == 0 or not str(row[0]).startswith('data:image')):
            try:
                with open(static_logo_file, 'rb') as f:
                    b64 = base64.b64encode(f.read()).decode('utf-8')
                cursor.execute("UPDATE settings SET logo_path = ? WHERE id = (SELECT id FROM settings ORDER BY id ASC LIMIT 1)", (f"data:image/png;base64,{b64}",))
            except Exception:
                pass

    # 2. Products Table (Tailored for Bags & Shoes)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sku TEXT UNIQUE NOT NULL,
        name TEXT NOT NULL,
        category TEXT NOT NULL,        -- 'Bags' or 'Shoes' or 'Accessories'
        subcategory TEXT,             -- 'Backpack', 'Luggage', 'Sneakers', 'Formal Shoes', etc.
        brand TEXT,
        size TEXT,                    -- e.g. '8', '9', '32L', 'Free Size'
        color TEXT,                   -- e.g. 'Black', 'Brown', 'Navy'
        unit TEXT DEFAULT 'Pairs',    -- 'Pairs', 'Pieces', 'Dozens', 'Cartons'
        cost_price REAL NOT NULL,     -- Purchase cost per unit
        wholesale_price REAL NOT NULL,-- Wholesale selling price per unit
        retail_price REAL NOT NULL,   -- Retail selling price per unit
        stock_qty INTEGER NOT NULL DEFAULT 0,
        min_alert_qty INTEGER NOT NULL DEFAULT 5,
        location TEXT DEFAULT 'Warehouse Rack A',
        notes TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # 3. Parties / Customer / Supplier Directory
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS parties (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        phone TEXT,
        party_type TEXT NOT NULL,      -- 'WHOLESALE_BUYER', 'RETAIL_CUSTOMER', 'SUPPLIER'
        address TEXT,
        gst_no TEXT,
        current_balance REAL DEFAULT 0.0, -- Positive means they owe us (Receivable), Negative means we owe them
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # 4. Purchases (Stock-In Bulk Shipments)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS purchases (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_no TEXT UNIQUE NOT NULL,
        supplier_id INTEGER,
        supplier_name TEXT NOT NULL,
        purchase_date DATE NOT NULL,
        total_amount REAL NOT NULL,
        paid_amount REAL DEFAULT 0.0,
        payment_status TEXT DEFAULT 'PAID', -- 'PAID', 'PARTIAL', 'DUE'
        notes TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (supplier_id) REFERENCES parties(id) ON DELETE SET NULL
    );
    """)

    # 5. Purchase Items
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS purchase_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        purchase_id INTEGER NOT NULL,
        product_id INTEGER NOT NULL,
        quantity INTEGER NOT NULL,
        unit_cost REAL NOT NULL,
        total_cost REAL NOT NULL,
        FOREIGN KEY (purchase_id) REFERENCES purchases(id) ON DELETE CASCADE,
        FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE CASCADE
    );
    """)

    # 6. Sales (Stock-Out Billing: Wholesale & Retail)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS sales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_no TEXT UNIQUE NOT NULL,
        sale_type TEXT NOT NULL,       -- 'WHOLESALE' or 'RETAIL'
        party_id INTEGER,
        customer_name TEXT NOT NULL DEFAULT 'Walk-in Customer',
        customer_phone TEXT,
        subtotal REAL NOT NULL,
        discount REAL DEFAULT 0.0,
        grand_total REAL NOT NULL,
        payment_mode TEXT NOT NULL DEFAULT 'CASH', -- 'CASH', 'UPI', 'CARD', 'CREDIT', 'MIXED'
        payment_status TEXT NOT NULL DEFAULT 'PAID', -- 'PAID', 'PARTIAL', 'CREDIT'
        paid_amount REAL NOT NULL,
        balance_amount REAL DEFAULT 0.0,
        notes TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (party_id) REFERENCES parties(id) ON DELETE SET NULL
    );
    """)

    # 7. Sale Items
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS sale_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_id INTEGER NOT NULL,
        product_id INTEGER NOT NULL,
        product_name TEXT NOT NULL,
        sku TEXT,
        size TEXT,
        color TEXT,
        unit TEXT,
        quantity INTEGER NOT NULL,
        unit_price REAL NOT NULL,
        cost_price REAL NOT NULL,     -- Snapshot of cost price to calculate accurate profit
        total_price REAL NOT NULL,
        FOREIGN KEY (sale_id) REFERENCES sales(id) ON DELETE CASCADE,
        FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE RESTRICT
    );
    """)

    # 8. Stock Movement Audit Log
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS stock_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_id INTEGER NOT NULL,
        change_type TEXT NOT NULL,     -- 'PURCHASE_IN', 'SALE_OUT', 'MANUAL_ADJUSTMENT', 'RETURN'
        quantity_change INTEGER NOT NULL,
        stock_after INTEGER NOT NULL,
        reference_no TEXT,             -- Invoice number or adjustment note
        reason TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE CASCADE
    );
    """)

    # 9. Estimates / Sales Orders (Advance Booking with Later Dispatch & Billing)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS estimates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        estimate_no TEXT UNIQUE NOT NULL,       -- e.g. 'EST-2026-0001'
        party_id INTEGER,
        customer_name TEXT NOT NULL,
        customer_phone TEXT,
        order_type TEXT NOT NULL DEFAULT 'WHOLESALE', -- 'WHOLESALE' or 'RETAIL'
        subtotal REAL NOT NULL,
        discount REAL DEFAULT 0.0,
        grand_total REAL NOT NULL,
        advance_amount REAL DEFAULT 0.0,
        advance_payment_mode TEXT DEFAULT 'CASH', -- 'CASH', 'UPI', 'BANK_TRANSFER', 'CHEQUE'
        balance_amount REAL DEFAULT 0.0,
        status TEXT NOT NULL DEFAULT 'BOOKED', -- 'BOOKED', 'DISPATCHED', 'CANCELLED'
        expected_dispatch_date DATE,
        converted_invoice_id INTEGER,          -- References sales(id) when dispatched
        converted_invoice_no TEXT,             -- Official invoice number
        salesman_id INTEGER,
        salesman_name TEXT DEFAULT 'Staff',
        notes TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (party_id) REFERENCES parties(id) ON DELETE SET NULL,
        FOREIGN KEY (converted_invoice_id) REFERENCES sales(id) ON DELETE SET NULL
    );
    """)

    # 10. Estimate Items
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS estimate_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        estimate_id INTEGER NOT NULL,
        product_id INTEGER NOT NULL,
        product_name TEXT NOT NULL,
        sku TEXT,
        size TEXT,
        color TEXT,
        unit TEXT,
        quantity INTEGER NOT NULL,
        unit_price REAL NOT NULL,
        total_price REAL NOT NULL,
        FOREIGN KEY (estimate_id) REFERENCES estimates(id) ON DELETE CASCADE,
        FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE RESTRICT
    );
    """)

    # 11. Invoice Sequences Table for Concurrency-Safe Bill Numbering across multiple salesmen
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS invoice_sequences (
        sequence_key TEXT PRIMARY KEY,
        current_val INTEGER NOT NULL DEFAULT 0
    );
    """)

    # Check if settings exists
    cursor.execute("SELECT COUNT(*) FROM settings")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO settings (shop_name) VALUES ('SSIPL Bags & Footwear Hub')")

    # Check and seed default role users
    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        default_users = [
            ('admin', generate_password_hash('admin123'), 'Store Owner (Admin)', 'ADMIN', '+91 98765 00001'),
            ('accounts', generate_password_hash('accounts123'), 'Shop Accountant', 'ACCOUNTING', '+91 98765 00003'),
            ('godown', generate_password_hash('godown123'), 'Warehouse Incharge', 'WAREHOUSE', '+91 98765 00004')
        ]
        cursor.executemany("""
            INSERT INTO users (username, password_hash, full_name, role, phone)
            VALUES (?, ?, ?, ?, ?)
        """, default_users)

    # Ensure 10 salesmen accounts (sales1 to sales10) exist for simultaneous multi-counter billing
    for i in range(1, 11):
        uname = f'sales{i}'
        cursor.execute("SELECT id FROM users WHERE username = ?", (uname,))
        if not cursor.fetchone():
            cursor.execute("""
                INSERT INTO users (username, password_hash, full_name, role, phone)
                VALUES (?, ?, ?, 'SALESMAN', ?)
            """, (uname, generate_password_hash('sales123'), f'Sales Staff {i}', f'+91 98765 000{i:02d}'))

    conn.commit()

    # Seed Sample Data if products table is empty
    seed_sample_data(conn)
    conn.close()

def get_next_invoice_number(conn, prefix='SSIPL', year_str=None):
    """
    Thread-safe & concurrency-safe atomic invoice generator.
    Guarantees that when 10 salesmen checkout at the same millisecond,
    each invoice receives a strictly consecutive, unique number (e.g. SSIPL-2026-0001).
    """
    if not year_str:
        year_str = datetime.now().strftime("%Y")
    clean_prefix = (prefix or 'SSIPL').strip()
    seq_key = f"{clean_prefix}-{year_str}"

    cur = conn.cursor()
    # Initialize sequence row if not present, seeding from highest existing bill number
    cur.execute("SELECT current_val FROM invoice_sequences WHERE sequence_key = ?", (seq_key,))
    row = cur.fetchone()
    if row is None:
        cur.execute("SELECT invoice_no FROM sales WHERE invoice_no LIKE ?", (f"{seq_key}-%",))
        sales = cur.fetchall()
        max_num = 0
        for s in sales:
            try:
                parts = s['invoice_no'].split('-')
                num = int(parts[-1])
                if num > max_num:
                    max_num = num
            except (ValueError, IndexError, TypeError, KeyError):
                pass
        cur.execute("INSERT OR REPLACE INTO invoice_sequences (sequence_key, current_val) VALUES (?, ?)", (seq_key, max_num))

    # Atomically increment
    cur.execute("UPDATE invoice_sequences SET current_val = current_val + 1 WHERE sequence_key = ?", (seq_key,))
    cur.execute("SELECT current_val FROM invoice_sequences WHERE sequence_key = ?", (seq_key,))
    next_val = cur.fetchone()[0]
    return f"{clean_prefix}-{year_str}-{next_val:04d}"

def seed_sample_data(conn):
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM products")
    if cursor.fetchone()[0] > 0:
        return

    sample_products = [
        # Bags
        ('BAG-BP-001', 'Pro Athlete 35L Laptop Backpack', 'Bags', 'Backpack', 'WildTrail', '35L', 'Black', 'Pieces', 480.0, 720.0, 1199.0, 45, 10, 'Rack B-1', 'Water resistant with rain cover'),
        ('BAG-BP-002', 'Pro Athlete 35L Laptop Backpack', 'Bags', 'Backpack', 'WildTrail', '35L', 'Navy Blue', 'Pieces', 480.0, 720.0, 1199.0, 32, 10, 'Rack B-1', 'Water resistant'),
        ('BAG-HB-101', 'Elegance Vegan Leather Handbag', 'Bags', 'Handbag', 'Lavie Aura', 'Medium', 'Brown', 'Pieces', 650.0, 950.0, 1599.0, 20, 5, 'Rack B-3', 'Includes sling belt'),
        ('BAG-HB-102', 'Elegance Vegan Leather Handbag', 'Bags', 'Handbag', 'Lavie Aura', 'Medium', 'Beige', 'Pieces', 650.0, 950.0, 1599.0, 4, 5, 'Rack B-3', 'Low stock alert!'),
        ('BAG-TR-201', 'Cabin Size Hard-Trolley Suitcase', 'Bags', 'Luggage', 'SkyFlyer', '20 Inch', 'Metallic Grey', 'Pieces', 1450.0, 2100.0, 3299.0, 15, 4, 'Rack L-1', '4-wheel 360 spinner'),
        ('BAG-SC-301', 'Kids Orthopedic School Bag', 'Bags', 'School Bag', 'WonderKids', 'Standard', 'Red/Blue', 'Pieces', 240.0, 360.0, 599.0, 60, 15, 'Rack B-4', 'Bulk seller item'),
        ('BAG-WL-401', 'Genuine Leather Men RFID Wallet', 'Bags', 'Wallet', 'UrbanHide', 'Standard', 'Black', 'Pieces', 120.0, 200.0, 399.0, 75, 15, 'Display Counter', 'Gift box packing'),

        # Shoes
        ('SH-SNK-01-07', 'Air Cushion Pace Running Sneakers', 'Shoes', 'Sneakers', 'AeroSport', 'UK 7', 'White/Grey', 'Pairs', 550.0, 800.0, 1299.0, 24, 6, 'Rack S-1', 'Soft EVA sole'),
        ('SH-SNK-01-08', 'Air Cushion Pace Running Sneakers', 'Shoes', 'Sneakers', 'AeroSport', 'UK 8', 'White/Grey', 'Pairs', 550.0, 800.0, 1299.0, 30, 6, 'Rack S-1', 'High demand size'),
        ('SH-SNK-01-09', 'Air Cushion Pace Running Sneakers', 'Shoes', 'Sneakers', 'AeroSport', 'UK 9', 'White/Grey', 'Pairs', 550.0, 800.0, 1299.0, 18, 6, 'Rack S-1', 'Breathable mesh'),
        ('SH-SNK-01-10', 'Air Cushion Pace Running Sneakers', 'Shoes', 'Sneakers', 'AeroSport', 'UK 10', 'White/Grey', 'Pairs', 550.0, 800.0, 1299.0, 3, 6, 'Rack S-1', 'Needs restocking!'),
        ('SH-FRM-02-08', 'Classic Oxford Formal Leather Shoes', 'Shoes', 'Formal', 'Kingsley', 'UK 8', 'Black', 'Pairs', 750.0, 1100.0, 1799.0, 16, 5, 'Rack S-3', 'Handcrafted finish'),
        ('SH-FRM-02-09', 'Classic Oxford Formal Leather Shoes', 'Shoes', 'Formal', 'Kingsley', 'UK 9', 'Black', 'Pairs', 750.0, 1100.0, 1799.0, 12, 5, 'Rack S-3', 'Handcrafted finish'),
        ('SH-WHL-03-06', 'Women Block Heel Party Sandals', 'Shoes', 'Heels', 'DivaWalk', 'UK 6', 'Rose Gold', 'Pairs', 380.0, 560.0, 899.0, 22, 6, 'Rack S-5', 'Comfort cushioned insole'),
        ('SH-SLP-04-08', 'Ultra-Light Daily Slide Slippers', 'Shoes', 'Slippers', 'ComfySole', 'UK 8', 'Navy Blue', 'Pairs', 95.0, 145.0, 249.0, 95, 20, 'Bulk Box 1', 'Fast retail turn-around')
    ]

    cursor.executemany("""
    INSERT INTO products (sku, name, category, subcategory, brand, size, color, unit, cost_price, wholesale_price, retail_price, stock_qty, min_alert_qty, location, notes)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, sample_products)

    # Sample Suppliers and Wholesale Buyers
    sample_parties = [
        ('Apex Factory Distributors (Direct Mill)', '+91 98200 11223', 'SUPPLIER', 'Industrial Area, Phase 2', '27AABCA1234F1Z8', 0.0),
        ('Zenith Footwear Manufacturers', '+91 97111 44556', 'SUPPLIER', 'Shoe Hub, Agra', '09AACCB5555G1Z2', 0.0),
        ('Metro Fashion Retailers (Wholesale Buyer)', '+91 98990 88776', 'WHOLESALE_BUYER', 'Main Market, City Center', '27BBDCE9988H1Z1', 12500.0),
        ('Gupta Shoe Store (Wholesale Buyer)', '+91 94555 22334', 'WHOLESALE_BUYER', 'Railway Road Bazaar', '', 4800.0),
        ('Rohan Traders', '+91 91234 56789', 'WHOLESALE_BUYER', 'Town Hall Cross Road', '', 0.0)
    ]

    cursor.executemany("""
    INSERT INTO parties (name, phone, party_type, address, gst_no, current_balance)
    VALUES (?, ?, ?, ?, ?, ?)
    """, sample_parties)

    conn.commit()

if __name__ == '__main__':
    init_db()
    print("Database initialized and sample data seeded successfully!")
