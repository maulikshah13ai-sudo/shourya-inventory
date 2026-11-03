"""
Excel & CSV Helper Module for Bags & Shoes Wholesale / Retail Inventory.
Provides template generation and robust parsing for Bulk Purchase and Bulk Sale uploads.
"""

import io
import csv
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADER_FILL_PURCHASE = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Dark Blue
HEADER_FILL_SALE = PatternFill(start_color="047857", end_color="047857", fill_type="solid")     # Emerald Green
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
THIN_BORDER = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)

def generate_bulk_purchase_template(file_format="xlsx"):
    """
    Generates a pre-formatted Bulk Purchase Template with sample data and column guidelines.
    Returns (bytes, filename, mimetype).
    """
    headers = [
        "Supplier Name",
        "Supplier Invoice No",
        "Purchase Date (YYYY-MM-DD)",
        "SKU",
        "Product Name",
        "Category",
        "Subcategory",
        "Brand",
        "Size",
        "Color",
        "Unit",
        "Quantity",
        "Unit Cost Price",
        "Wholesale Price",
        "Retail Price",
        "Location",
        "Notes"
    ]

    sample_rows = [
        [
            "Apex Factory Distributors", "INV-2026-8812", date.today().isoformat(),
            "BAG-BP-001", "Pro Athlete 35L Laptop Backpack", "Bags", "Backpack",
            "WildTrail", "35L", "Black", "Pieces", 50, 480.0, 720.0, 1199.0, "Rack B-1", "Factory direct shipment"
        ],
        [
            "Zenith Footwear Manufacturers", "INV-2026-9021", date.today().isoformat(),
            "SH-SNK-01-08", "Air Cushion Pace Running Sneakers", "Shoes", "Sneakers",
            "AeroSport", "UK 8", "White/Grey", "Pairs", 36, 550.0, 800.0, 1299.0, "Rack S-1", "Seasonal restock"
        ],
        [
            "Apex Factory Distributors", "INV-2026-8812", date.today().isoformat(),
            "BAG-TR-201", "Cabin Size Hard-Trolley Suitcase", "Bags", "Luggage",
            "SkyFlyer", "20 Inch", "Metallic Grey", "Pieces", 20, 1450.0, 2100.0, 3299.0, "Rack L-1", "Same invoice item 2"
        ]
    ]

    if file_format.lower() == "csv":
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(headers)
        for r in sample_rows:
            writer.writerow(r)
        return output.getvalue().encode('utf-8-sig'), "bulk_purchase_template.csv", "text/csv"

    # Excel XLSX
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bulk_Purchase_Data"
    ws.views.sheetView[0].showGridLines = True

    # Write Header
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL_PURCHASE
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Write Sample Rows
    for row_data in sample_rows:
        ws.append(row_data)

    # Styling cells
    for row in ws.iter_rows(min_row=2, max_row=len(sample_rows) + 1, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = THIN_BORDER
            cell.font = Font(name="Calibri", size=10)
            if isinstance(cell.value, (int, float)):
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="left")

    # Column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Add Instructions Sheet
    ws_notes = wb.create_sheet(title="Instructions")
    ws_notes.views.sheetView[0].showGridLines = True
    instructions = [
        ["Bulk Purchase Excel Upload Guidelines"],
        [""],
        ["Column Name", "Required?", "Description / Rules"],
        ["Supplier Name", "Yes", "Name of the manufacturer / wholesale supplier. If not in system, will be added."],
        ["Supplier Invoice No", "Yes", "Invoice number from your supplier. Rows with the same invoice number are grouped into 1 shipment."],
        ["Purchase Date", "Optional", "Date in YYYY-MM-DD or DD-MM-YYYY format. Defaults to today if left blank."],
        ["SKU", "Yes", "Unique product SKU code (e.g. BAG-BP-001, SH-SNK-01-08)."],
        ["Product Name", "Yes", "Name or description of the bag/shoe."],
        ["Category", "Yes", "'Bags', 'Shoes', or 'Accessories'."],
        ["Quantity", "Yes", "Number of pieces or pairs purchased (positive whole number)."],
        ["Unit Cost Price", "Yes", "Purchase cost per unit from supplier (INR)."],
        ["Wholesale Price", "Optional", "Wholesale selling rate. Updates product catalog rate if provided."],
        ["Retail Price", "Optional", "Retail MRP selling rate. Updates product catalog rate if provided."],
        ["Location", "Optional", "Shelf or rack location (e.g. 'Rack B-1', 'Godown Box 4')."]
    ]
    for r in instructions:
        ws_notes.append(r)
    
    ws_notes.cell(row=1, column=1).font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    for c in range(1, 4):
        header_cell = ws_notes.cell(row=3, column=c)
        header_cell.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
        header_cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    for col in ws_notes.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_notes.column_dimensions[col_letter].width = max(max_len + 3, 15)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), "bulk_purchase_template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def generate_bulk_sale_template(file_format="xlsx"):
    """
    Generates a pre-formatted Bulk Sale Template with sample data and column guidelines.
    Returns (bytes, filename, mimetype).
    """
    headers = [
        "Customer Name",
        "Customer Phone",
        "Sale Type",
        "Payment Mode",
        "Order Reference",
        "SKU",
        "Quantity",
        "Unit Price",
        "Discount",
        "Paid Amount",
        "Notes"
    ]

    sample_rows = [
        [
            "Metro Fashion Retailers", "9899088776", "WHOLESALE", "CREDIT",
            "ORD-101", "BAG-BP-001", 10, 720.0, 0, 0, "Bulk wholesale delivery 1"
        ],
        [
            "Metro Fashion Retailers", "9899088776", "WHOLESALE", "CREDIT",
            "ORD-101", "SH-SNK-01-08", 12, 800.0, 0, 0, "Same invoice item 2"
        ],
        [
            "Walk-in Customer", "9812345678", "RETAIL", "UPI",
            "ORD-102", "BAG-WL-401", 3, 399.0, 50.0, 1147.0, "Retail counter purchase"
        ]
    ]

    if file_format.lower() == "csv":
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(headers)
        for r in sample_rows:
            writer.writerow(r)
        return output.getvalue().encode('utf-8-sig'), "bulk_sale_template.csv", "text/csv"

    # Excel XLSX
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bulk_Sale_Data"
    ws.views.sheetView[0].showGridLines = True

    # Write Header
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL_SALE
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Write Sample Rows
    for row_data in sample_rows:
        ws.append(row_data)

    # Styling cells
    for row in ws.iter_rows(min_row=2, max_row=len(sample_rows) + 1, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = THIN_BORDER
            cell.font = Font(name="Calibri", size=10)
            if isinstance(cell.value, (int, float)):
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="left")

    # Column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

    # Add Instructions Sheet
    ws_notes = wb.create_sheet(title="Instructions")
    ws_notes.views.sheetView[0].showGridLines = True
    instructions = [
        ["Bulk Sale Excel Upload Guidelines"],
        [""],
        ["Column Name", "Required?", "Description / Rules"],
        ["Customer Name", "Yes", "Buyer name (e.g. 'Metro Fashion' or 'Walk-in Customer')."],
        ["Customer Phone", "Optional", "Customer mobile / WhatsApp number."],
        ["Sale Type", "Yes", "'WHOLESALE' or 'RETAIL'."],
        ["Payment Mode", "Yes", "'CASH', 'UPI', 'CARD', or 'CREDIT' (Udhar)."],
        ["Order Reference", "Optional", "Rows with the same Order Reference will be merged into a SINGLE invoice with multiple items."],
        ["SKU", "Yes", "Product SKU code must match an existing product in inventory."],
        ["Quantity", "Yes", "Number of pieces or pairs sold."],
        ["Unit Price", "Yes", "Selling rate per unit charged to customer."],
        ["Discount", "Optional", "Discount amount (INR) on the bill, if any."],
        ["Paid Amount", "Optional", "Amount customer paid today. For full credit, leave 0 or empty."]
    ]
    for r in instructions:
        ws_notes.append(r)
    
    ws_notes.cell(row=1, column=1).font = Font(name="Calibri", size=14, bold=True, color="047857")
    for c in range(1, 4):
        header_cell = ws_notes.cell(row=3, column=c)
        header_cell.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
        header_cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    for col in ws_notes.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_notes.column_dimensions[col_letter].width = max(max_len + 3, 15)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), "bulk_sale_template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def normalize_col_name(name):
    """Normalize column headers: lower, strip, remove spaces, parentheses, underscores."""
    if not name:
        return ""
    s = str(name).strip().lower()
    for ch in [' ', '_', '-', '(', ')', '/', '.', ':']:
        s = s.replace(ch, '')
    return s


def parse_uploaded_file(file_storage):
    """
    Parses an uploaded file (either .xlsx or .csv) from Flask request.files.
    Returns list of dicts with normalized keys.
    """
    filename = (file_storage.filename or '').lower()
    rows = []

    if filename.endswith('.xlsx'):
        if hasattr(file_storage, 'read'):
            raw_content = file_storage.read()
            if hasattr(file_storage, 'seek'):
                file_storage.seek(0)
            wb = openpyxl.load_workbook(io.BytesIO(raw_content), data_only=True)
        else:
            wb = openpyxl.load_workbook(file_storage, data_only=True)
        # Choose the first sheet or the one with 'data' in name
        sheet = wb.active
        for name in wb.sheetnames:
            if 'data' in name.lower() or 'purchase' in name.lower() or 'sale' in name.lower():
                sheet = wb[name]
                break

        header_cells = [cell.value for cell in sheet[1]]
        # Map column index to normalized key
        header_map = {}
        for idx, h in enumerate(header_cells):
            if h:
                norm_h = normalize_col_name(h)
                header_map[idx] = norm_h

        for r_idx in range(2, sheet.max_row + 1):
            row_dict = {}
            has_data = False
            for c_idx, key in header_map.items():
                val = sheet.cell(row=r_idx, column=c_idx + 1).value
                if val is not None and str(val).strip() != "":
                    has_data = True
                row_dict[key] = val
            if has_data:
                rows.append(row_dict)

    elif filename.endswith('.csv'):
        raw_bytes = file_storage.read()
        try:
            text = raw_bytes.decode('utf-8-sig')
        except UnicodeDecodeError:
            text = raw_bytes.decode('latin1')

        reader = csv.reader(io.StringIO(text))
        header_row = next(reader, None)
        if not header_row:
            return []

        header_map = {}
        for idx, h in enumerate(header_row):
            if h:
                header_map[idx] = normalize_col_name(h)

        for row in reader:
            if not row or not any(x.strip() for x in row):
                continue
            row_dict = {}
            for idx, key in header_map.items():
                if idx < len(row):
                    row_dict[key] = row[idx].strip()
                else:
                    row_dict[key] = ""
            rows.append(row_dict)
    else:
        raise ValueError("Unsupported file format! Please upload an .xlsx Excel or .csv file.")

    return rows


def get_field_val(row_dict, candidate_keys, default=''):
    """Finds the first matching key in candidate_keys that has a non-empty value."""
    for k in candidate_keys:
        val = row_dict.get(k)
        if val is not None and str(val).strip() != "":
            return val
    return default


def extract_purchase_row_data(row):
    """Safely extracts and standardizes a single bulk purchase row."""
    supplier_name = str(get_field_val(row, ['suppliername', 'supplier', 'partyname', 'vendor', 'name'], 'Direct Supplier')).strip()
    invoice_no = str(get_field_val(row, ['supplierinvoiceno', 'invoiceno', 'billno', 'invoice', 'refno', 'reference'], '')).strip()
    
    raw_date = get_field_val(row, ['purchasedateyyyymmdd', 'purchasedate', 'date'], date.today().isoformat())
    if isinstance(raw_date, (datetime, date)):
        purchase_date = raw_date.strftime("%Y-%m-%d")
    else:
        purchase_date = str(raw_date).strip()[:10] if str(raw_date).strip() else date.today().isoformat()

    sku = str(get_field_val(row, ['sku', 'itemcode', 'barcode', 'productcode', 'code'], '')).strip().upper()
    name = str(get_field_val(row, ['productname', 'name', 'itemname', 'description'], '')).strip()
    if not name and sku:
        name = f"Product {sku}"

    category = str(get_field_val(row, ['category', 'cat'], 'Bags')).strip().capitalize()
    if 'shoe' in category.lower():
        category = 'Shoes'
    elif 'bag' in category.lower():
        category = 'Bags'

    subcategory = str(get_field_val(row, ['subcategory', 'subcat', 'type'], '')).strip()
    brand = str(get_field_val(row, ['brand', 'maker', 'company'], '')).strip()
    size = str(get_field_val(row, ['size'], '')).strip()
    color = str(get_field_val(row, ['color', 'colour'], '')).strip()
    unit = str(get_field_val(row, ['unit', 'uom'], 'Pieces' if category == 'Bags' else 'Pairs')).strip()

    try:
        quantity = int(float(get_field_val(row, ['quantity', 'qty', 'count', 'units'], 1)))
    except (ValueError, TypeError):
        quantity = 1

    try:
        cost_price = float(get_field_val(row, ['unitcostprice', 'costprice', 'purchaseprice', 'rate', 'cost', 'unitcost'], 0.0))
    except (ValueError, TypeError):
        cost_price = 0.0

    try:
        wholesale_price = float(get_field_val(row, ['wholesaleprice', 'wholesalerate', 'wsprice'], cost_price * 1.3))
    except (ValueError, TypeError):
        wholesale_price = cost_price * 1.3

    try:
        retail_price = float(get_field_val(row, ['retailprice', 'mrp', 'sellingprice'], wholesale_price * 1.4))
    except (ValueError, TypeError):
        retail_price = wholesale_price * 1.4

    location = str(get_field_val(row, ['location', 'rack', 'shelf'], 'Warehouse Rack A')).strip()
    notes = str(get_field_val(row, ['notes', 'note', 'remark', 'remarks'], '')).strip()

    return {
        'supplier_name': supplier_name,
        'invoice_no': invoice_no,
        'purchase_date': purchase_date,
        'sku': sku,
        'name': name,
        'category': category,
        'subcategory': subcategory,
        'brand': brand,
        'size': size,
        'color': color,
        'unit': unit,
        'quantity': max(1, quantity),
        'cost_price': max(0.0, cost_price),
        'wholesale_price': max(0.0, wholesale_price),
        'retail_price': max(0.0, retail_price),
        'location': location,
        'notes': notes
    }


def extract_sale_row_data(row):
    """Safely extracts and standardizes a single bulk sale row."""
    customer_name = str(get_field_val(row, ['customername', 'customer', 'partyname', 'buyer', 'party'], 'Walk-in Customer')).strip()
    customer_phone = str(get_field_val(row, ['customerphone', 'phone', 'mobile', 'whatsapp', 'contact'], '')).strip()
    
    sale_type = str(get_field_val(row, ['saletype', 'type'], 'WHOLESALE')).strip().upper()
    if 'retail' in sale_type.lower():
        sale_type = 'RETAIL'
    else:
        sale_type = 'WHOLESALE'

    payment_mode = str(get_field_val(row, ['paymentmode', 'paymode', 'mode'], 'CASH')).strip().upper()
    if payment_mode not in ['CASH', 'UPI', 'CARD', 'CREDIT', 'MIXED', 'BANK_TRANSFER']:
        payment_mode = 'CASH'

    order_ref = str(get_field_val(row, ['orderreference', 'orderref', 'refno', 'invoiceref', 'billno', 'orderid'], '')).strip()
    sku = str(get_field_val(row, ['sku', 'itemcode', 'barcode', 'code'], '')).strip().upper()

    try:
        quantity = int(float(get_field_val(row, ['quantity', 'qty', 'count'], 1)))
    except (ValueError, TypeError):
        quantity = 1

    try:
        unit_price = float(get_field_val(row, ['unitprice', 'price', 'rate', 'sellingprice'], 0.0))
    except (ValueError, TypeError):
        unit_price = 0.0

    try:
        discount = float(get_field_val(row, ['discount', 'disc'], 0.0))
    except (ValueError, TypeError):
        discount = 0.0

    raw_paid = get_field_val(row, ['paidamount', 'paid', 'advance'], None)
    if raw_paid is not None and str(raw_paid).strip() != '':
        try:
            paid_amount = float(raw_paid)
        except (ValueError, TypeError):
            paid_amount = None
    else:
        paid_amount = None

    notes = str(get_field_val(row, ['notes', 'note', 'remark'], '')).strip()

    return {
        'customer_name': customer_name,
        'customer_phone': customer_phone,
        'sale_type': sale_type,
        'payment_mode': payment_mode,
        'order_ref': order_ref,
        'sku': sku,
        'quantity': max(1, quantity),
        'unit_price': max(0.0, unit_price),
        'discount': max(0.0, discount),
        'paid_amount': paid_amount,
        'notes': notes
    }

