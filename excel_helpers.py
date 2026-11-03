"""
Excel & CSV Helper Module for Bags & Shoes Wholesale / Retail Inventory.
Provides template generation, spreadsheet parsing, Indian GST tax calculations,
and official GSTR-1 & GSTR-3B report generation compliant with Indian GST laws.
"""

import io
import csv
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADER_FILL_PURCHASE = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Dark Blue
HEADER_FILL_SALE = PatternFill(start_color="047857", end_color="047857", fill_type="solid")     # Emerald Green
HEADER_FILL_GST = PatternFill(start_color="4C1D95", end_color="4C1D95", fill_type="solid")      # Deep Purple
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
THIN_BORDER = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)


# ----------------- INDIAN GST TAX CALCULATIONS -----------------
def calculate_gst_breakdown(items, is_interstate=False, prices_include_gst=True):
    """
    Computes exact Indian GST breakdown for line items.
    Bags default HSN: 4202 (18% GST).
    Footwear default HSN: 6404 (18% or 12% GST).

    Returns:
      taxable_total, cgst_total, sgst_total, igst_total, total_tax, grand_total, processed_items, hsn_summary
    """
    taxable_total = 0.0
    cgst_total = 0.0
    sgst_total = 0.0
    igst_total = 0.0
    total_tax = 0.0
    grand_total = 0.0

    processed_items = []
    hsn_dict = {}

    for item in items:
        qty = float(item.get('quantity') or 1)
        unit_price = float(item.get('price') or item.get('unit_price') or item.get('unit_cost') or 0.0)
        gst_rate = float(item.get('gst_rate') if item.get('gst_rate') is not None else 18.0)
        
        name_str = str(item.get('name') or item.get('product_name') or '').lower()
        default_hsn = '4202' if ('bag' in name_str or 'wallet' in name_str or 'luggage' in name_str) else '6404'
        hsn = str(item.get('hsn_code') or default_hsn).strip() or default_hsn
        unit = str(item.get('unit') or ('Pieces' if hsn == '4202' else 'Pairs')).strip()
        prod_name = str(item.get('name') or item.get('product_name') or f"Product {hsn}").strip()

        gross_line_total = qty * unit_price

        if prices_include_gst:
            # Inclusive: Total is gross_line_total, calculate taxable base
            taxable_val = round(gross_line_total / (1.0 + (gst_rate / 100.0)), 2)
            tax_val = round(gross_line_total - taxable_val, 2)
            line_total = gross_line_total
        else:
            # Exclusive: Rate is base price, add GST
            taxable_val = round(gross_line_total, 2)
            tax_val = round(taxable_val * (gst_rate / 100.0), 2)
            line_total = round(taxable_val + tax_val, 2)

        if is_interstate:
            cgst = 0.0
            sgst = 0.0
            igst = tax_val
        else:
            cgst = round(tax_val / 2.0, 2)
            sgst = round(tax_val - cgst, 2)
            igst = 0.0

        item_res = dict(item)
        item_res['hsn_code'] = hsn
        item_res['gst_rate'] = gst_rate
        item_res['taxable_amount'] = taxable_val
        item_res['tax_amount'] = tax_val
        item_res['cgst_amount'] = cgst
        item_res['sgst_amount'] = sgst
        item_res['igst_amount'] = igst
        item_res['total_price'] = line_total
        processed_items.append(item_res)

        taxable_total += taxable_val
        cgst_total += cgst
        sgst_total += sgst
        igst_total += igst
        total_tax += tax_val
        grand_total += line_total

        # Aggregate HSN Table (Indian GST Rule 46 compliance)
        hsn_key = (hsn, gst_rate)
        if hsn_key not in hsn_dict:
            hsn_dict[hsn_key] = {
                'hsn_code': hsn,
                'description': prod_name,
                'uqc': unit,
                'total_qty': 0,
                'gst_rate': gst_rate,
                'taxable_value': 0.0,
                'cgst_amount': 0.0,
                'sgst_amount': 0.0,
                'igst_amount': 0.0,
                'total_tax': 0.0,
                'total_value': 0.0
            }
        hsn_dict[hsn_key]['total_qty'] += qty
        hsn_dict[hsn_key]['taxable_value'] += taxable_val
        hsn_dict[hsn_key]['cgst_amount'] += cgst
        hsn_dict[hsn_key]['sgst_amount'] += sgst
        hsn_dict[hsn_key]['igst_amount'] += igst
        hsn_dict[hsn_key]['total_tax'] += tax_val
        hsn_dict[hsn_key]['total_value'] += line_total

    # Round HSN aggregates
    hsn_list = []
    for h in hsn_dict.values():
        h['taxable_value'] = round(h['taxable_value'], 2)
        h['cgst_amount'] = round(h['cgst_amount'], 2)
        h['sgst_amount'] = round(h['sgst_amount'], 2)
        h['igst_amount'] = round(h['igst_amount'], 2)
        h['total_tax'] = round(h['total_tax'], 2)
        h['total_value'] = round(h['total_value'], 2)
        hsn_list.append(h)

    return {
        'taxable_total': round(taxable_total, 2),
        'cgst_total': round(cgst_total, 2),
        'sgst_total': round(sgst_total, 2),
        'igst_total': round(igst_total, 2),
        'total_tax': round(total_tax, 2),
        'grand_total': round(grand_total, 2),
        'items': processed_items,
        'hsn_summary': hsn_list
    }


# ----------------- TEMPLATE GENERATORS -----------------
def generate_bulk_purchase_template(file_format="xlsx"):
    """
    Generates a pre-formatted Bulk Purchase Template with HSN code and GST rate columns.
    """
    headers = [
        "Supplier Name",
        "Supplier Invoice No",
        "Purchase Date (YYYY-MM-DD)",
        "SKU",
        "Product Name",
        "Category",
        "HSN Code",
        "GST Rate (%)",
        "Subcategory",
        "Brand",
        "Size",
        "Color",
        "Unit",
        "Quantity",
        "Unit Cost Price (Incl GST)",
        "Wholesale Price",
        "Retail Price",
        "Location",
        "Notes"
    ]

    sample_rows = [
        [
            "Apex Factory Distributors", "INV-2026-8812", date.today().isoformat(),
            "BAG-BP-001", "Pro Athlete 35L Laptop Backpack", "Bags", "4202", 18.0, "Backpack",
            "WildTrail", "35L", "Black", "Pieces", 50, 480.0, 720.0, 1199.0, "Rack B-1", "Factory direct shipment"
        ],
        [
            "Zenith Footwear Manufacturers", "INV-2026-9021", date.today().isoformat(),
            "SH-SNK-01-08", "Air Cushion Pace Running Sneakers", "Shoes", "6404", 18.0, "Sneakers",
            "AeroSport", "UK 8", "White/Grey", "Pairs", 36, 550.0, 800.0, 1299.0, "Rack S-1", "Seasonal restock"
        ],
        [
            "Apex Factory Distributors", "INV-2026-8812", date.today().isoformat(),
            "BAG-TR-201", "Cabin Size Hard-Trolley Suitcase", "Bags", "4202", 18.0, "Luggage",
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

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bulk_Purchase_Data"
    ws.views.sheetView[0].showGridLines = True

    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL_PURCHASE
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_data in sample_rows:
        ws.append(row_data)

    for row in ws.iter_rows(min_row=2, max_row=len(sample_rows) + 1, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = THIN_BORDER
            cell.font = Font(name="Calibri", size=10)
            if isinstance(cell.value, (int, float)):
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="left")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Instructions Sheet
    ws_notes = wb.create_sheet(title="GST Guidelines")
    ws_notes.views.sheetView[0].showGridLines = True
    instructions = [
        ["Bulk Purchase Excel Upload Guidelines (Indian GST & HSN Compliant)"],
        [""],
        ["Column Name", "Required?", "Standard Indian GST Rule"],
        ["Supplier Name", "Yes", "Registered Manufacturer / Distributor Name."],
        ["Supplier Invoice No", "Yes", "Supplier Tax Invoice Number."],
        ["Purchase Date", "Optional", "Date in YYYY-MM-DD or DD-MM-YYYY format."],
        ["SKU", "Yes", "Unique product SKU code."],
        ["Product Name", "Yes", "Product description."],
        ["Category", "Yes", "'Bags', 'Shoes', or 'Accessories'."],
        ["HSN Code", "Yes", "HSN Code under Indian GST: Bags/Luggage = 4202, Footwear = 6402/6403/6404."],
        ["GST Rate (%)", "Yes", "Applicable GST % (e.g. 18, 12, 5)."],
        ["Quantity", "Yes", "Units purchased."],
        ["Unit Cost Price", "Yes", "Purchase cost per unit (including GST)."],
        ["Wholesale Price", "Optional", "Wholesale selling rate."],
        ["Retail Price", "Optional", "Retail MRP selling rate."]
    ]
    for r in instructions:
        ws_notes.append(r)
    
    ws_notes.cell(row=1, column=1).font = Font(name="Calibri", size=13, bold=True, color="1E3A8A")
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


# Standard Indian GST State Codes Map (2-digit GSTIN prefix)
GST_STATE_CODES = {
    "01": "Jammu & Kashmir", "02": "Himachal Pradesh", "03": "Punjab", "04": "Chandigarh",
    "05": "Uttarakhand", "06": "Haryana", "07": "Delhi", "08": "Rajasthan",
    "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim", "12": "Arunachal Pradesh",
    "13": "Nagaland", "14": "Manipur", "15": "Mizoram", "16": "Tripura",
    "17": "Meghalaya", "18": "Assam", "19": "West Bengal", "20": "Jharkhand",
    "21": "Odisha", "22": "Chhattisgarh", "23": "Madhya Pradesh", "24": "Gujarat",
    "25": "Daman and Diu", "26": "Dadra and Nagar Haveli", "27": "Maharashtra",
    "29": "Karnataka", "30": "Goa", "31": "Lakshadweep", "32": "Kerala",
    "33": "Tamil Nadu", "34": "Puducherry", "35": "Andaman & Nicobar",
    "36": "Telangana", "37": "Andhra Pradesh", "38": "Ladakh", "97": "Other Territory"
}

def generate_bulk_sale_template(file_format="xlsx"):
    """
    Generates a pre-formatted Bulk Sale Template with HSN code, GST rate, and Party GST Number columns.
    """
    headers = [
        "Customer Name",
        "Customer Phone",
        "Party GST Number",
        "Sale Type",
        "Payment Mode",
        "Order Reference",
        "SKU",
        "HSN Code",
        "GST Rate (%)",
        "Quantity",
        "Unit Price (Incl GST)",
        "Discount",
        "Paid Amount",
        "Notes"
    ]

    sample_rows = [
        [
            "Metro Fashion Retailers", "9899088776", "27AABCS1429B1ZB", "WHOLESALE", "CREDIT",
            "ORD-101", "BAG-BP-001", "4202", 18.0, 10, 720.0, 0, 0, "Bulk wholesale delivery 1"
        ],
        [
            "Metro Fashion Retailers", "9899088776", "27AABCS1429B1ZB", "WHOLESALE", "CREDIT",
            "ORD-101", "SH-SNK-01-08", "6404", 18.0, 12, 800.0, 0, 0, "Same invoice item 2"
        ],
        [
            "Walk-in Customer", "9812345678", "", "RETAIL", "UPI",
            "ORD-102", "BAG-WL-401", "4202", 18.0, 3, 399.0, 50.0, 1147.0, "Retail counter purchase"
        ]
    ]

    if file_format.lower() == "csv":
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(headers)
        for r in sample_rows:
            writer.writerow(r)
        return output.getvalue().encode('utf-8-sig'), "bulk_sale_template.csv", "text/csv"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bulk_Sale_Data"
    ws.views.sheetView[0].showGridLines = True

    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL_SALE
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_data in sample_rows:
        ws.append(row_data)

    for row in ws.iter_rows(min_row=2, max_row=len(sample_rows) + 1, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = THIN_BORDER
            cell.font = Font(name="Calibri", size=10)
            if isinstance(cell.value, (int, float)):
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="left")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), "bulk_sale_template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ----------------- PARSING SPREADSHEETS -----------------
def normalize_col_name(name):
    if not name:
        return ""
    s = str(name).strip().lower()
    for ch in [' ', '_', '-', '(', ')', '/', '.', ':', '%']:
        s = s.replace(ch, '')
    return s


def parse_uploaded_file(file_storage):
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

        sheet = wb.active
        for name in wb.sheetnames:
            if 'data' in name.lower() or 'purchase' in name.lower() or 'sale' in name.lower():
                sheet = wb[name]
                break

        header_cells = [cell.value for cell in sheet[1]]
        header_map = {}
        for idx, h in enumerate(header_cells):
            if h:
                header_map[idx] = normalize_col_name(h)

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
        raise ValueError("Unsupported format! Please upload an .xlsx Excel or .csv file.")

    return rows


def get_field_val(row_dict, candidate_keys, default=''):
    for k in candidate_keys:
        val = row_dict.get(k)
        if val is not None and str(val).strip() != "":
            return val
    return default


def extract_purchase_row_data(row):
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

    hsn_code = str(get_field_val(row, ['hsncode', 'hsn', 'sac', 'hsnsac'], '4202' if category == 'Bags' else '6404')).strip()
    try:
        gst_rate = float(get_field_val(row, ['gstrate', 'gst', 'taxrate', 'tax'], 18.0))
    except (ValueError, TypeError):
        gst_rate = 18.0

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
        cost_price = float(get_field_val(row, ['unitcostpriceinclgst', 'unitcostprice', 'costprice', 'purchaseprice', 'rate', 'cost', 'unitcost'], 0.0))
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
        'hsn_code': hsn_code,
        'gst_rate': gst_rate,
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
    customer_name = str(get_field_val(row, ['customername', 'customer', 'partyname', 'buyer', 'party'], 'Walk-in Customer')).strip()
    customer_phone = str(get_field_val(row, ['customerphone', 'phone', 'mobile', 'whatsapp', 'contact'], '')).strip()
    customer_gst = str(get_field_val(row, ['partygstnumber', 'partygst', 'customergst', 'gstnumber', 'gstno', 'gstin'], '')).strip().upper()
    
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
    hsn_code = str(get_field_val(row, ['hsncode', 'hsn', 'sac'], '')).strip()

    try:
        gst_rate = float(get_field_val(row, ['gstrate', 'gst', 'taxrate'], 18.0))
    except (ValueError, TypeError):
        gst_rate = 18.0

    try:
        quantity = int(float(get_field_val(row, ['quantity', 'qty', 'count'], 1)))
    except (ValueError, TypeError):
        quantity = 1

    try:
        unit_price = float(get_field_val(row, ['unitpriceinclgst', 'unitprice', 'price', 'rate', 'sellingprice'], 0.0))
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
        'customer_gst': customer_gst,
        'sale_type': sale_type,
        'payment_mode': payment_mode,
        'order_ref': order_ref,
        'sku': sku,
        'hsn_code': hsn_code,
        'gst_rate': gst_rate,
        'quantity': max(1, quantity),
        'unit_price': max(0.0, unit_price),
        'discount': max(0.0, discount),
        'paid_amount': paid_amount,
        'notes': notes
    }


# ----------------- INDIAN GSTR-1 & GSTR-3B REPORT GENERATION -----------------
def generate_gstr1_excel_report(conn, start_date=None, end_date=None):
    """
    Generates official Chartered Accountant / GST Portal compatible GSTR-1 Excel Return:
    - Tab 1: Executive Summary
    - Tab 2: B2B Invoices (Table 4 - GST Registered Wholesale Buyers)
    - Tab 3: B2C Invoices (Table 7 - Retail Consumers)
    - Tab 4: HSN Summary (Table 12 - HSN wise Quantity, Taxable Value, CGST, SGST, IGST)
    """
    cur = conn.cursor()
    cur.execute("SELECT * FROM settings ORDER BY id ASC LIMIT 1")
    settings = dict(cur.fetchone())
    shop_gstin = settings.get('tax_number', '').replace('GSTIN:', '').strip()
    shop_state = settings.get('state_name', 'Maharashtra')
    shop_state_code = settings.get('state_code', '27')

    query = """
        SELECT s.id, s.invoice_no, s.sale_type, s.customer_name, s.customer_phone,
               s.subtotal, s.discount, s.grand_total, s.taxable_amount, s.cgst_amount,
               s.sgst_amount, s.igst_amount, s.total_tax, s.is_interstate, s.created_at,
               COALESCE(NULLIF(s.customer_gst, ''), p.gst_no) as gst_no,
               p.state_code as buyer_state_code, p.address as buyer_address
        FROM sales s
        LEFT JOIN parties p ON s.party_id = p.id
        WHERE 1=1
    """
    params = []
    if start_date:
        query += " AND DATE(s.created_at) >= ?"
        params.append(start_date)
    if end_date:
        query += " AND DATE(s.created_at) <= ?"
        params.append(end_date)
    query += " ORDER BY s.id ASC"

    cur.execute(query, params)
    sales = [dict(r) for r in cur.fetchall()]

    wb = openpyxl.Workbook()

    # --- TAB 1: GSTR-1 SUMMARY ---
    ws_sum = wb.active
    ws_sum.title = "GSTR-1 Overview"
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum.append(["GOVERNMENT OF INDIA - GOODS AND SERVICES TAX"])
    ws_sum.append([f"FORM GSTR-1: Outward Supplies Return for {settings.get('shop_name', 'SSIPL')}"])
    ws_sum.append([f"GSTIN: {shop_gstin} | State: {shop_state} ({shop_state_code}) | Period: {start_date or 'Beginning'} to {end_date or 'Current'}"])
    ws_sum.append([])

    tot_b2b_inv = 0
    tot_b2b_taxable = 0.0
    tot_b2b_cgst = 0.0
    tot_b2b_sgst = 0.0
    tot_b2b_igst = 0.0
    tot_b2b_val = 0.0

    tot_b2c_inv = 0
    tot_b2c_taxable = 0.0
    tot_b2c_cgst = 0.0
    tot_b2c_sgst = 0.0
    tot_b2c_igst = 0.0
    tot_b2c_val = 0.0

    b2b_rows = []
    b2c_rows = []

    for s in sales:
        has_gstin = bool(s.get('gst_no') and len(s['gst_no'].strip()) >= 15)
        pos_code = s.get('buyer_state_code') or (s['gst_no'][:2] if has_gstin else shop_state_code)
        
        inv_val = float(s['grand_total'])
        taxable = float(s['taxable_amount'] or (inv_val / 1.18))
        cgst = float(s['cgst_amount'] or 0.0)
        sgst = float(s['sgst_amount'] or 0.0)
        igst = float(s['igst_amount'] or 0.0)

        if has_gstin:
            tot_b2b_inv += 1
            tot_b2b_taxable += taxable
            tot_b2b_cgst += cgst
            tot_b2b_sgst += sgst
            tot_b2b_igst += igst
            tot_b2b_val += inv_val

            b2b_rows.append([
                s['gst_no'], s['customer_name'], s['invoice_no'], s['created_at'][:10],
                inv_val, f"{pos_code}-{shop_state}", "N", "Regular", taxable, igst, cgst, sgst
            ])
        else:
            tot_b2c_inv += 1
            tot_b2c_taxable += taxable
            tot_b2c_cgst += cgst
            tot_b2c_sgst += sgst
            tot_b2c_igst += igst
            tot_b2c_val += inv_val

            b2c_rows.append([
                s['invoice_no'], s['created_at'][:10], s['customer_name'],
                f"{pos_code}-{shop_state}", taxable, 18.0, igst, cgst, sgst, inv_val
            ])

    ws_sum.append(["Table", "Category Description", "Total Invoices", "Taxable Value (₹)", "IGST (₹)", "CGST (₹)", "SGST (₹)", "Total Value (₹)"])
    ws_sum.append(["4A, 4B", "B2B Wholesale Invoices (GST Registered)", tot_b2b_inv, tot_b2b_taxable, tot_b2b_igst, tot_b2b_cgst, tot_b2b_sgst, tot_b2b_val])
    ws_sum.append(["7", "B2C Consumer & Retail Sales (Unregistered)", tot_b2c_inv, tot_b2c_taxable, tot_b2c_igst, tot_b2c_cgst, tot_b2c_sgst, tot_b2c_val])
    ws_sum.append(["TOTAL", "Total Outward Supplies (GSTR-1)", tot_b2b_inv + tot_b2c_inv,
                   tot_b2b_taxable + tot_b2c_taxable, tot_b2b_igst + tot_b2c_igst,
                   tot_b2b_cgst + tot_b2c_cgst, tot_b2b_sgst + tot_b2c_sgst, tot_b2b_val + tot_b2c_val])

    # Style overview
    ws_sum.cell(row=1, column=1).font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    for c in range(1, 9):
        hdr = ws_sum.cell(row=5, column=c)
        hdr.fill = HEADER_FILL_GST
        hdr.font = HEADER_FONT

    # --- TAB 2: B2B INVOICES ---
    ws_b2b = wb.create_sheet(title="4A_B2B_Invoices")
    ws_b2b.views.sheetView[0].showGridLines = True
    b2b_headers = ["GSTIN of Customer", "Customer / Firm Name", "Invoice No", "Invoice Date", "Invoice Value (₹)", "Place of Supply", "Reverse Charge", "Invoice Type", "Taxable Value (₹)", "Integrated Tax (₹)", "Central Tax (₹)", "State Tax (₹)"]
    ws_b2b.append(b2b_headers)
    for c in range(1, len(b2b_headers) + 1):
        cell = ws_b2b.cell(row=1, column=c)
        cell.fill = HEADER_FILL_PURCHASE
        cell.font = HEADER_FONT
    for r in b2b_rows:
        ws_b2b.append(r)

    # --- TAB 3: B2C INVOICES ---
    ws_b2c = wb.create_sheet(title="7_B2C_Retail_Sales")
    ws_b2c.views.sheetView[0].showGridLines = True
    b2c_headers = ["Invoice No", "Date", "Customer Name", "Place of Supply", "Taxable Value (₹)", "Rate (%)", "Integrated Tax (₹)", "Central Tax (₹)", "State Tax (₹)", "Total Value (₹)"]
    ws_b2c.append(b2c_headers)
    for c in range(1, len(b2c_headers) + 1):
        cell = ws_b2c.cell(row=1, column=c)
        cell.fill = HEADER_FILL_SALE
        cell.font = HEADER_FONT
    for r in b2c_rows:
        ws_b2c.append(r)

    # --- TAB 4: HSN SUMMARY (TABLE 12) ---
    ws_hsn = wb.create_sheet(title="12_HSN_Summary")
    ws_hsn.views.sheetView[0].showGridLines = True
    hsn_headers = ["HSN Code", "Description", "UQC / Unit", "Total Quantity", "Total Value (₹)", "Taxable Value (₹)", "Integrated Tax (₹)", "Central Tax (₹)", "State Tax (₹)", "Total Tax (₹)"]
    ws_hsn.append(hsn_headers)
    for c in range(1, len(hsn_headers) + 1):
        cell = ws_hsn.cell(row=1, column=c)
        cell.fill = HEADER_FILL_GST
        cell.font = HEADER_FONT

    # Query HSN aggregate across sales items
    hsn_query = """
        SELECT COALESCE(NULLIF(si.hsn_code, ''), p.hsn_code, '4202') as hsn,
               COALESCE(NULLIF(si.unit, ''), p.unit, 'PCS') as uqc,
               p.category,
               SUM(si.quantity) as tot_qty,
               SUM(si.total_price) as tot_val,
               SUM(COALESCE(NULLIF(si.taxable_amount, 0), si.total_price / 1.18)) as tot_taxable,
               SUM(si.cgst_amount) as tot_cgst,
               SUM(si.sgst_amount) as tot_sgst,
               SUM(si.igst_amount) as tot_igst,
               SUM(si.tax_amount) as tot_tax
        FROM sale_items si
        JOIN sales s ON si.sale_id = s.id
        LEFT JOIN products p ON si.product_id = p.id
        WHERE 1=1
    """
    hsn_params = []
    if start_date:
        hsn_query += " AND DATE(s.created_at) >= ?"
        hsn_params.append(start_date)
    if end_date:
        hsn_query += " AND DATE(s.created_at) <= ?"
        hsn_params.append(end_date)
    hsn_query += " GROUP BY hsn, uqc ORDER BY hsn ASC"

    cur.execute(hsn_query, hsn_params)
    hsn_rows = cur.fetchall()

    for hr in hsn_rows:
        desc = "Trunks, suit-cases, bags, handbags, wallets" if hr['hsn'] == '4202' else "Footwear, shoes, sneakers, sandals"
        ws_hsn.append([
            hr['hsn'], desc, hr['uqc'], hr['tot_qty'],
            round(hr['tot_val'], 2), round(hr['tot_taxable'], 2),
            round(hr['tot_igst'] or 0.0, 2), round(hr['tot_cgst'] or 0.0, 2),
            round(hr['tot_sgst'] or 0.0, 2), round(hr['tot_tax'] or 0.0, 2)
        ])

    # Auto-adjust column widths
    for sheet in [ws_sum, ws_b2b, ws_b2c, ws_hsn]:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"GSTR1_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return buf.getvalue(), filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def generate_gstr3b_summary_data(conn, start_date=None, end_date=None):
    """
    Computes monthly GSTR-3B tax figures:
    - Outward Supplies (Sales Output GST)
    - Eligible ITC (Purchases Input Tax Credit)
    - Net Tax Liability (Output Tax - Input Tax)
    """
    cur = conn.cursor()

    # 1. Output Supplies (Sales)
    s_query = "SELECT COUNT(*), COALESCE(SUM(taxable_amount), 0), COALESCE(SUM(cgst_amount), 0), COALESCE(SUM(sgst_amount), 0), COALESCE(SUM(igst_amount), 0), COALESCE(SUM(grand_total), 0) FROM sales WHERE 1=1"
    s_params = []
    if start_date:
        s_query += " AND DATE(created_at) >= ?"
        s_params.append(start_date)
    if end_date:
        s_query += " AND DATE(created_at) <= ?"
        s_params.append(end_date)
    
    cur.execute(s_query, s_params)
    s_res = cur.fetchone()

    sales_taxable = round(s_res[1] or 0.0, 2)
    sales_cgst = round(s_res[2] or 0.0, 2)
    sales_sgst = round(s_res[3] or 0.0, 2)
    sales_igst = round(s_res[4] or 0.0, 2)
    sales_total_tax = round(sales_cgst + sales_sgst + sales_igst, 2)
    sales_grand_total = round(s_res[5] or 0.0, 2)

    # If legacy sales had 0 tax populated, approximate base and 18% tax
    if sales_total_tax == 0 and sales_grand_total > 0:
        sales_taxable = round(sales_grand_total / 1.18, 2)
        tot_tax = round(sales_grand_total - sales_taxable, 2)
        sales_cgst = round(tot_tax / 2.0, 2)
        sales_sgst = round(tot_tax - sales_cgst, 2)
        sales_total_tax = tot_tax

    # 2. Inward Supplies (Purchases Input Tax Credit - ITC)
    p_query = "SELECT COUNT(*), COALESCE(SUM(taxable_amount), 0), COALESCE(SUM(cgst_amount), 0), COALESCE(SUM(sgst_amount), 0), COALESCE(SUM(igst_amount), 0), COALESCE(SUM(total_amount), 0) FROM purchases WHERE 1=1"
    p_params = []
    if start_date:
        p_query += " AND DATE(purchase_date) >= ?"
        p_params.append(start_date)
    if end_date:
        p_query += " AND DATE(purchase_date) <= ?"
        p_params.append(end_date)

    cur.execute(p_query, p_params)
    p_res = cur.fetchone()

    pur_taxable = round(p_res[1] or 0.0, 2)
    pur_cgst = round(p_res[2] or 0.0, 2)
    pur_sgst = round(p_res[3] or 0.0, 2)
    pur_igst = round(p_res[4] or 0.0, 2)
    pur_total_itc = round(pur_cgst + pur_sgst + pur_igst, 2)
    pur_total_val = round(p_res[5] or 0.0, 2)

    if pur_total_itc == 0 and pur_total_val > 0:
        pur_taxable = round(pur_total_val / 1.18, 2)
        tot_itc = round(pur_total_val - pur_taxable, 2)
        pur_cgst = round(tot_itc / 2.0, 2)
        pur_sgst = round(tot_itc - pur_cgst, 2)
        pur_total_itc = tot_itc

    # 3. Net Tax Payable = Output Tax - Input Tax Credit
    net_cgst = max(0.0, round(sales_cgst - pur_cgst, 2))
    net_sgst = max(0.0, round(sales_sgst - pur_sgst, 2))
    net_igst = max(0.0, round(sales_igst - pur_igst, 2))
    net_total_payable = round(net_cgst + net_sgst + net_igst, 2)

    # 4. HSN Summary for Table 12
    hsn_query = """
        SELECT COALESCE(NULLIF(si.hsn_code, ''), CASE WHEN lower(si.product_name) LIKE '%bag%' OR lower(si.product_name) LIKE '%wallet%' THEN '4202' ELSE '6404' END) as hsn,
               COALESCE(NULLIF(si.unit, ''), 'Pairs') as uqc,
               SUM(si.quantity) as tot_qty,
               SUM(si.total_price) as tot_val,
               SUM(COALESCE(NULLIF(si.taxable_amount, 0), ROUND(si.total_price / 1.18, 2))) as tot_taxable,
               SUM(COALESCE(NULLIF(si.cgst_amount, 0), ROUND((si.total_price - (si.total_price / 1.18)) / 2.0, 2))) as tot_cgst,
               SUM(COALESCE(NULLIF(si.sgst_amount, 0), ROUND((si.total_price - (si.total_price / 1.18)) / 2.0, 2))) as tot_sgst,
               SUM(COALESCE(si.igst_amount, 0)) as tot_igst,
               SUM(COALESCE(NULLIF(si.tax_amount, 0), ROUND(si.total_price - (si.total_price / 1.18), 2))) as tot_tax
        FROM sale_items si
        JOIN sales s ON s.id = si.sale_id
        WHERE 1=1
    """
    hsn_params = []
    if start_date:
        hsn_query += " AND DATE(s.created_at) >= ?"
        hsn_params.append(start_date)
    if end_date:
        hsn_query += " AND DATE(s.created_at) <= ?"
        hsn_params.append(end_date)
    hsn_query += " GROUP BY hsn, uqc ORDER BY tot_val DESC"

    cur.execute(hsn_query, hsn_params)
    hsn_rows = [dict(r) for r in cur.fetchall()]

    return {
        'sales_count': s_res[0] or 0,
        'taxable_sales': sales_taxable,
        'cgst_output': sales_cgst,
        'sgst_output': sales_sgst,
        'igst_output': sales_igst,
        'total_output_tax': sales_total_tax,
        'gross_sales': sales_grand_total,

        'purchases_count': p_res[0] or 0,
        'taxable_purchases': pur_taxable,
        'cgst_itc': pur_cgst,
        'sgst_itc': pur_sgst,
        'igst_itc': pur_igst,
        'total_itc': pur_total_itc,
        'gross_purchases': pur_total_val,

        'net_cgst_payable': net_cgst,
        'net_sgst_payable': net_sgst,
        'net_igst_payable': net_igst,
        'net_tax_payable': net_total_payable,

        'hsn_rows': hsn_rows,

        'sales': {
            'count': s_res[0],
            'taxable': sales_taxable,
            'cgst': sales_cgst,
            'sgst': sales_sgst,
            'igst': sales_igst,
            'total_tax': sales_total_tax,
            'grand_total': sales_grand_total
        },
        'purchases': {
            'count': p_res[0],
            'taxable': pur_taxable,
            'cgst': pur_cgst,
            'sgst': pur_sgst,
            'igst': pur_igst,
            'total_itc': pur_total_itc,
            'total_amount': pur_total_val
        },
        'net_payable': {
            'cgst': net_cgst,
            'sgst': net_sgst,
            'igst': net_igst,
            'total': net_total_payable
        }
    }
