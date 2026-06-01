from flask import Flask, request, jsonify, Response
from flask_cors import CORS
import json, os
from datetime import datetime
from google.oauth2.service_account import Credentials
import gspread

app = Flask(__name__)
CORS(app)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

SHEET_ID   = os.environ.get("GOOGLE_SHEET_ID", "")
CREDS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")

FIXED_COLS = ["Date", "Slot", "Account", "Order Name", "Status"]
STATUS_COL = 5   # 1-based, column E

def get_client():
    raw = CREDS_JSON.strip()
    if (raw.startswith("'") and raw.endswith("'")) or \
       (raw.startswith('"') and raw.endswith('"')):
        raw = raw[1:-1]
    data  = json.loads(raw)
    creds = Credentials.from_service_account_info(data, scopes=SCOPES)
    return gspread.authorize(creds)

def get_sheet():
    return get_client().open_by_key(SHEET_ID.strip())

def inv_title():
    return f"Inventory_{datetime.now().strftime('%b_%Y')}"

def get_inv_ws(sh):
    return sh.worksheet(inv_title())

def get_products_from_sheet(sh):
    try:
        vals = sh.worksheet("Product_Master").col_values(1)
        return [v.strip() for v in vals[1:] if v.strip()]
    except:
        return []

def get_slots_from_sheet(sh):
    try:
        vals = sh.worksheet("Slot_Master").col_values(1)
        return [v.strip() for v in vals[1:] if v.strip()]
    except:
        return ["10 PM", "12 PM", "8 AM", "2 PM", "4 PM"]

# ── /api/health ───────────────────────────────────────────────────────────────
@app.route("/api/health")
def health():
    if not CREDS_JSON:
        return jsonify({"ok": False, "error": "GOOGLE_CREDENTIALS_JSON env var not set"}), 500
    if not SHEET_ID:
        return jsonify({"ok": False, "error": "GOOGLE_SHEET_ID env var not set"}), 500
    try:
        sh     = get_sheet()
        titles = [ws.title for ws in sh.worksheets()]
        return jsonify({"ok": True, "sheet_name": sh.title, "tabs": titles})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

# ── /api/setup ────────────────────────────────────────────────────────────────
@app.route("/api/setup", methods=["POST"])
def setup_db():
    try:
        body     = request.json or {}
        products = [p.strip() for p in body.get("products", []) if p.strip()]
        slots    = [s.strip() for s in body.get("slots",    []) if s.strip()]
        if not products:
            return jsonify({"error": "No products provided"}), 400

        sh       = get_sheet()
        existing = {ws.title: ws for ws in sh.worksheets()}

        def get_or_create(title, rows=500, cols=60):
            if title in existing:
                ws = existing[title]; ws.clear(); return ws
            return sh.add_worksheet(title=title, rows=str(rows), cols=str(cols))

        ws_prod = get_or_create("Product_Master", rows=200, cols=3)
        ws_prod.update("A1", [["Product Name"]])
        ws_prod.update(f"A2:A{len(products)+1}", [[p] for p in products])

        if not slots:
            slots = ["10 PM", "12 PM", "8 AM", "2 PM", "4 PM"]
        ws_slot = get_or_create("Slot_Master", rows=20, cols=2)
        ws_slot.update("A1", [["Slots"]])
        ws_slot.update(f"A2:A{len(slots)+1}", [[s] for s in slots])

        ws_sales = get_or_create("Sales_Log", rows=1000, cols=5)
        ws_sales.update("A1", [["Date", "Buyer Name", "Product Name", "Quantity Sold"]])

        month_key  = datetime.now().strftime("%b_%Y")
        sheet_name = f"Inventory_{month_key}"
        headers    = FIXED_COLS + products

        ws_inv = get_or_create(sheet_name, rows=500, cols=len(headers)+5)
        ws_inv.update("A1", [headers])
        ws_inv.update("A2", [["-", "-", "Old Stock", "-", "Delivered"] + ["0"]*len(products)])
        try:
            ws_inv.format("A1:ZZ1", {
                "textFormat":      {"bold": True},
                "backgroundColor": {"red": 0.13, "green": 0.37, "blue": 0.18}
            })
        except: pass

        try: sh.del_worksheet(sh.worksheet("Sheet1"))
        except: pass

        return jsonify({
            "success": True, "month": month_key,
            "inv_sheet": sheet_name, "product_count": len(products),
            "tabs_created": ["Product_Master","Slot_Master","Sales_Log", sheet_name]
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/init-data ────────────────────────────────────────────────────────────
@app.route("/api/init-data")
def init_data():
    try:
        sh       = get_sheet()
        products = get_products_from_sheet(sh)
        slots    = get_slots_from_sheet(sh)
        today    = datetime.now().strftime("%Y-%m-%d")
        month    = datetime.now().strftime("%b_%Y")
        titles   = [ws.title for ws in sh.worksheets()]
        inv_ok   = f"Inventory_{month}" in titles
        return jsonify({
            "products": products,
            "slots":    slots,
            "today":    today,
            "month":    month,
            "inv_ok":   inv_ok
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/add-order ────────────────────────────────────────────────────────────
@app.route("/api/add-order", methods=["POST"])
def add_order():
    try:
        body       = request.json or {}
        date       = body.get("date",       datetime.now().strftime("%Y-%m-%d"))
        slot       = body.get("slot",       "")
        account    = body.get("account",    "")
        order_name = body.get("order_name", "")
        quantities = body.get("quantities", {})

        sh       = get_sheet()
        ws       = get_inv_ws(sh)
        headers  = ws.row_values(1)
        products = headers[len(FIXED_COLS):]

        row = [date, slot, account, order_name, "Pending"]
        for p in products:
            row.append(str(quantities.get(p, "") or ""))

        ws.append_row(row, value_input_option="USER_ENTERED")
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/receiver ─────────────────────────────────────────────────────────────
@app.route("/api/receiver")
def receiver():
    try:
        date_filter = request.args.get("date", datetime.now().strftime("%Y-%m-%d"))
        sh      = get_sheet()
        ws      = get_inv_ws(sh)
        data    = ws.get_all_values()
        if not data:
            return jsonify({"slots": [], "summary": {}, "orders": [], "date": date_filter})

        headers  = data[0]
        products = headers[len(FIXED_COLS):]
        orders       = []
        slot_summary = {}

        for i, row in enumerate(data[1:], start=2):
            if not any(c.strip() for c in row):
                continue
            padded     = row + [""] * max(0, len(headers) - len(row))
            row_date   = padded[0].strip()
            row_slot   = padded[1].strip()
            row_acct   = padded[2].strip()
            row_oname  = padded[3].strip()
            row_status = padded[4].strip().lower()

            if row_acct.lower() in ("old stock", "current stock", ""):
                continue
            if row_date != date_filter or row_status != "pending":
                continue

            qtys = {}
            for j, p in enumerate(products):
                col = len(FIXED_COLS) + j
                val = padded[col] if col < len(padded) else ""
                try:    qtys[p] = int(float(val)) if val else 0
                except: qtys[p] = 0

            orders.append({"row_index": i, "date": row_date, "slot": row_slot,
                           "account": row_acct, "order_name": row_oname, "quantities": qtys})
            if row_slot not in slot_summary:
                slot_summary[row_slot] = {p: 0 for p in products}
            for p, q in qtys.items():
                slot_summary[row_slot][p] = slot_summary[row_slot].get(p, 0) + q

        return jsonify({
            "products": products, "slots": list(slot_summary.keys()),
            "summary": slot_summary, "orders": orders, "date": date_filter
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/mark-delivered ───────────────────────────────────────────────────────
@app.route("/api/mark-delivered", methods=["POST"])
def mark_delivered():
    try:
        body      = request.json or {}
        row_index = body.get("row_index")
        month_key = body.get("month_key", None)
        sh  = get_sheet()
        ws  = sh.worksheet(f"Inventory_{month_key}") if month_key else get_inv_ws(sh)
        ws.update_cell(row_index, STATUS_COL, "Delivered")
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/months ───────────────────────────────────────────────────────────────
@app.route("/api/months")
def list_months():
    try:
        sh     = get_sheet()
        months = [ws.title[len("Inventory_"):] for ws in sh.worksheets()
                  if ws.title.startswith("Inventory_")]
        return jsonify({"months": months})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/month-dashboard/<month_key> ─────────────────────────────────────────
@app.route("/api/month-dashboard/<month_key>")
def month_dashboard(month_key):
    try:
        sh   = get_sheet()
        ws   = sh.worksheet(f"Inventory_{month_key}")
        data = ws.get_all_values()
        if not data:
            return jsonify({"error": "Sheet is empty"}), 404

        headers  = data[0]
        products = headers[len(FIXED_COLS):]

        total_orders = delivered = pending = 0
        prod_totals  = {p: 0 for p in products}
        date_counts  = {}
        slot_counts  = {}
        orders       = []

        for i, row in enumerate(data[1:], start=2):
            if not any(c.strip() for c in row):
                continue
            padded = row + [""] * max(0, len(headers) - len(row))
            acct   = padded[2].strip()
            status = padded[4].strip()
            if acct.lower() in ("old stock", "current stock", ""):
                continue

            date  = padded[0].strip()
            slot  = padded[1].strip()
            oname = padded[3].strip()

            total_orders += 1
            if status.lower() == "delivered":
                delivered += 1
            else:
                pending += 1

            date_counts[date] = date_counts.get(date, 0) + 1
            if slot:
                slot_counts[slot] = slot_counts.get(slot, 0) + 1

            qtys = {}
            for j, p in enumerate(products):
                col = len(FIXED_COLS) + j
                val = padded[col] if col < len(padded) else ""
                try:    n = int(float(val)) if val else 0
                except: n = 0
                qtys[p] = n
                if status.lower() == "delivered":
                    prod_totals[p] += n

            orders.append({
                "row_index": i, "date": date, "slot": slot,
                "account": acct, "order_name": oname,
                "status": status, "quantities": qtys
            })

        return jsonify({
            "month_key": month_key, "products": products,
            "total_orders": total_orders, "delivered": delivered, "pending": pending,
            "prod_totals": prod_totals,
            "date_counts": dict(sorted(date_counts.items())),
            "slot_counts": slot_counts, "orders": orders
        })
    except gspread.WorksheetNotFound:
        return jsonify({"error": f"Inventory_{month_key} not found"}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/sales-manager/<month_key> ───────────────────────────────────────────
@app.route("/api/sales-manager/<month_key>")
def sales_manager(month_key):
    try:
        sh       = get_sheet()
        products = get_products_from_sheet(sh)
        SALESMEN = ["Gourab", "Souvik", "Rajkumar"]

        try:
            ws_inv   = sh.worksheet(f"Inventory_{month_key}")
            inv_data = ws_inv.get_all_values()
            headers  = inv_data[0] if inv_data else []
            prod_cols = {p: headers.index(p) for p in products if p in headers}
            delivered_totals = {p: 0 for p in products}
            for row in inv_data[1:]:
                if not row or len(row) < 5: continue
                acct   = row[2].strip()
                status = row[4].strip().lower()
                is_old_stock = acct.lower() == "old stock"
                # Include Old Stock row (transferred from previous month)
                # AND all rows with status = delivered
                if is_old_stock:
                    for p, ci in prod_cols.items():
                        try: delivered_totals[p] += int(float(row[ci])) if ci < len(row) and row[ci] else 0
                        except: pass
                elif acct.lower() in ("current stock", "") or status != "delivered":
                    continue
                else:
                    for p, ci in prod_cols.items():
                        try: delivered_totals[p] += int(float(row[ci])) if ci < len(row) and row[ci] else 0
                        except: pass
        except:
            delivered_totals = {p: 0 for p in products}

        sales = {s: {p: 0 for p in products} for s in SALESMEN}
        try:
            import calendar
            ws_sales   = sh.worksheet("Sales_Log")
            sales_data = ws_sales.get_all_values()
            parts    = month_key.split("_")
            mon_abbr = parts[0]; year = parts[1]
            mon_num  = list(calendar.month_abbr).index(mon_abbr)
            prefix   = f"{year}-{mon_num:02d}"
            for row in sales_data[1:]:
                if len(row) < 4: continue
                date, buyer, product, qty = row[0].strip(), row[1].strip(), row[2].strip(), row[3].strip()
                if not date.startswith(prefix): continue
                matched = next((s for s in SALESMEN if s.lower() in buyer.lower()), None)
                if matched and product in sales[matched]:
                    try: sales[matched][product] += int(float(qty)) if qty else 0
                    except: pass
        except: pass

        remaining = {}
        for p in products:
            sold = sum(sales[s][p] for s in SALESMEN)
            remaining[p] = max(0, delivered_totals[p] - sold)

        return jsonify({
            "month_key": month_key, "products": products, "salesmen": SALESMEN,
            "delivered": delivered_totals, "sales": sales, "remaining": remaining
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/log-sale ─────────────────────────────────────────────────────────────
@app.route("/api/log-sale", methods=["POST"])
def log_sale():
    """Single sale entry — append only (used for quick logging)."""
    try:
        body    = request.json or {}
        date    = body.get("date",    datetime.now().strftime("%Y-%m-%d"))
        buyer   = body.get("buyer",   "")
        product = body.get("product", "")
        qty     = int(body.get("qty", 0))
        if not buyer or not product or qty <= 0:
            return jsonify({"error": "buyer, product and qty required"}), 400
        sh = get_sheet()
        ws = sh.worksheet("Sales_Log")
        ws.append_row([date, buyer, product, qty], value_input_option="USER_ENTERED")
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/save-sales  (batch save — clears old entries for month+buyer first) ──
@app.route("/api/save-sales", methods=["POST"])
def save_sales():
    """
    Receives full sales data for a month and replaces all existing Sales_Log
    entries for that month+salesman combination with fresh data.
    Body: { month_key: "May_2026", sales: { "Gourab": { "Product": qty, ... }, ... } }
    """
    try:
        import calendar
        body      = request.json or {}
        month_key = body.get("month_key", "")
        sales_in  = body.get("sales", {})   # { salesman: { product: qty } }

        if not month_key:
            return jsonify({"error": "month_key required"}), 400

        # Build month date prefix e.g. "2026-05"
        parts    = month_key.split("_")
        mon_abbr = parts[0]; year = parts[1]
        mon_num  = list(calendar.month_abbr).index(mon_abbr)
        prefix   = f"{year}-{mon_num:02d}"
        save_date = f"{prefix}-01"   # use 1st of month as canonical date

        sh = get_sheet()
        ws = sh.worksheet("Sales_Log")
        all_rows = ws.get_all_values()

        # Find row indices to DELETE (existing entries for this month)
        rows_to_delete = []
        for i, row in enumerate(all_rows[1:], start=2):
            if not row: continue
            row_date  = row[0].strip() if len(row) > 0 else ""
            row_buyer = row[1].strip() if len(row) > 1 else ""
            if row_date.startswith(prefix) and row_buyer in sales_in:
                rows_to_delete.append(i)

        # Delete from bottom up to preserve row indices
        for idx in reversed(rows_to_delete):
            ws.delete_rows(idx)

        # Append fresh rows for every salesman + product with qty > 0
        new_rows = []
        for salesman, products in sales_in.items():
            for product, qty in products.items():
                if qty and int(qty) > 0:
                    new_rows.append([save_date, salesman, product, int(qty)])

        if new_rows:
            ws.append_rows(new_rows, value_input_option="USER_ENTERED")

        return jsonify({"success": True, "saved": len(new_rows), "deleted": len(rows_to_delete)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/transfer-stock ───────────────────────────────────────────────────────
@app.route("/api/transfer-stock", methods=["POST"])
def transfer_stock():
    try:
        import calendar
        body      = request.json or {}
        from_key  = body.get("from_month")
        remaining = body.get("remaining", {})
        if not from_key:
            return jsonify({"error": "from_month required"}), 400

        sh       = get_sheet()
        products = get_products_from_sheet(sh)
        parts    = from_key.split("_")
        mon_abbr = parts[0]; year = int(parts[1])
        mon_num  = list(calendar.month_abbr).index(mon_abbr)
        if mon_num == 12:
            next_mon_num = 1; next_year = year + 1
        else:
            next_mon_num = mon_num + 1; next_year = year
        next_key   = f"{calendar.month_abbr[next_mon_num]}_{next_year}"
        next_title = f"Inventory_{next_key}"

        existing = {ws.title for ws in sh.worksheets()}
        headers  = FIXED_COLS + products
        cols     = len(headers) + 5

        if next_title not in existing:
            ws_next = sh.add_worksheet(title=next_title, rows=500, cols=cols)
        else:
            ws_next = sh.worksheet(next_title)

        ws_next.update("A1", [headers])
        old_row = ["-", "-", "Old Stock", "-", "Delivered"] + \
                  [str(remaining.get(p, 0)) for p in products]
        ws_next.update("A2", [old_row])
        try:
            ws_next.format("A1:ZZ1", {
                "textFormat":      {"bold": True},
                "backgroundColor": {"red": 0.13, "green": 0.37, "blue": 0.18}
            })
        except: pass

        return jsonify({"success": True, "next_month": next_key, "sheet": next_title})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── frontend ──────────────────────────────────────────────────────────────────
TEMPLATE = os.path.join(os.path.dirname(__file__), "templates", "index.html")

@app.route("/")
def index():
    with open(TEMPLATE, encoding="utf-8") as f:
        return Response(f.read(), mimetype="text/html")

if __name__ == "__main__":
    app.run(debug=True)
