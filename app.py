from flask import Flask, request, jsonify, Response
from flask_cors import CORS
import json, os, traceback
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

# ── Google Sheets client ──────────────────────────────────────────────────────

def get_sheet():
    if not CREDS_JSON:
        raise RuntimeError("GOOGLE_CREDENTIALS_JSON env var not set")
    if not SHEET_ID:
        raise RuntimeError("GOOGLE_SHEET_ID env var not set")
    data  = json.loads(CREDS_JSON)
    creds = Credentials.from_service_account_info(data, scopes=SCOPES)
    gc    = gspread.authorize(creds)
    return gc.open_by_key(SHEET_ID)

def ws_or_create(sh, title, rows=500, cols=60):
    try:
        return sh.worksheet(title)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=title, rows=rows, cols=cols)

# ── Sheet-name helpers ────────────────────────────────────────────────────────

def inv_title(month_key):
    """month_key like 'May_2026'  →  'Inventory_May_2026'"""
    return f"Inventory_{month_key}"

def strip_inv(title):
    """'Inventory_May_2026' → 'May_2026' (returns None for non-inventory sheets)"""
    if title.startswith("Inventory_"):
        return title[len("Inventory_"):]
    return None

FIXED_SHEETS = {"Product_Master", "Slot_Master", "Sales_Log"}

# ── Inventory sheet structure ─────────────────────────────────────────────────
# Row 1  : headers  → Date | Slot | Account | Order Name | Status | <products…>
# Row 2  : Old Stock row
# Row 3+ : order rows  (Status = "Pending" or "Delivered")

FIXED_COLS = ["Date", "Slot", "Account", "Order Name", "Status"]
STATUS_COL = 5   # 1-based column index of Status

def get_products_list(sh):
    try:
        vals = sh.worksheet("Product_Master").col_values(1)
        return [v for v in vals[1:] if v.strip()]
    except:
        return []

def get_slots_list(sh):
    try:
        vals = sh.worksheet("Slot_Master").col_values(1)
        return [v for v in vals[1:] if v.strip()]
    except:
        return ["10 PM", "12 PM", "8 AM"]

# ── /api/health ──────────────────────────────────────────────────────────────

@app.route("/api/health")
def health():
    try:
        sh = get_sheet()
        titles = [ws.title for ws in sh.worksheets()]
        return jsonify({"ok": True, "sheets": titles})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

# ── /api/products ─────────────────────────────────────────────────────────────

@app.route("/api/products", methods=["GET"])
def get_products():
    try:
        sh = get_sheet()
        ws = ws_or_create(sh, "Product_Master", rows=100, cols=2)
        rows = ws.get_all_values()
        products = [r[0] for r in rows if r and r[0].strip() and r[0] != "Product Name"]
        return jsonify({"products": products})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/products", methods=["POST"])
def save_products():
    try:
        products = request.json.get("products", [])
        sh = get_sheet()
        ws = ws_or_create(sh, "Product_Master", rows=100, cols=2)
        ws.clear()
        ws.update("A1", [["Product Name"]] + [[p] for p in products])
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/slots ────────────────────────────────────────────────────────────────

@app.route("/api/slots", methods=["GET"])
def get_slots():
    try:
        sh = get_sheet()
        return jsonify({"slots": get_slots_list(sh)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/months ───────────────────────────────────────────────────────────────

@app.route("/api/months", methods=["GET"])
def list_months():
    try:
        sh = get_sheet()
        keys = [strip_inv(ws.title) for ws in sh.worksheets() if strip_inv(ws.title)]
        return jsonify({"months": keys})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/months", methods=["POST"])
def create_month():
    try:
        month_key = request.json.get("month")          # e.g. "May_2026"
        sh        = get_sheet()
        title     = inv_title(month_key)

        existing = [ws.title for ws in sh.worksheets()]
        if title in existing:
            return jsonify({"error": "Month already exists"}), 400

        products = get_products_list(sh)
        headers  = FIXED_COLS + products

        ws = sh.add_worksheet(title=title, rows=500, cols=len(headers) + 5)
        ws.update("A1", [headers])

        # Old Stock row — carry forward Current Stock from previous month if available
        old_stock_qtys = [""] * len(products)
        inv_sheets = [t for t in existing if t.startswith("Inventory_")]
        if inv_sheets:
            prev_ws   = sh.worksheet(inv_sheets[-1])
            prev_data = prev_ws.get_all_values()
            for row in prev_data:
                if row and row[0].strip() == "Current Stock":
                    old_stock_qtys = row[len(FIXED_COLS):]
                    break

        old_stock_row = ["-", "-", "Old Stock", "-", "Delivered"] + old_stock_qtys
        ws.update("A2", [old_stock_row])

        # Bold + green header
        ws.format("A1:ZZ1", {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.13, "green": 0.37, "blue": 0.18}
        })
        return jsonify({"success": True, "month": month_key})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/sync  (one-click DB setup like Streamlit version) ───────────────────

@app.route("/api/sync", methods=["POST"])
def sync_db():
    """Ensure all required sheets exist with correct headers for current month."""
    try:
        sh = get_sheet()
        existing = {ws.title for ws in sh.worksheets()}

        # Product_Master
        if "Product_Master" not in existing:
            ws = sh.add_worksheet(title="Product_Master", rows=100, cols=2)
            ws.update("A1", [["Product Name"]])

        # Slot_Master
        if "Slot_Master" not in existing:
            ws = sh.add_worksheet(title="Slot_Master", rows=20, cols=2)
            ws.update("A1", [["Slots"], ["10 PM"], ["12 PM"], ["8 AM"], ["2 PM"], ["4 PM"]])

        # Sales_Log
        if "Sales_Log" not in existing:
            ws = sh.add_worksheet(title="Sales_Log", rows=1000, cols=5)
            ws.update("A1", [["Date", "Buyer Name", "Product Name", "Quantity Sold"]])

        # Current month inventory sheet
        month_key = datetime.now().strftime("%b_%Y")   # e.g. "May_2026"
        title     = inv_title(month_key)
        products  = get_products_list(sh)
        headers   = FIXED_COLS + products

        if title not in existing:
            ws = sh.add_worksheet(title=title, rows=500, cols=len(headers) + 5)
            ws.update("A1", [headers])
            ws.update("A2", [["-", "-", "Old Stock", "-", "Delivered"] + ["0"] * len(products)])
            ws.format("A1:ZZ1", {
                "textFormat": {"bold": True},
                "backgroundColor": {"red": 0.13, "green": 0.37, "blue": 0.18}
            })
        else:
            ws = sh.worksheet(title)
            ws.update("A1", [headers])

        return jsonify({"success": True, "month": month_key})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/orders/<month_key> ───────────────────────────────────────────────────

@app.route("/api/orders/<month_key>", methods=["GET"])
def get_orders(month_key):
    try:
        sh   = get_sheet()
        ws   = sh.worksheet(inv_title(month_key))
        data = ws.get_all_values()
        if not data:
            return jsonify({"headers": [], "rows": []})

        headers = data[0]
        rows    = []
        for i, row in enumerate(data[1:], start=2):
            if not any(c.strip() for c in row):
                continue
            # Pad row to header length
            padded = row + [""] * max(0, len(headers) - len(row))
            rows.append({
                "row_index": i,
                "date":       padded[0] if len(padded) > 0 else "",
                "slot":       padded[1] if len(padded) > 1 else "",
                "account":    padded[2] if len(padded) > 2 else "",
                "order_name": padded[3] if len(padded) > 3 else "",
                "status":     padded[4] if len(padded) > 4 else "",
                "quantities": {
                    headers[j]: padded[j]
                    for j in range(len(FIXED_COLS), len(headers))
                    if j < len(padded)
                }
            })
        return jsonify({"headers": headers, "rows": rows, "products": headers[len(FIXED_COLS):]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/orders/<month_key>", methods=["POST"])
def add_order(month_key):
    try:
        body       = request.json
        date       = body.get("date", str(datetime.now().date()))
        slot       = body.get("slot", "")
        account    = body.get("account", "")
        order_name = body.get("order_name", "")
        quantities = body.get("quantities", {})

        sh      = get_sheet()
        ws      = sh.worksheet(inv_title(month_key))
        headers = ws.row_values(1)
        products = headers[len(FIXED_COLS):]

        row = [date, slot, account, order_name, "Pending"]
        row += [str(quantities.get(p, "")) for p in products]

        ws.append_row(row, value_input_option="USER_ENTERED")
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/orders/<month_key>/verify ────────────────────────────────────────────

@app.route("/api/orders/<month_key>/verify", methods=["POST"])
def verify_order(month_key):
    try:
        row_index = request.json.get("row_index")
        sh = get_sheet()
        ws = sh.worksheet(inv_title(month_key))
        ws.update_cell(row_index, STATUS_COL, "Delivered")
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/orders/<month_key>/row/<n>  DELETE ───────────────────────────────────

@app.route("/api/orders/<month_key>/row/<int:row_index>", methods=["DELETE"])
def delete_order_row(month_key, row_index):
    try:
        sh = get_sheet()
        ws = sh.worksheet(inv_title(month_key))
        ws.delete_rows(row_index)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/receiver/<month_key>  (slot summary + pending orders) ────────────────

@app.route("/api/debug/<month_key>", methods=["GET"])
def debug_sheet(month_key):
    """Shows raw sheet data for debugging."""
    try:
        sh   = get_sheet()
        ws   = sh.worksheet(inv_title(month_key))
        data = ws.get_all_values()
        headers = data[0] if data else []
        return jsonify({
            "headers": headers,
            "header_count": len(headers),
            "row_count": len(data),
            "first_3_rows": data[1:4] if len(data) > 1 else [],
            "products_detected": headers[5:] if len(headers) > 5 else [],
            "fixed_cols_used": 5
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/receiver/<month_key>", methods=["GET"])
def receiver_view(month_key):
    """Returns ALL orders (pending + delivered) with slot summary."""
    try:
        sh   = get_sheet()
        ws   = sh.worksheet(inv_title(month_key))
        data = ws.get_all_values()
        if not data:
            return jsonify({"products": [], "all_orders": [], "pending_orders": [], "slot_totals": {}})

        headers = data[0]

        # Dynamically find column indices from actual header row
        def col_idx(name):
            """Find 0-based index of a column by name (case-insensitive)."""
            for i, h in enumerate(headers):
                if h.strip().lower() == name.lower():
                    return i
            return None

        date_col   = col_idx("Date")   if col_idx("Date")   is not None else 0
        slot_col   = col_idx("Slot")   if col_idx("Slot")   is not None else 1
        acct_col   = col_idx("Account") if col_idx("Account") is not None else 2
        oname_col  = col_idx("Order Name") if col_idx("Order Name") is not None else 3
        status_col = col_idx("Status") if col_idx("Status") is not None else 4

        # Products start after Status column (whichever is last of fixed cols)
        prod_start = max(date_col, slot_col, acct_col, oname_col, status_col) + 1
        products   = [h.strip() for h in headers[prod_start:] if h.strip()]

        all_orders    = []
        pending_orders = []
        # Overall totals (all pending)
        overall_totals = {p: 0 for p in products}
        # Slot totals (pending only)
        slot_totals    = {}

        for i, row in enumerate(data[1:], start=2):
            if not any(c.strip() for c in row):
                continue
            padded = row + [""] * max(0, len(headers) - len(row))

            acct = padded[acct_col].strip() if acct_col < len(padded) else ""
            # Skip structural rows
            if acct.lower() in ("old stock", "current stock", ""):
                continue

            date       = padded[date_col].strip()   if date_col < len(padded)  else ""
            slot       = padded[slot_col].strip()   if slot_col < len(padded)  else ""
            order_name = padded[oname_col].strip()  if oname_col < len(padded) else ""
            status     = padded[status_col].strip() if status_col < len(padded) else "Pending"

            qtys = {}
            for j, p in enumerate(products):
                col = prod_start + j
                val = padded[col].strip() if col < len(padded) else ""
                try:
                    qtys[p] = int(float(val)) if val else 0
                except:
                    qtys[p] = 0

            order = {
                "row_index":  i,
                "date":       date,
                "slot":       slot,
                "account":    acct,
                "order_name": order_name,
                "status":     status,
                "quantities": qtys
            }
            all_orders.append(order)

            if status.lower() == "pending":
                pending_orders.append(order)
                # Slot totals
                if slot not in slot_totals:
                    slot_totals[slot] = {p: 0 for p in products}
                for p, q in qtys.items():
                    slot_totals[slot][p] = slot_totals[slot].get(p, 0) + q
                # Overall totals
                for p, q in qtys.items():
                    overall_totals[p] = overall_totals.get(p, 0) + q

        return jsonify({
            "products":       products,
            "all_orders":     all_orders,
            "pending_orders": pending_orders,
            "slot_totals":    slot_totals,
            "overall_totals": overall_totals,
            "pending_count":  len(pending_orders),
            "total_count":    len(all_orders)
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/sales/<month_key> ────────────────────────────────────────────────────

@app.route("/api/sales/<month_key>", methods=["POST"])
def log_sale(month_key):
    """Append a single sale entry to Sales_Log."""
    try:
        body    = request.json
        buyer   = body.get("buyer", "")
        product = body.get("product", "")
        qty     = body.get("qty", 0)
        date    = body.get("date", str(datetime.now().date()))

        sh = get_sheet()
        ws = ws_or_create(sh, "Sales_Log", rows=1000, cols=5)
        ws.append_row([date, buyer, product, qty], value_input_option="USER_ENTERED")
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/dashboard/<month_key> ───────────────────────────────────────────────

@app.route("/api/dashboard/<month_key>", methods=["GET"])
def dashboard(month_key):
    try:
        sh   = get_sheet()
        ws   = sh.worksheet(inv_title(month_key))
        data = ws.get_all_values()
        if not data:
            return jsonify({"products": [], "delivered": {}, "pending": {}, "order_count": 0, "pending_count": 0, "delivered_count": 0})

        headers  = data[0]
        products = headers[len(FIXED_COLS):]

        delivered_qty = {p: 0 for p in products}
        pending_qty   = {p: 0 for p in products}
        order_count = delivered_count = pending_count = 0

        for row in data[1:]:
            if not any(c.strip() for c in row):
                continue
            padded = row + [""] * max(0, len(headers) - len(row))
            acct   = padded[2].strip()
            status = padded[4].strip().lower() if len(padded) > 4 else ""

            if acct in ("Old Stock", "Current Stock", ""):
                continue
            order_count += 1

            for j, p in enumerate(products):
                col = len(FIXED_COLS) + j
                val = padded[col] if col < len(padded) else ""
                try:
                    n = int(val) if val else 0
                except:
                    n = 0
                if status == "delivered":
                    delivered_qty[p] += n
                else:
                    pending_qty[p] += n

            if status == "delivered":
                delivered_count += 1
            else:
                pending_count += 1

        # Current stock from Old Stock row + all delivered
        old_stock = {p: 0 for p in products}
        for row in data[1:]:
            if row and row[2].strip() == "Old Stock":
                padded = row + [""] * max(0, len(headers) - len(row))
                for j, p in enumerate(products):
                    col = len(FIXED_COLS) + j
                    try:
                        old_stock[p] = int(padded[col]) if col < len(padded) and padded[col] else 0
                    except:
                        old_stock[p] = 0
                break

        current_stock = {p: old_stock[p] + delivered_qty[p] for p in products}

        return jsonify({
            "products":       products,
            "current_stock":  current_stock,
            "delivered":      delivered_qty,
            "pending":        pending_qty,
            "order_count":    order_count,
            "pending_count":  pending_count,
            "delivered_count": delivered_count
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── Frontend ──────────────────────────────────────────────────────────────────

with open(os.path.join(os.path.dirname(__file__), "templates", "index.html"), encoding="utf-8") as _f:
    _HTML = _f.read()

@app.route("/")
def index():
    return Response(_HTML, mimetype="text/html")

if __name__ == "__main__":
    app.run(debug=True)
