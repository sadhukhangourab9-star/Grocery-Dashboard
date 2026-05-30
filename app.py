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

# ── /api/init-data  (products + slots in one call for UI) ─────────────────────
@app.route("/api/init-data")
def init_data():
    """Returns products, slots and today's date — used when loading order/receiver pages."""
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
        quantities = body.get("quantities", {})   # { "Product Name": qty }

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

# ── /api/receiver  (today's pending orders) ───────────────────────────────────
@app.route("/api/receiver")
def receiver():
    """
    Returns today's pending orders, grouped by slot.
    Also returns per-slot product totals for the summary table.
    """
    try:
        sh      = get_sheet()
        ws      = get_inv_ws(sh)
        data    = ws.get_all_values()
        if not data:
            return jsonify({"slots": [], "summary": {}, "orders": []})

        headers  = data[0]
        products = headers[len(FIXED_COLS):]
        today    = datetime.now().strftime("%Y-%m-%d")

        orders       = []
        slot_summary = {}   # { slot: { product: total } }

        for i, row in enumerate(data[1:], start=2):
            if not any(c.strip() for c in row):
                continue
            padded = row + [""] * max(0, len(headers) - len(row))

            row_date   = padded[0].strip()
            row_slot   = padded[1].strip()
            row_acct   = padded[2].strip()
            row_oname  = padded[3].strip()
            row_status = padded[4].strip().lower()

            # Skip structural rows
            if row_acct.lower() in ("old stock", "current stock", ""):
                continue
            # Only today's pending
            if row_date != today or row_status != "pending":
                continue

            qtys = {}
            for j, p in enumerate(products):
                col = len(FIXED_COLS) + j
                val = padded[col] if col < len(padded) else ""
                try:
                    qtys[p] = int(float(val)) if val else 0
                except:
                    qtys[p] = 0

            orders.append({
                "row_index":  i,
                "date":       row_date,
                "slot":       row_slot,
                "account":    row_acct,
                "order_name": row_oname,
                "quantities": qtys
            })

            if row_slot not in slot_summary:
                slot_summary[row_slot] = {p: 0 for p in products}
            for p, q in qtys.items():
                slot_summary[row_slot][p] = slot_summary[row_slot].get(p, 0) + q

        # Sort slots naturally
        all_slots = list(slot_summary.keys())

        return jsonify({
            "products":    products,
            "slots":       all_slots,
            "summary":     slot_summary,
            "orders":      orders,
            "today":       today
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── /api/mark-delivered ───────────────────────────────────────────────────────
@app.route("/api/mark-delivered", methods=["POST"])
def mark_delivered():
    try:
        row_index = request.json.get("row_index")
        sh  = get_sheet()
        ws  = get_inv_ws(sh)
        ws.update_cell(row_index, STATUS_COL, "Delivered")
        return jsonify({"success": True})
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
