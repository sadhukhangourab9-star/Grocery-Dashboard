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


# ── /api/setup  (full DB initialisation) ─────────────────────────────────────
@app.route("/api/setup", methods=["POST"])
def setup_db():
    """
    Expects JSON:
    {
      "products": ["Product A", "Product B", ...],
      "slots":    ["10 PM", "12 PM", ...]      (optional)
    }
    Creates / resets: Product_Master, Slot_Master, Sales_Log,
    and the current-month Inventory sheet.
    """
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
                ws = existing[title]
                ws.clear()
                return ws
            return sh.add_worksheet(title=title, rows=str(rows), cols=str(cols))

        # 1. Product_Master
        ws_prod = get_or_create("Product_Master", rows=200, cols=3)
        ws_prod.update("A1", [["Product Name"]])
        ws_prod.update(f"A2:A{len(products)+1}", [[p] for p in products])

        # 2. Slot_Master
        if not slots:
            slots = ["10 PM", "12 PM", "8 AM", "2 PM", "4 PM"]
        ws_slot = get_or_create("Slot_Master", rows=20, cols=2)
        ws_slot.update("A1", [["Slots"]])
        ws_slot.update(f"A2:A{len(slots)+1}", [[s] for s in slots])

        # 3. Sales_Log
        ws_sales = get_or_create("Sales_Log", rows=1000, cols=5)
        ws_sales.update("A1", [["Date", "Buyer Name", "Product Name", "Quantity Sold"]])

        # 4. Current-month Inventory sheet
        month_key = datetime.now().strftime("%b_%Y")   # e.g. May_2026
        inv_title = f"Inventory_{month_key}"
        headers   = FIXED_COLS + products
        cols_needed = len(headers) + 5

        ws_inv = get_or_create(inv_title, rows=500, cols=cols_needed)
        ws_inv.update("A1", [headers])
        old_row = ["-", "-", "Old Stock", "-", "Delivered"] + ["0"] * len(products)
        ws_inv.update("A2", [old_row])

        # Bold + colour the header row
        try:
            ws_inv.format("A1:ZZ1", {
                "textFormat":      {"bold": True},
                "backgroundColor": {"red": 0.13, "green": 0.37, "blue": 0.18}
            })
        except Exception:
            pass

        # Remove leftover "Sheet1" if it exists
        try:
            sh.del_worksheet(sh.worksheet("Sheet1"))
        except Exception:
            pass

        return jsonify({
            "success":       True,
            "month":         month_key,
            "inv_sheet":     inv_title,
            "product_count": len(products),
            "tabs_created":  ["Product_Master", "Slot_Master", "Sales_Log", inv_title]
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── /api/products  (read list from sheet) ────────────────────────────────────
@app.route("/api/products", methods=["GET"])
def get_products():
    try:
        sh   = get_sheet()
        ws   = sh.worksheet("Product_Master")
        vals = ws.col_values(1)
        return jsonify({"products": [v for v in vals[1:] if v.strip()]})
    except Exception as e:
        return jsonify({"error": str(e), "products": []}), 200


# ── frontend ──────────────────────────────────────────────────────────────────
TEMPLATE = os.path.join(os.path.dirname(__file__), "templates", "index.html")

@app.route("/")
def index():
    with open(TEMPLATE, encoding="utf-8") as f:
        return Response(f.read(), mimetype="text/html")

if __name__ == "__main__":
    app.run(debug=True)
