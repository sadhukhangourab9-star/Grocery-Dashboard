from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
import json, os
from datetime import datetime
from google.oauth2.service_account import Credentials
import gspread

app = Flask(__name__)
CORS(app)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
CREDS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")

def get_sheet():
    creds_data = json.loads(CREDS_JSON)
    creds = Credentials.from_service_account_info(creds_data, scopes=SCOPES)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SHEET_ID)

def get_or_create_worksheet(spreadsheet, title):
    try:
        return spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=title, rows=200, cols=50)

# ─── PRODUCTS ───────────────────────────────────────────────────────────────

@app.route("/api/products", methods=["GET"])
def get_products():
    try:
        sh = get_sheet()
        ws = get_or_create_worksheet(sh, "Products")
        rows = ws.get_all_values()
        products = [r[0] for r in rows if r and r[0].strip()]
        return jsonify({"products": products})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/products", methods=["POST"])
def save_products():
    try:
        data = request.json
        products = data.get("products", [])
        sh = get_sheet()
        ws = get_or_create_worksheet(sh, "Products")
        ws.clear()
        if products:
            ws.update("A1", [[p] for p in products])
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ─── MONTHS ─────────────────────────────────────────────────────────────────

@app.route("/api/months", methods=["GET"])
def get_months():
    try:
        sh = get_sheet()
        titles = [ws.title for ws in sh.worksheets() if ws.title not in ("Products",)]
        return jsonify({"months": titles})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/months", methods=["POST"])
def create_month():
    try:
        data = request.json
        month_name = data.get("month")
        sh = get_sheet()
        # Check if month already exists
        existing = [ws.title for ws in sh.worksheets()]
        if month_name in existing:
            return jsonify({"error": "Month already exists"}), 400

        # Get products
        products_ws = get_or_create_worksheet(sh, "Products")
        products = [r[0] for r in products_ws.get_all_values() if r and r[0].strip()]

        ws = sh.add_worksheet(title=month_name, rows=200, cols=len(products)+5)
        
        # Build header row: Account + products
        header = ["Account / Slot"] + products
        ws.update("A1", [header])

        # Pre-fill Old Stock row
        old_stock_row = ["Old Stock"] + [""] * len(products)

        # Check if previous month exists to carry forward old stock
        if len(existing) > 1:
            # find last month sheet (not Products)
            prev_months = [t for t in existing if t not in ("Products",)]
            if prev_months:
                prev_ws = sh.worksheet(prev_months[-1])
                prev_data = prev_ws.get_all_values()
                # find Current Stock row
                for row in prev_data:
                    if row and row[0].strip().lower() == "current stock":
                        old_stock_row = ["Old Stock"] + row[1:]
                        break

        ws.update("A2", [old_stock_row])

        # Add section labels
        ws.update(f"A{len(products)+10}", [["Total"]])
        ws.update(f"A{len(products)+12}", [["Rajkumar da sold"]])
        ws.update(f"A{len(products)+13}", [["Souvik da Sold"]])
        ws.update(f"A{len(products)+14}", [["Gourab Sold"]])
        ws.update(f"A{len(products)+16}", [["Current Stock"]])

        # Format header bold
        ws.format("A1:Z1", {"textFormat": {"bold": True}, "backgroundColor": {"red": 0.2, "green": 0.6, "blue": 0.3}})

        return jsonify({"success": True, "month": month_name})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ─── ORDERS ─────────────────────────────────────────────────────────────────

@app.route("/api/orders/<month>", methods=["GET"])
def get_orders(month):
    try:
        sh = get_sheet()
        ws = sh.worksheet(month)
        data = ws.get_all_values()
        if not data:
            return jsonify({"headers": [], "rows": []})
        headers = data[0]
        rows = []
        for i, row in enumerate(data[1:], start=2):
            rows.append({"row_index": i, "values": row})
        return jsonify({"headers": headers, "rows": rows})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/orders/<month>", methods=["POST"])
def add_order(month):
    try:
        data = request.json
        account = data.get("account")  # e.g. "Amazon (9/1/26) 12:00 PM SLOT"
        quantities = data.get("quantities", {})  # {product: qty}

        sh = get_sheet()
        ws = sh.worksheet(month)
        all_data = ws.get_all_values()
        headers = all_data[0] if all_data else []
        products = headers[1:] if headers else []

        row_values = [account]
        for p in products:
            row_values.append(str(quantities.get(p, "")))

        # Find insertion point: before Total row
        insert_row = len(all_data) + 1
        for i, row in enumerate(all_data):
            if row and row[0].strip().lower() == "total":
                insert_row = i + 1  # 1-indexed
                break

        ws.insert_row(row_values, index=insert_row)
        _recalculate_totals(ws, headers)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/orders/<month>/verify", methods=["POST"])
def verify_order(month):
    try:
        data = request.json
        row_index = data.get("row_index")
        sh = get_sheet()
        ws = sh.worksheet(month)
        all_data = ws.get_all_values()
        headers = all_data[0] if all_data else []
        n_cols = len(headers)
        # Add ✓ to last used col or a status col
        verify_col = n_cols + 1
        ws.update_cell(row_index, verify_col, "✓ Delivered")
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ─── SALES ──────────────────────────────────────────────────────────────────

@app.route("/api/sales/<month>", methods=["POST"])
def save_sales(month):
    try:
        data = request.json
        sales = data.get("sales", {})  # {seller: {product: qty}}

        sh = get_sheet()
        ws = sh.worksheet(month)
        all_data = ws.get_all_values()
        headers = all_data[0] if all_data else []
        products = headers[1:] if len(headers) > 1 else []

        sellers = ["Rajkumar da sold", "Souvik da Sold", "Gourab Sold"]
        
        # Find or create seller rows
        for seller in sellers:
            row_idx = None
            for i, row in enumerate(all_data):
                if row and row[0].strip().lower() == seller.lower():
                    row_idx = i + 1
                    break
            
            seller_quantities = sales.get(seller, {})
            row_values = [seller] + [str(seller_quantities.get(p, "")) for p in products]
            
            if row_idx:
                ws.update(f"A{row_idx}", [row_values])
            else:
                ws.append_row(row_values)

        _recalculate_current_stock(ws, headers, all_data)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def _recalculate_totals(ws, headers):
    """Recalculate Total row = sum of all order rows (excluding Old Stock, Total, sold, Current Stock)"""
    try:
        all_data = ws.get_all_values()
        skip_labels = {"total", "rajkumar da sold", "souvik da sold", "gourab sold", "current stock"}
        products = headers[1:]

        totals = {p: 0 for p in products}
        for row in all_data[1:]:
            if not row or not row[0].strip():
                continue
            label = row[0].strip().lower()
            if label in skip_labels:
                continue
            for i, p in enumerate(products):
                col_i = i + 1
                val = row[col_i] if col_i < len(row) else ""
                try:
                    totals[p] += int(val) if val else 0
                except:
                    pass

        total_row_idx = None
        for i, row in enumerate(all_data):
            if row and row[0].strip().lower() == "total":
                total_row_idx = i + 1
                break

        if total_row_idx:
            total_vals = ["Total"] + [str(totals.get(p, 0)) for p in products]
            ws.update(f"A{total_row_idx}", [total_vals])
    except:
        pass

def _recalculate_current_stock(ws, headers, all_data):
    """Current Stock = Total - sum(sold rows)"""
    try:
        products = headers[1:]
        total_row = {}
        sold_totals = {p: 0 for p in products}
        sold_labels = {"rajkumar da sold", "souvik da sold", "gourab sold"}

        for row in all_data:
            if not row:
                continue
            label = row[0].strip().lower()
            if label == "total":
                for i, p in enumerate(products):
                    try:
                        total_row[p] = int(row[i+1]) if i+1 < len(row) and row[i+1] else 0
                    except:
                        total_row[p] = 0
            elif label in sold_labels:
                for i, p in enumerate(products):
                    try:
                        sold_totals[p] += int(row[i+1]) if i+1 < len(row) and row[i+1] else 0
                    except:
                        pass

        current_stock = {p: total_row.get(p, 0) - sold_totals.get(p, 0) for p in products}

        cs_row_idx = None
        for i, row in enumerate(all_data):
            if row and row[0].strip().lower() == "current stock":
                cs_row_idx = i + 1
                break

        if cs_row_idx:
            cs_vals = ["Current Stock"] + [str(current_stock.get(p, 0)) for p in products]
            ws.update(f"A{cs_row_idx}", [cs_vals])
    except:
        pass

@app.route("/api/dashboard/<month>", methods=["GET"])
def get_dashboard(month):
    """Returns current stock row and order count for the dashboard summary."""
    try:
        sh = get_sheet()
        ws = sh.worksheet(month)
        all_data = ws.get_all_values()
        if not all_data:
            return jsonify({"headers": [], "current_stock": [], "order_count": 0})

        headers = all_data[0]
        products = headers[1:] if len(headers) > 1 else []
        current_stock = {}
        order_count = 0
        skip_labels = {"total", "rajkumar da sold", "souvik da sold", "gourab sold",
                       "current stock", "old stock"}

        for row in all_data[1:]:
            if not row or not row[0].strip():
                continue
            label = row[0].strip().lower()
            if label == "current stock":
                for i, p in enumerate(products):
                    try:
                        current_stock[p] = int(row[i + 1]) if i + 1 < len(row) and row[i + 1] else 0
                    except:
                        current_stock[p] = 0
            elif label not in skip_labels:
                order_count += 1

        return jsonify({
            "headers": products,
            "current_stock": current_stock,
            "order_count": order_count
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/orders/<month>/row/<int:row_index>", methods=["DELETE"])
def delete_order_row(month, row_index):
    """Delete a specific order row by its 1-based sheet row index."""
    try:
        sh = get_sheet()
        ws = sh.worksheet(month)
        ws.delete_rows(row_index)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/")
def index():
    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
