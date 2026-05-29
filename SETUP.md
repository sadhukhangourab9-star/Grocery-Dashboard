# Grocery Dashboard — Setup Guide

## Step 1: Google Sheets & Service Account

1. Go to https://console.cloud.google.com
2. Create a new project (e.g. "Grocery Dashboard")
3. Enable **Google Sheets API** and **Google Drive API**
4. Go to **IAM & Admin → Service Accounts → Create Service Account**
5. Give it any name, click Done
6. Click the service account → **Keys → Add Key → JSON** → Download
7. Create a new Google Sheet at https://sheets.google.com
8. Copy the Sheet ID from the URL:
   `https://docs.google.com/spreadsheets/d/**THIS_PART**/edit`
9. Share the Google Sheet with the service account email (from the JSON file, field: `client_email`)
   → Give it **Editor** access

---

## Step 2: GitHub Setup

1. Create a new repo on GitHub (e.g. `grocery-dashboard`)
2. Push all project files:
```bash
git init
git add .
git commit -m "initial commit"
git remote add origin https://github.com/YOUR_USERNAME/grocery-dashboard.git
git push -u origin main
```

---

## Step 3: Deploy on Render

1. Go to https://render.com → New → Web Service
2. Connect your GitHub repo
3. Settings:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
   - **Environment:** Python 3

4. Add Environment Variables:
   - `GOOGLE_SHEET_ID` → paste your Sheet ID
   - `GOOGLE_CREDENTIALS_JSON` → paste the **entire contents** of the JSON key file (all on one line or as-is)

5. Click **Deploy**

---

## Step 4: First Use

1. Open your deployed URL
2. Go to **Products** tab → add all your products → click **Save to Sheets**
3. Go to **Months** → click **+ New Month** → type e.g. `June 2026`
   - This auto-creates a new sheet tab in Google Sheets with header + old stock row
4. Go to **Add Order** → select month, platform (Amazon/Flipkart/Blinkit), date, time slot → enter quantities → Submit
5. Click **Verify Delivery** button in Month View after receiving the order
6. At end of month → go to **Record Sales** → enter sold qty per seller → Save
   - Current Stock is auto-calculated
   - Next month's "Old Stock" row is auto-populated from Current Stock

---

## How it maps to your Excel

| Excel Row | Dashboard Equivalent |
|-----------|---------------------|
| Old stock | Auto-carried from previous month's Current Stock |
| Amazon 16/10/25 | Add Order → Amazon (16/10/25) 10 AM SLOT |
| Total | Auto-calculated in sheet |
| Rajkumar da sold | Record Sales → Rajkumar da tab |
| Souvik da Sold | Record Sales → Souvik da tab |
| Gourab Sold | Record Sales → Gourab tab |
| Current Stock | Auto-calculated = Total − Sales |
| OK (payment) | Verify Delivery button |

---

## Google Sheet Structure (auto-created)

Each month becomes a new tab:
- Row 1: Header (Account / Slot | Product1 | Product2 | ...)
- Row 2: Old Stock (carried from previous month)
- Rows 3–N: Order rows (Amazon, Flipkart, etc.)
- Row N+1: Total
- Row N+2: blank
- Row N+3: Rajkumar da sold
- Row N+4: Souvik da Sold
- Row N+5: Gourab Sold
- Row N+6: blank
- Row N+7: Current Stock
