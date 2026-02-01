# MiLCo Supplier Sales Dashboard — Quick Start

## 1) Start the Web App

From the project folder:

```bash
python3 -m http.server 8000
```

Open in your browser:

- Dashboard: `http://localhost:8000/index.html`
- Batch Generator: `http://localhost:8000/admin.html`
- Supplier Details: `http://localhost:8000/supplier-admin.html`

## 2) Generate Monthly Batches

1. Go to **Batch Generator** (`admin.html`).
2. Upload the **raw POS sales XLSX** (first sheet only).
3. Click **Generate Monthly Sales Summary (XLSX)** and download it.
4. Click **Process & Generate Files**.
5. Download the batch files and the updated `transactions-*.csv`.

## 3) Update Supplier Payment Details

1. Go to **Supplier Details** (`supplier-admin.html`).
2. Upload:
   - Raw Sales Report (XLSX) and/or
   - MoMo Update Sheet (XLSX)
3. Click **Generate Updated banking-details.csv**.
4. Download and replace `data/banking-details.csv`.

## 4) Final Step

Replace the old CSV in `data/` with the new one, then commit + push.
