# MiLCo Supplier Sales Dashboard — Web Admin Guide

This repository provides a read-only supplier sales dashboard and two admin web tools:

- **Batch Generator** (`admin.html`) — turns raw POS sales XLSX into monthly batches.
- **Supplier Payment Details** (`supplier-admin.html`) — views and updates `banking-details.csv`.

No backend is required. Everything runs in the browser.

## Quick Start (Local Web)

From the repo root:

```bash
python3 -m http.server 8000
```

Open in your browser:

- Dashboard: `http://localhost:8000/index.html`
- Admin (Batch Generator): `http://localhost:8000/admin.html`
- Supplier Admin: `http://localhost:8000/supplier-admin.html`

## 1) Batch Generator (admin.html)

### Step A — Generate Monthly Sales Summary (XLSX)

1. Upload the **raw POS sales report** (XLSX).  
   - The **first sheet** is used.
   - Cell **A1** must contain a title like: `MILCO JANUARY SALES 2026`
2. Click **Generate Monthly Sales Summary (XLSX)**.
3. Download the summary XLSX.  
   - If any invalid rows exist, download the invalid rows CSV and fix the raw source.

### Step B — Generate Bank/MoMo Batch Files (CSV)

1. Set the **Bank threshold** (default 400).
2. Click **Process & Generate Files**.
3. Download:
   - `bank-payment-batch-<mon>-<year>.csv`
   - `momo-payment-batch-<mon>-<year>.csv`
   - `transactions-<mon>-<year>.csv`
   - `exceptions-<mon>-<year>.csv` (if any)
   - `invalid-rows-<mon>-<year>.csv` (if any)

The verification box shows:
- Counts per category
- Totals for BANK / MOMO / Exceptions
- Total payable amount

### Final Step — Update Ledger

Replace `data/transactions.csv` with the downloaded `transactions-*.csv`, then commit and push.

## 2) Supplier Payment Details (supplier-admin.html)

### View Current Details

The table shows all suppliers and highlights missing details in red.

Use **Download Table (CSV/XLSX)** to export the current list.

### Update Missing Details Only

Upload either or both:

1. **Raw Sales Report (XLSX)**  
   Used to fill **BANK** and **ACCOUNT** for suppliers with missing data.
2. **MoMo Update Sheet (XLSX)**  
   Used to fill:
   - `MOMO` from **5. Select Your Mobile Money Platform**
   - `MOMO NUMBER` from **3. MPESA / Ecocash Number**
   - `MOMO NAMES` from **2. Contact Person Full Name**

Rules:
- Existing details are **never overwritten**.
- Only **missing fields** are filled.
- Suppliers not in `banking-details.csv` are **added**.
- Output columns are always uppercase for:
  - `COMPANY NAME`
  - `BANK`
  - `MOMO NAMES`

Click **Generate Updated banking-details.csv**, then download and replace `data/banking-details.csv`.

## Data Files

- `data/transactions.csv` — ledger used by the dashboard.
- `data/banking-details.csv` — supplier payment details.

## Troubleshooting

**“XLSX is not defined”**  
Make sure you are running via `http://localhost:8000` and not opening the HTML file directly.

**Missing columns detected**  
Check that your XLSX headers match the expected POS format.

---

If you want additional automation (auto-upload, validations, or custom exports), open an issue or request a feature update.
