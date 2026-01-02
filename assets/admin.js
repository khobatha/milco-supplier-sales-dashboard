// assets/admin.js
// MiLCo Admin — Monthly Batch Generator (v2 template: includes MONTH + YEAR)

const BANKING_DETAILS_URL = "data/banking-details.csv";
const LEDGER_URL = "data/transactions.csv";

const $ = (id) => document.getElementById(id);

const MONTHS = [
  "January","February","March","April","May","June",
  "July","August","September","October","November","December"
];

function normName(s) {
  return (s ?? "")
    .toString()
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}

function clean(s) {
  return (s ?? "").toString().trim();
}

function monthIndex(monthName) {
  const m = clean(monthName);
  const idx = MONTHS.findIndex(x => x.toUpperCase() === m.toUpperCase());
  return idx; // -1 if not found
}

function parseMonthYearFromText(text) {
  // Best effort: find "October 2025" inside a string like "MiLCo October 2025 Sales"
  const t = clean(text);
  const re = new RegExp(`\\b(${MONTHS.join("|")})\\b\\s+(\\d{4})`, "i");
  const m = t.match(re);
  if (!m) return null;
  return { month: m[1], year: Number(m[2]), period: `${m[1]} ${m[2]}` };
}

function parseCsvFile(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (res) => resolve(res.data),
      error: reject
    });
  });
}

function parseCsvUrl(url) {
  return new Promise((resolve, reject) => {
    Papa.parse(url, {
      download: true,
      header: true,
      skipEmptyLines: true,
      complete: (res) => resolve(res.data),
      error: reject
    });
  });
}

function toCsv(rows, headers) {
  return Papa.unparse(rows, { columns: headers });
}

function downloadText(filename, content) {
  const blob = new Blob([content], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

function money(n) {
  const x = Number(n);
  return Number.isFinite(x) ? x.toFixed(2) : "";
}

function normalizeLedgerRows(ledgerRows) {
  // Supports old ledger formats by upgrading rows into:
  // MONTH,YEAR,PERIOD,COMPANY NAME,AMOUNT,MODE,REFERENCE
  const upgraded = [];

  for (const r of (ledgerRows || [])) {
    const company = clean(r["COMPANY NAME"] ?? r["COMPANY"] ?? r["NAME"]);
    const amount = Number(r["AMOUNT"] ?? r["Sum"] ?? r["TOTAL"]);
    const mode = clean(r["MODE"]);
    const reference = clean(r["REFERENCE"] ?? r["COMMENT"] ?? r["PERIOD"] ?? "");

    // Prefer explicit MONTH/YEAR if present
    let month = clean(r["MONTH"]);
    let year = Number(r["YEAR"]);

    let period = clean(r["PERIOD"]);

    if ((!month || !Number.isFinite(year)) && period) {
      const parsed = parseMonthYearFromText(period);
      if (parsed) {
        month = parsed.month;
        year = parsed.year;
        period = parsed.period;
      }
    }

    if ((!month || !Number.isFinite(year)) && reference) {
      const parsed = parseMonthYearFromText(reference);
      if (parsed) {
        month = month || parsed.month;
        year = Number.isFinite(year) ? year : parsed.year;
        period = period || parsed.period;
      }
    }

    if (!period && month && Number.isFinite(year)) period = `${month} ${year}`;

    // Only include valid rows
    if (!company || !Number.isFinite(amount)) continue;

    upgraded.push({
      "MONTH": month || "",
      "YEAR": Number.isFinite(year) ? String(year) : "",
      "PERIOD": period || "",
      "COMPANY NAME": company,
      "AMOUNT": money(amount),
      "MODE": mode || "",
      "REFERENCE": reference || ""
    });
  }

  return upgraded;
}

$("processBtn").addEventListener("click", async () => {
  $("status").textContent = "Loading banking details + ledger...";
  $("downloads").innerHTML = "";
  $("exceptions").textContent = "";

  const file = $("salesFile").files?.[0];
  if (!file) {
    $("status").textContent = "Please choose a monthly sales report CSV first.";
    return;
  }

  const threshold = Number($("threshold").value ?? 400);

  const [banking, ledgerExistingRaw, sales] = await Promise.all([
    parseCsvUrl(BANKING_DETAILS_URL),
    parseCsvUrl(LEDGER_URL).catch(() => []),
    parseCsvFile(file)
  ]);

  const ledgerExisting = normalizeLedgerRows(ledgerExistingRaw);

  // Index supplier master data by normalized name
  const bankMap = new Map();
  banking.forEach(r => bankMap.set(normName(r["COMPANY NAME"]), r));

  const bankBatch = [];
  const momoBatch = [];
  const exceptions = [];
  const ledgerNew = [];

  // Validate sales template v2 columns exist
  const requiredCols = ["COMPANY NAME", "SUM of COST", "MONTH", "YEAR"];
  const missing = requiredCols.filter(c => !Object.prototype.hasOwnProperty.call(sales[0] || {}, c));
  if (missing.length) {
    $("status").textContent =
      `Sales report template mismatch. Missing column(s): ${missing.join(", ")}. ` +
      `Expected: COMPANY NAME, SUM of COST, COMMENT, MONTH, YEAR`;
    return;
  }

  for (const r of sales) {
    const rawName = r["COMPANY NAME"];
    const nameKey = normName(rawName);

    const amount = Number(r["SUM of COST"]);
    const comment = clean(r["COMMENT"]);
    const month = clean(r["MONTH"]);
    const year = Number(r["YEAR"]);

    if (!nameKey || !Number.isFinite(amount) || !month || !Number.isFinite(year)) continue;

    const supplier = bankMap.get(nameKey);
    const mode = (amount >= threshold) ? "BANK" : "MOMO";

    const period = `${month} ${year}`;
    const reference = comment || `MiLCo ${period} Sales`;

    if (!supplier) {
      exceptions.push({
        "COMPANY NAME": rawName,
        "AMOUNT": money(amount),
        "MODE": mode,
        "MONTH": month,
        "YEAR": String(year),
        "ISSUE": "Supplier not found in banking-details.csv"
      });
      continue;
    }

    if (mode === "BANK") {
      const account = clean(supplier["ACCOUNT"]);
      const branch = clean(supplier["BRANCH"]);

      if (!account || !branch) {
        exceptions.push({
          "COMPANY NAME": supplier["COMPANY NAME"],
          "AMOUNT": money(amount),
          "MODE": "BANK",
          "MONTH": month,
          "YEAR": String(year),
          "ISSUE": "Missing ACCOUNT or BRANCH"
        });
      } else {
        bankBatch.push({
          "NAME": supplier["COMPANY NAME"],
          "ACCOUNT": account,
          "BRANCH": branch,
          "AMOUNT": money(amount),
          "COMMENT": reference
        });

        ledgerNew.push({
          "MONTH": month,
          "YEAR": String(year),
          "PERIOD": period,
          "COMPANY NAME": supplier["COMPANY NAME"],
          "AMOUNT": money(amount),
          "MODE": "BANK",
          "REFERENCE": reference
        });
      }
    } else {
      const momoProvider = clean(supplier["MOMO"]);
      const momoNumber = clean(supplier["MOMO NUMBER"]);
      const momoNames = clean(supplier["MOMO NAMES"]);

      if (!momoProvider || !momoNumber) {
        exceptions.push({
          "COMPANY NAME": supplier["COMPANY NAME"],
          "AMOUNT": money(amount),
          "MODE": "MOMO",
          "MONTH": month,
          "YEAR": String(year),
          "ISSUE": "Missing MOMO or MOMO NUMBER"
        });
      } else {
        momoBatch.push({
          "NAME": supplier["COMPANY NAME"],
          "MOMO PROVIDER": momoProvider,
          "MOMO NUMBER": momoNumber,
          "MOMO NAMES": momoNames,
          "AMOUNT": money(amount),
          "COMMENT": reference
        });

        ledgerNew.push({
          "MONTH": month,
          "YEAR": String(year),
          "PERIOD": period,
          "COMPANY NAME": supplier["COMPANY NAME"],
          "AMOUNT": money(amount),
          "MODE": "MOMO",
          "REFERENCE": reference
        });
      }
    }
  }

  const ledgerUpdated = [...ledgerExisting, ...ledgerNew];

  // Sort ledger by YEAR then MONTH order (nice for charts/history)
  ledgerUpdated.sort((a, b) => {
    const ya = Number(a.YEAR) || 0, yb = Number(b.YEAR) || 0;
    if (ya !== yb) return ya - yb;
    const ma = monthIndex(a.MONTH), mb = monthIndex(b.MONTH);
    return (ma === -1 ? 99 : ma) - (mb === -1 ? 99 : mb);
  });

  const files = [];

  files.push({
    name: "bank-payment-batch.csv",
    csv: toCsv(bankBatch, ["NAME","ACCOUNT","BRANCH","AMOUNT","COMMENT"])
  });

  files.push({
    name: "momo-payment-batch.csv",
    csv: toCsv(momoBatch, ["NAME","MOMO PROVIDER","MOMO NUMBER","MOMO NAMES","AMOUNT","COMMENT"])
  });

  files.push({
    name: "transactions.csv",
    csv: toCsv(ledgerUpdated, ["MONTH","YEAR","PERIOD","COMPANY NAME","AMOUNT","MODE","REFERENCE"])
  });

  if (exceptions.length) {
    files.push({
      name: "exceptions.csv",
      csv: toCsv(exceptions, ["COMPANY NAME","AMOUNT","MODE","MONTH","YEAR","ISSUE"])
    });
  }

  // Render download links
  const ul = $("downloads");
  for (const f of files) {
    const li = document.createElement("li");
    const btn = document.createElement("button");
    btn.textContent = `Download ${f.name}`;
    btn.addEventListener("click", () => downloadText(f.name, f.csv));
    li.appendChild(btn);
    ul.appendChild(li);
  }

  $("exceptions").innerHTML = exceptions.length
    ? `<b>${exceptions.length}</b> issue(s) found. Download <b>exceptions.csv</b> and fix <code>data/banking-details.csv</code>.`
    : "No exceptions. All good.";

  $("status").textContent =
    `Done. Bank rows: ${bankBatch.length}, MoMo rows: ${momoBatch.length}, Ledger added: ${ledgerNew.length}.`;
});
