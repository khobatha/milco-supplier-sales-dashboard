// assets/admin.js
// MiLCo Admin — Monthly Batch Generator (Template v2: includes MONTH + YEAR)
// - Robust parsing + verification counters (so no row is silently skipped)
// - Generates BANK / MOMO batches, updated transactions ledger
// - Appends period suffix to downloads e.g. bank-payment-batch-dec-2025.csv

const BANKING_DETAILS_URL = "data/banking-details.csv";
const LEDGER_URL = "data/transactions.csv";

const $ = (id) => document.getElementById(id);

const MONTHS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
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
  const idx = MONTHS.findIndex((x) => x.toUpperCase() === m.toUpperCase());
  return idx; // -1 if not found
}

// "December" -> "dec"
function monthShort(monthName) {
  const idx = monthIndex(monthName);
  if (idx === -1) return clean(monthName).slice(0, 3).toLowerCase();
  return MONTHS[idx].slice(0, 3).toLowerCase();
}

function parseMonthYearFromText(text) {
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
      skipEmptyLines: false, // we count/drop empties ourselves for verification
      transformHeader: (h) => clean(h),
      complete: (res) => resolve(res),
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
      transformHeader: (h) => clean(h),
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

// tolerant numeric parsing for amounts like "1,200" or "M 1,200.50"
function parseAmount(val) {
  const s = clean(val);
  if (!s) return NaN;
  const cleaned = s.replace(/,/g, "").replace(/^M\s*/i, "");
  return Number(cleaned);
}

function normalizeLedgerRows(ledgerRows) {
  // Ensures ledger rows adhere to:
  // MONTH,YEAR,PERIOD,COMPANY NAME,AMOUNT,MODE,REFERENCE
  const upgraded = [];

  for (const r of (ledgerRows || [])) {
    const company = clean(r["COMPANY NAME"] ?? r["COMPANY"] ?? r["NAME"]);
    const amount = Number(r["AMOUNT"] ?? r["Sum"] ?? r["TOTAL"]);
    const mode = clean(r["MODE"]);
    const reference = clean(r["REFERENCE"] ?? r["COMMENT"] ?? r["PERIOD"] ?? "");

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

function isAllEmptyRow(obj) {
  return !Object.values(obj || {}).some((v) => clean(v) !== "");
}

function setVerifyVisible(visible) {
  const box = $("verifyBox");
  if (!box) return;
  box.style.display = visible ? "block" : "none";
}

function setBadge(id, text) {
  const el = $(id);
  if (el) el.textContent = text;
}

function renderVerifyTable(metrics) {
  const tbody = document.querySelector("#verifyTable tbody");
  if (!tbody) return;
  tbody.innerHTML = "";
  for (const [k, v] of metrics) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${k}</td><td>${v}</td>`;
    tbody.appendChild(tr);
  }
}

$("processBtn")?.addEventListener("click", async () => {
  $("status").textContent = "Loading banking details + ledger...";
  $("downloads").innerHTML = "";
  $("exceptions").textContent = "";
  setVerifyVisible(false);

  const file = $("salesFile").files?.[0];
  if (!file) {
    $("status").textContent = "Please choose a monthly sales report CSV first.";
    return;
  }

  const threshold = Number($("threshold").value ?? 400);

  const [banking, ledgerExistingRaw, salesParse] = await Promise.all([
    parseCsvUrl(BANKING_DETAILS_URL),
    parseCsvUrl(LEDGER_URL).catch(() => []),
    parseCsvFile(file)
  ]);

  const ledgerExisting = normalizeLedgerRows(ledgerExistingRaw);

  // Sales parsing details
  const salesRaw = salesParse.data || [];
  const salesErrors = salesParse.errors || [];

  const parsedRowCount = salesRaw.length;
  const emptyRowCount = salesRaw.filter(isAllEmptyRow).length;

  // Remove fully-empty rows but KEEP count for verification
  const sales = salesRaw.filter((r) => !isAllEmptyRow(r));
  const nonEmpty = sales.length;

  // Validate required columns exist (after header trimming)
  const requiredCols = ["COMPANY NAME", "SUM of COST", "MONTH", "YEAR"];
  const sampleRow = sales[0] || {};
  const missing = requiredCols.filter(
    (c) => !Object.prototype.hasOwnProperty.call(sampleRow, c)
  );
  if (missing.length) {
    $("status").textContent =
      `Sales report template mismatch. Missing column(s): ${missing.join(", ")}. ` +
      `Expected: COMPANY NAME, SUM of COST, COMMENT, MONTH, YEAR`;
    return;
  }

  // Index supplier master data by normalized name
  const supplierMap = new Map();
  banking.forEach((r) => supplierMap.set(normName(r["COMPANY NAME"]), r));

  const bankBatch = [];
  const momoBatch = [];
  const exceptions = [];
  const invalidRows = []; // rows skipped due to missing/invalid required fields
  const ledgerNew = [];

  // Process sales rows with full classification (no silent skipping)
  for (let i = 0; i < sales.length; i++) {
    const r = sales[i];

    const rawName = r["COMPANY NAME"];
    const nameKey = normName(rawName);

    const amount = parseAmount(r["SUM of COST"]);
    const comment = clean(r["COMMENT"]);
    const month = clean(r["MONTH"]);
    const year = Number(clean(r["YEAR"]));

    // Validate fields and log invalid rows instead of silently skipping
    const reasons = [];
    if (!nameKey) reasons.push("Missing COMPANY NAME");
    if (!Number.isFinite(amount)) reasons.push("Invalid SUM of COST (not a number)");
    if (!month) reasons.push("Missing MONTH");
    if (!Number.isFinite(year)) reasons.push("Invalid YEAR");

    if (reasons.length) {
      invalidRows.push({
        "ROW_NUMBER": i + 2, // +2 (header + 1-based)
        "COMPANY NAME": clean(rawName),
        "SUM of COST": clean(r["SUM of COST"]),
        "COMMENT": comment,
        "MONTH": month,
        "YEAR": clean(r["YEAR"]),
        "ISSUE": reasons.join("; ")
      });
      continue;
    }

    const supplier = supplierMap.get(nameKey);
    const mode = amount >= threshold ? "BANK" : "MOMO";
    const period = `${month} ${year}`;
    const reference = comment || `MiLCo ${period} Sales`;

    if (!supplier) {
      exceptions.push({
        "COMPANY NAME": clean(rawName),
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

  // Determine period suffix from processed rows (assumes single month/year per upload)
  let periodMonth = "";
  let periodYear = "";
  if (ledgerNew.length > 0) {
    periodMonth = ledgerNew[0].MONTH;
    periodYear = ledgerNew[0].YEAR;
  } else if (exceptions.length > 0) {
    // fall back if everything landed in exceptions (rare, but possible)
    periodMonth = exceptions[0].MONTH;
    periodYear = exceptions[0].YEAR;
  } else if (invalidRows.length > 0) {
    // fall back if everything invalid
    periodMonth = invalidRows[0].MONTH;
    periodYear = invalidRows[0].YEAR;
  }

  const periodSuffix =
    periodMonth && periodYear ? `-${monthShort(periodMonth)}-${periodYear}` : "";

  const ledgerUpdated = [...ledgerExisting, ...ledgerNew];

  // Sort ledger by YEAR then MONTH
  ledgerUpdated.sort((a, b) => {
    const ya = Number(a.YEAR) || 0;
    const yb = Number(b.YEAR) || 0;
    if (ya !== yb) return ya - yb;
    const ma = monthIndex(a.MONTH);
    const mb = monthIndex(b.MONTH);
    return (ma === -1 ? 99 : ma) - (mb === -1 ? 99 : mb);
  });

  // Prepare downloadable files (with period suffix)
  const files = [];

  files.push({
    name: `bank-payment-batch${periodSuffix}.csv`,
    csv: toCsv(bankBatch, ["NAME", "ACCOUNT", "BRANCH", "AMOUNT", "COMMENT"])
  });

  files.push({
    name: `momo-payment-batch${periodSuffix}.csv`,
    csv: toCsv(momoBatch, ["NAME", "MOMO PROVIDER", "MOMO NUMBER", "MOMO NAMES", "AMOUNT", "COMMENT"])
  });

  files.push({
    name: `transactions${periodSuffix}.csv`,
    csv: toCsv(ledgerUpdated, ["MONTH", "YEAR", "PERIOD", "COMPANY NAME", "AMOUNT", "MODE", "REFERENCE"])
  });

  if (exceptions.length) {
    files.push({
      name: `exceptions${periodSuffix}.csv`,
      csv: toCsv(exceptions, ["COMPANY NAME", "AMOUNT", "MODE", "MONTH", "YEAR", "ISSUE"])
    });
  }

  if (invalidRows.length) {
    files.push({
      name: `invalid-rows${periodSuffix}.csv`,
      csv: toCsv(invalidRows, ["ROW_NUMBER", "COMPANY NAME", "SUM of COST", "COMMENT", "MONTH", "YEAR", "ISSUE"])
    });
  }

  // Render download buttons
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
    ? `<b>${exceptions.length}</b> issue(s) found. Download <b>exceptions${periodSuffix}.csv</b> and fix <code>data/banking-details.csv</code>.`
    : "No exceptions. All good.";

  // --- Verification summary
  const bankCount = bankBatch.length;
  const momoCount = momoBatch.length;
  const excCount = exceptions.length;
  const invalidCount = invalidRows.length;
  const ledgerAdd = ledgerNew.length;

  // verification: every non-empty sales row must end up in exactly one bucket:
  // bankBatch OR momoBatch OR exceptions OR invalidRows
  const verificationOk = nonEmpty === (bankCount + momoCount + excCount + invalidCount);

  setVerifyVisible(true);
  setBadge("vParsed", `Parsed: ${parsedRowCount}`);
  setBadge("vNonEmpty", `Non-empty: ${nonEmpty}`);
  setBadge("vInvalid", `Invalid: ${invalidCount}`);
  setBadge("vBank", `BANK: ${bankCount}`);
  setBadge("vMomo", `MOMO: ${momoCount}`);
  setBadge("vExceptions", `Exceptions: ${excCount}`);
  setBadge("vLedger", `Ledger add: ${ledgerAdd}`);

  renderVerifyTable([
    ["Rows parsed by parser (including empties)", parsedRowCount],
    ["Fully empty rows dropped", emptyRowCount],
    ["Non-empty rows processed", nonEmpty],
    ["BANK rows generated", bankCount],
    ["MOMO rows generated", momoCount],
    ["Exceptions (missing supplier/details)", excCount],
    ["Invalid rows (skipped with reason)", invalidCount],
    ["Ledger rows added (BANK+MOMO)", ledgerAdd],
    ["Parser reported errors (if any)", salesErrors.length],
    ["Verification check passed", verificationOk ? "YES" : "NO"]
  ]);

  const note = $("verifyNote");
  if (note) {
    note.innerHTML = verificationOk
      ? `Verification passed. Non-empty rows (${nonEmpty}) = BANK (${bankCount}) + MOMO (${momoCount}) + Exceptions (${excCount}) + Invalid (${invalidCount}).`
      : `<b>Verification failed.</b> Non-empty rows (${nonEmpty}) do not match outputs. Download <b>invalid-rows${periodSuffix}.csv</b> to see which rows were skipped and why.`;
  }

  $("status").textContent =
    `Done${periodSuffix}. Parsed: ${parsedRowCount}. Non-empty: ${nonEmpty}. ` +
    `Bank: ${bankCount}, MoMo: ${momoCount}, Exceptions: ${excCount}, Invalid: ${invalidCount}. ` +
    `Ledger added: ${ledgerAdd}.`;
});
