// assets/admin.js
// MiLCo Admin — Monthly Batch Generator
// - Parse raw XLSX sales report, generate monthly summary XLSX
// - Generate BANK / MOMO batches + updated transactions ledger (CSV)
// - Robust parsing + verification counters (no row is silently skipped)
// - Appends period suffix to downloads e.g. bank-payment-batch-dec-2025.csv

const BANKING_DETAILS_URL = "data/banking-details.csv";
const LEDGER_URL = "data/transactions.csv";

const $ = (id) => document.getElementById(id);

const MONTHS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
];

let summaryRows = null;
let summaryMeta = null;

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
  const monthRe = new RegExp(`\\b(${MONTHS.join("|")})\\b`, "i");
  const yearRe = /\b(\d{4})\b/;
  const m = t.match(monthRe);
  const y = t.match(yearRe);
  if (!m || !y) return null;
  return { month: m[1], year: Number(y[1]), period: `${m[1]} ${y[1]}` };
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

function downloadXlsx(filename, rows, sheetName) {
  const wb = XLSX.utils.book_new();
  const headers = ["COMPANY NAME", "SUM of COST", "COMMENT", "MONTH", "YEAR"];
  const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
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

function isAllEmptyArrayRow(arr) {
  return !(arr || []).some((v) => clean(v) !== "");
}

function normHeader(s) {
  return clean(s).toUpperCase().replace(/\s+/g, " ");
}

function pickCompanyNameCell(row, map) {
  if (map.has("COMPANY NAME")) return clean(row[map.get("COMPANY NAME")] || "");
  for (const [key, idx] of map.entries()) {
    if (key.includes("COMPANY") && key.includes("NAME")) {
      return clean(row[idx] || "");
    }
  }
  return "";
}

function paymentModeLabel(row) {
  const hasBank = clean(row["ACCOUNT"]) && clean(row["BANK"]);
  const hasMomo = clean(row["MOMO NUMBER"]) && clean(row["MOMO"]);
  if (hasBank && hasMomo) return "BANK + MOMO";
  if (hasBank) return "BANK";
  if (hasMomo) return "MOMO";
  return "MISSING";
}

function isMissingDetails(row) {
  const bankMissing = !clean(row["BANK"]) || !clean(row["ACCOUNT"]) || !clean(row["BRANCH"]);
  const momoMissing = !clean(row["MOMO"]) || !clean(row["MOMO NUMBER"]) || !clean(row["MOMO NAMES"]);
  return bankMissing || momoMissing;
}

function renderBankingTable(rows) {
  const tbody = document.querySelector("#bankingTable tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  rows.forEach((r) => {
    const tr = document.createElement("tr");
    const missing = isMissingDetails(r);
    tr.classList.add(missing ? "row-missing" : "row-complete");
    tr.innerHTML = `
      <td>${clean(r["COMPANY NAME"])}</td>
      <td>${paymentModeLabel(r)}</td>
      <td>${clean(r["BANK"])}</td>
      <td>${clean(r["ACCOUNT"])}</td>
      <td>${clean(r["BRANCH"])}</td>
      <td>${clean(r["MOMO"])}</td>
      <td>${clean(r["MOMO NUMBER"])}</td>
      <td>${clean(r["MOMO NAMES"])}</td>
      <td>${missing ? "Missing details" : "Complete"}</td>
    `;
    tbody.appendChild(tr);
  });
}

function normalizeBankingRows(rows) {
  const out = [];
  for (const r of rows || []) {
    const name = clean(r["COMPANY NAME"] ?? r["COMPANY"] ?? r["NAME"]);
    if (!name) continue;
    out.push({
      "COMPANY NAME": name,
      "BANK": clean(r["BANK"] ?? r["BANK NAME"]),
      "ACCOUNT": clean(r["ACCOUNT"] ?? r["BANK ACCOUNT/MOBILE"]),
      "BRANCH": clean(r["BRANCH"]),
      "MOMO": clean(r["MOMO"]),
      "MOMO NUMBER": clean(r["MOMO NUMBER"]),
      "MOMO NAMES": clean(r["MOMO NAMES"])
    });
  }
  return out;
}

function updateMissing(target, source, fields) {
  let changed = false;
  for (const f of fields) {
    if (!clean(target[f]) && clean(source[f])) {
      target[f] = clean(source[f]);
      changed = true;
    }
  }
  return changed;
}

function readXlsxFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const firstSheetName = wb.SheetNames[0];
        if (!firstSheetName) {
          reject(new Error("No sheets found in workbook."));
          return;
        }
        const ws = wb.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(ws, {
          header: 1,
          raw: false,
          defval: ""
        });
        resolve({ rows, sheetName: firstSheetName });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function findHeaderRow(rows) {
  const required = ["COMPANY NAME", "COST"];
  const expected = [
    "PRODUCTS", "QUANTITY", "SELLING", "COST", "PROFIT",
    "COMPANY NAME", "BANK ACCOUNT/MOBILE", "BANK NAME", "CONTACTS"
  ];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const map = new Map();
    row.forEach((cell, idx) => {
      const key = normHeader(cell);
      if (key) map.set(key, idx);
    });
    const hasRequired = required.every((h) => map.has(h));
    if (hasRequired) {
      const missingExpected = expected.filter((h) => !map.has(h));
      return { rowIndex: i, map, missingExpected };
    }
  }
  return null;
}

function parseRawBankingDetails(rows) {
  const headerInfo = findHeaderRow(rows);
  if (!headerInfo) return { error: "Could not find header row in raw sales report." };
  const { rowIndex, map } = headerInfo;
  const dataRows = rows.slice(rowIndex + 1);

  const out = new Map();
  dataRows.forEach((row) => {
    if (isAllEmptyArrayRow(row)) return;
    const company = pickCompanyNameCell(row, map);
    if (!company) return;
    const bank = clean(row[map.get("BANK NAME")] || "");
    const account = clean(row[map.get("BANK ACCOUNT/MOBILE")] || "");
    if (!bank && !account) return;
    const key = normName(company);
    if (!out.has(key)) {
      out.set(key, {
        "COMPANY NAME": company,
        "BANK": bank,
        "ACCOUNT": account
      });
    } else {
      const existing = out.get(key);
      if (!existing["BANK"] && bank) existing["BANK"] = bank;
      if (!existing["ACCOUNT"] && account) existing["ACCOUNT"] = account;
    }
  });

  return { map: out };
}

function parseMomoUpdateSheet(rows) {
  const headerInfo = findHeaderRow(rows) || { rowIndex: 0, map: new Map() };
  const startRow = headerInfo.rowIndex + 1;
  const map = headerInfo.map;

  const fallbackMap = new Map(map);
  rows[headerInfo.rowIndex]?.forEach((cell, idx) => {
    const key = normHeader(cell);
    if (key) fallbackMap.set(key, idx);
  });

  const out = new Map();
  for (let i = startRow; i < rows.length; i++) {
    const row = rows[i] || [];
    if (isAllEmptyArrayRow(row)) continue;
    const company = pickCompanyNameCell(row, fallbackMap);
    if (!company) continue;

    const momoProvider =
      clean(row[fallbackMap.get("5. SELECT YOUR MOBILE MONEY PLATFORM")] || "") ||
      clean(row[fallbackMap.get("SELECT YOUR MOBILE MONEY PLATFORM")] || "") ||
      clean(row[fallbackMap.get("MOBILE MONEY PLATFORM")] || "");

    const momoNumber =
      clean(row[fallbackMap.get("3. MPESA / ECOCASH NUMBER")] || "") ||
      clean(row[fallbackMap.get("MPESA / ECOCASH NUMBER")] || "") ||
      clean(row[fallbackMap.get("MOMO NUMBER")] || "");

    const momoNames =
      clean(row[fallbackMap.get("2. CONTACT PERSON FULL NAME")] || "") ||
      clean(row[fallbackMap.get("CONTACT PERSON FULL NAME")] || "") ||
      clean(row[fallbackMap.get("CONTACT NAME")] || "");

    if (!momoProvider && !momoNumber && !momoNames) continue;

    const key = normName(company);
    out.set(key, {
      "COMPANY NAME": company,
      "MOMO": momoProvider,
      "MOMO NUMBER": momoNumber,
      "MOMO NAMES": momoNames
    });
  }

  return { map: out };
}

async function loadBankingDetails() {
  try {
    const rows = await parseCsvUrl(BANKING_DETAILS_URL);
    const normalized = normalizeBankingRows(rows);
    renderBankingTable(normalized);
    return normalized;
  } catch (err) {
    const tbody = document.querySelector("#bankingTable tbody");
    if (tbody) {
      tbody.innerHTML = `<tr><td colspan="9">Failed to load banking-details.csv</td></tr>`;
    }
    return [];
  }
}

function parseRawSalesReport(rows) {
  const titleCell = (rows[0] || []).find((v) => clean(v) !== "") || "";
  const title = clean(titleCell);
  if (!title) {
    return { error: "Missing title in A1 (or first row)." };
  }

  const period = parseMonthYearFromText(title);
  if (!period) {
    return { error: "Could not extract MONTH and YEAR from the title." };
  }

  const headerInfo = findHeaderRow(rows);
  if (!headerInfo) {
    return { error: "Could not find header row with COMPANY NAME and COST." };
  }

  const { rowIndex, map, missingExpected } = headerInfo;
  const dataRows = rows.slice(rowIndex + 1);

  const suppliers = new Map();
  const supplierNames = new Map();
  const invalidRows = [];

  let parsedRowCount = dataRows.length;
  let nonEmpty = 0;
  let invalidCount = 0;
  let rawTotal = 0;

  dataRows.forEach((row, idx) => {
    if (isAllEmptyArrayRow(row)) return;
    nonEmpty++;

    const company = clean(row[map.get("COMPANY NAME")] || "");
    const costRaw = row[map.get("COST")] || "";
    const cost = parseAmount(costRaw);

    const reasons = [];
    if (!company) reasons.push("Missing COMPANY NAME");
    if (!Number.isFinite(cost)) reasons.push("Invalid COST (not a number)");

    if (reasons.length) {
      invalidCount++;
      invalidRows.push({
        "ROW_NUMBER": rowIndex + 2 + idx,
        "COMPANY NAME": company,
        "COST": clean(costRaw),
        "ISSUE": reasons.join("; ")
      });
      return;
    }

    const key = normName(company);
    supplierNames.set(key, company);
    suppliers.set(key, (suppliers.get(key) || 0) + cost);
    rawTotal += cost;
  });

  const summaryRows = [];
  for (const [key, sum] of suppliers.entries()) {
    summaryRows.push({
      "COMPANY NAME": supplierNames.get(key),
      "SUM of COST": money(sum),
      "COMMENT": title,
      "MONTH": period.month,
      "YEAR": String(period.year)
    });
  }

  const summaryTotal = summaryRows.reduce(
    (a, r) => a + (parseAmount(r["SUM of COST"]) || 0), 0
  );

  return {
    summaryRows,
    invalidRows,
    title,
    month: period.month,
    year: period.year,
    missingExpected,
    metrics: {
      parsedRowCount,
      nonEmpty,
      invalidCount,
      suppliers: summaryRows.length,
      rawTotal,
      summaryTotal
    }
  };
}

function setVerifyVisible(visible) {
  const box = $("verifyBox");
  if (!box) return;
  box.style.display = visible ? "block" : "none";
}

function setRawVerifyVisible(visible) {
  const box = $("rawVerifyBox");
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

$("generateSummaryBtn")?.addEventListener("click", async () => {
  $("summaryStatus").textContent = "Reading raw XLSX...";
  $("summaryDownloads").innerHTML = "";
  setRawVerifyVisible(false);
  summaryRows = null;
  summaryMeta = null;

  const file = $("rawFile").files?.[0];
  if (!file) {
    $("summaryStatus").textContent = "Please choose a raw sales XLSX file first.";
    return;
  }

  try {
    const { rows } = await readXlsxFile(file);
    const parsed = parseRawSalesReport(rows);
    if (parsed.error) {
      $("summaryStatus").textContent = parsed.error;
      return;
    }

    summaryRows = parsed.summaryRows;
    summaryMeta = {
      title: parsed.title,
      month: parsed.month,
      year: parsed.year
    };

    if (!summaryRows.length) {
      $("summaryStatus").textContent = "No valid data rows found after the header.";
      return;
    }

    const month = parsed.month;
    const year = parsed.year;
    const sheetName = `MILCO SUMMARY ${month} ${year}`;
    const filename = `milco-sales-summary-${monthShort(month)}-${year}.xlsx`;

    const dlWrap = $("summaryDownloads");
    const btn = document.createElement("button");
    btn.textContent = `Download ${filename}`;
    btn.classList.add("btn-success");
    btn.addEventListener("click", () => downloadXlsx(filename, summaryRows, sheetName));
    dlWrap.appendChild(btn);

    if (parsed.invalidRows.length) {
      const invalidName = `invalid-raw-rows-${monthShort(month)}-${year}.csv`;
      const invalidBtn = document.createElement("button");
      invalidBtn.style.marginLeft = "8px";
      invalidBtn.textContent = `Download ${invalidName}`;
      invalidBtn.classList.add("btn-danger");
      invalidBtn.addEventListener("click", () => {
        const csv = toCsv(parsed.invalidRows, ["ROW_NUMBER", "COMPANY NAME", "COST", "ISSUE"]);
        downloadText(invalidName, csv);
      });
      dlWrap.appendChild(invalidBtn);
    }

    if (parsed.missingExpected.length) {
      $("summaryStatus").textContent =
        `Summary ready. Warning: missing expected column(s): ${parsed.missingExpected.join(", ")}.`;
    } else {
      $("summaryStatus").textContent = "Summary ready. You can now generate batch files.";
    }

    const m = parsed.metrics;
    setRawVerifyVisible(true);
    setBadge("rParsed", `Parsed: ${m.parsedRowCount}`);
    setBadge("rNonEmpty", `Non-empty: ${m.nonEmpty}`);
    setBadge("rInvalid", `Invalid: ${m.invalidCount}`);
    setBadge("rSuppliers", `Suppliers: ${m.suppliers}`);
    setBadge("rRawTotal", `Raw total: ${money(m.rawTotal)}`);
    setBadge("rSumTotal", `Summary total: ${money(m.summaryTotal)}`);

    const verifyOk = Math.abs(m.rawTotal - m.summaryTotal) < 0.01;
    const note = $("rawVerifyNote");
    if (note) {
      note.innerHTML = verifyOk
        ? "Verification passed. Raw total COST equals Summary total COST."
        : "<b>Verification failed.</b> Raw total does not match Summary total.";
    }
  } catch (err) {
    $("summaryStatus").textContent = `Failed to read XLSX: ${err.message || err}`;
  }
});

$("updateBankingBtn")?.addEventListener("click", async () => {
  $("bankingStatus").textContent = "Loading banking details...";
  $("bankingDownloads").innerHTML = "";

  const [bankingRows] = await Promise.all([parseCsvUrl(BANKING_DETAILS_URL).catch(() => [])]);
  const banking = normalizeBankingRows(bankingRows);
  const bankMap = new Map();
  banking.forEach((r) => bankMap.set(normName(r["COMPANY NAME"]), r));

  const salesFile = $("bankUpdateSalesFile").files?.[0];
  const momoFile = $("bankUpdateMomoFile").files?.[0];

  if (!salesFile && !momoFile) {
    $("bankingStatus").textContent = "Please upload at least one XLSX file.";
    return;
  }

  let salesMap = new Map();
  let momoMap = new Map();

  if (salesFile) {
    try {
      const { rows } = await readXlsxFile(salesFile);
      const parsed = parseRawBankingDetails(rows);
      if (parsed.error) {
        $("bankingStatus").textContent = parsed.error;
        return;
      }
      salesMap = parsed.map;
    } catch (err) {
      $("bankingStatus").textContent = `Failed to read raw sales XLSX: ${err.message || err}`;
      return;
    }
  }

  if (momoFile) {
    try {
      const { rows } = await readXlsxFile(momoFile);
      const parsed = parseMomoUpdateSheet(rows);
      momoMap = parsed.map;
    } catch (err) {
      $("bankingStatus").textContent = `Failed to read MoMo update XLSX: ${err.message || err}`;
      return;
    }
  }

  let updatedCount = 0;
  let addedCount = 0;

  const allKeys = new Set([
    ...Array.from(bankMap.keys()),
    ...Array.from(salesMap.keys()),
    ...Array.from(momoMap.keys())
  ]);

  for (const key of allKeys) {
    const existing = bankMap.get(key);
    const sales = salesMap.get(key) || {};
    const momo = momoMap.get(key) || {};

    if (existing) {
      const changedBank = updateMissing(existing, sales, ["BANK", "ACCOUNT", "BRANCH"]);
      const changedMomo = updateMissing(existing, momo, ["MOMO", "MOMO NUMBER", "MOMO NAMES"]);
      if (changedBank || changedMomo) updatedCount++;
    } else {
      const name = sales["COMPANY NAME"] || momo["COMPANY NAME"];
      if (!name) continue;
      bankMap.set(key, {
        "COMPANY NAME": name,
        "BANK": clean(sales["BANK"] || ""),
        "ACCOUNT": clean(sales["ACCOUNT"] || ""),
        "BRANCH": "",
        "MOMO": clean(momo["MOMO"] || ""),
        "MOMO NUMBER": clean(momo["MOMO NUMBER"] || ""),
        "MOMO NAMES": clean(momo["MOMO NAMES"] || "")
      });
      addedCount++;
    }
  }

  const updatedRows = Array.from(bankMap.values()).sort((a, b) => {
    return clean(a["COMPANY NAME"]).localeCompare(clean(b["COMPANY NAME"]));
  });

  renderBankingTable(updatedRows);

  const csv = toCsv(updatedRows, [
    "COMPANY NAME",
    "BANK",
    "ACCOUNT",
    "BRANCH",
    "MOMO",
    "MOMO NUMBER",
    "MOMO NAMES"
  ]);

  const btn = document.createElement("button");
  btn.textContent = "Download updated banking-details.csv";
  btn.classList.add("btn-success");
  btn.addEventListener("click", () => downloadText("banking-details.csv", csv));
  $("bankingDownloads").appendChild(btn);

  $("bankingStatus").textContent =
    `Done. Updated: ${updatedCount}, Added: ${addedCount}. Download the updated file.`;
});

$("processBtn")?.addEventListener("click", async () => {
  $("status").textContent = "Loading banking details + ledger...";
  $("downloads").innerHTML = "";
  $("exceptions").textContent = "";
  setVerifyVisible(false);

  if (!summaryRows || !summaryRows.length) {
    $("status").textContent = "Please generate the Monthly Sales Summary (XLSX) first.";
    return;
  }

  const threshold = Number($("threshold").value ?? 400);

  const [banking, ledgerExistingRaw] = await Promise.all([
    parseCsvUrl(BANKING_DETAILS_URL),
    parseCsvUrl(LEDGER_URL).catch(() => [])
  ]);

  const ledgerExisting = normalizeLedgerRows(ledgerExistingRaw);

  // Sales parsing details (from generated summary)
  const salesRaw = summaryRows;
  const salesErrors = [];

  const parsedRowCount = salesRaw.length;
  const emptyRowCount = salesRaw.filter(isAllEmptyRow).length;

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
        const missing = [];
        if (!account) missing.push("Missing Bank Account");
        if (!branch) missing.push("Missing Branch Code");
        exceptions.push({
          "COMPANY NAME": supplier["COMPANY NAME"],
          "AMOUNT": money(amount),
          "MODE": "BANK",
          "MONTH": month,
          "YEAR": String(year),
          "ISSUE": missing.join(", ")
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

      if (!momoProvider || !momoNumber || !momoNames) {
        const missing = [];
        if (!momoProvider) missing.push("Missing MOMO Provider");
        if (!momoNumber) missing.push("Missing MOMO Number");
        if (!momoNames) missing.push("Missing MOMO Names");
        exceptions.push({
          "COMPANY NAME": supplier["COMPANY NAME"],
          "AMOUNT": money(amount),
          "MODE": "MOMO",
          "MONTH": month,
          "YEAR": String(year),
          "ISSUE": missing.join(", ")
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
    if (f.name.startsWith("exceptions") || f.name.startsWith("invalid-rows")) {
      btn.classList.add("btn-danger");
    } else {
      btn.classList.add("btn-success");
    }
    btn.addEventListener("click", () => downloadText(f.name, f.csv));
    li.appendChild(btn);
    ul.appendChild(li);
  }

  $("exceptions").innerHTML = exceptions.length
    ? `<b>${exceptions.length}</b> issue(s) found. Download <b>exceptions${periodSuffix}.csv</b> and fix <code>data/banking-details.csv</code>.`
    : "No exceptions. All good.";

  if (exceptions.length) {
    $("exceptions")?.classList.add("result-danger");
  } else {
    $("exceptions")?.classList.remove("result-danger");
  }

  // --- Verification summary
  const bankCount = bankBatch.length;
  const momoCount = momoBatch.length;
  const excCount = exceptions.length;
  const invalidCount = invalidRows.length;
  const ledgerAdd = ledgerNew.length;
  const bankTotal = bankBatch.reduce((a, r) => a + (parseAmount(r["AMOUNT"]) || 0), 0);
  const momoTotal = momoBatch.reduce((a, r) => a + (parseAmount(r["AMOUNT"]) || 0), 0);
  const excTotal = exceptions.reduce((a, r) => a + (parseAmount(r["AMOUNT"]) || 0), 0);
  const totalPayable = bankTotal + momoTotal + excTotal;

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
    ["BANK amount total", money(bankTotal)],
    ["MOMO amount total", money(momoTotal)],
    ["Exceptions amount total", money(excTotal)],
    ["Total payable (BANK+MOMO+Exceptions)", money(totalPayable)],
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
    `Ledger added: ${ledgerAdd}. ` +
    `Totals — Bank: ${money(bankTotal)}, MoMo: ${money(momoTotal)}, Exceptions: ${money(excTotal)}, Total: ${money(totalPayable)}.`;
});
