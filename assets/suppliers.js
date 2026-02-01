// assets/suppliers.js
// Supplier payment details management

const BANKING_DETAILS_URL = "data/banking-details.csv";

const $ = (id) => document.getElementById(id);

let currentBankingRows = [];

function clean(s) {
  return (s ?? "").toString().trim();
}

function upper(s) {
  return clean(s).toUpperCase();
}

function normName(s) {
  return (s ?? "")
    .toString()
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}

function normHeader(s) {
  return clean(s).toUpperCase().replace(/\s+/g, " ");
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

function downloadXlsx(filename, rows, sheetName) {
  const wb = XLSX.utils.book_new();
  const headers = [
    "COMPANY NAME", "BANK", "ACCOUNT", "BRANCH",
    "MOMO", "MOMO NUMBER", "MOMO NAMES"
  ];
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

function isAllEmptyArrayRow(arr) {
  return !(arr || []).some((v) => clean(v) !== "");
}

function findHeaderRow(rows) {
  const required = ["COMPANY NAME", "COST"];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const map = new Map();
    row.forEach((cell, idx) => {
      const key = normHeader(cell);
      if (key) map.set(key, idx);
    });
    const hasRequired = required.every((h) => map.has(h));
    if (hasRequired) {
      return { rowIndex: i, map };
    }
  }
  return null;
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

function normalizeBankingRows(rows) {
  const out = [];
  for (const r of rows || []) {
    const name = clean(r["COMPANY NAME"] ?? r["COMPANY"] ?? r["NAME"]);
    if (!name) continue;
    out.push({
      "COMPANY NAME": upper(name),
      "BANK": upper(r["BANK"] ?? r["BANK NAME"]),
      "ACCOUNT": clean(r["ACCOUNT"] ?? r["BANK ACCOUNT/MOBILE"]),
      "BRANCH": clean(r["BRANCH"]),
      "MOMO": clean(r["MOMO"]),
      "MOMO NUMBER": clean(r["MOMO NUMBER"]),
      "MOMO NAMES": upper(r["MOMO NAMES"])
    });
  }
  return out;
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
    currentBankingRows = normalized.slice();
    renderBankingTable(normalized);
  } catch (err) {
    const tbody = document.querySelector("#bankingTable tbody");
    if (tbody) {
      tbody.innerHTML = `<tr><td colspan="9">Failed to load banking-details.csv</td></tr>`;
    }
  }
}

function exportCurrentBankingCsv() {
  const csv = toCsv(currentBankingRows, [
    "COMPANY NAME",
    "BANK",
    "ACCOUNT",
    "BRANCH",
    "MOMO",
    "MOMO NUMBER",
    "MOMO NAMES"
  ]);
  downloadText("banking-details.csv", csv);
}

function exportCurrentBankingXlsx() {
  downloadXlsx("banking-details.xlsx", currentBankingRows, "BANKING DETAILS");
}

$("exportBankingCsv")?.addEventListener("click", exportCurrentBankingCsv);
$("exportBankingXlsx")?.addEventListener("click", exportCurrentBankingXlsx);
$("refreshBanking")?.addEventListener("click", loadBankingDetails);

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
        "COMPANY NAME": upper(name),
        "BANK": upper(sales["BANK"] || ""),
        "ACCOUNT": clean(sales["ACCOUNT"] || ""),
        "BRANCH": "",
        "MOMO": clean(momo["MOMO"] || ""),
        "MOMO NUMBER": clean(momo["MOMO NUMBER"] || ""),
        "MOMO NAMES": upper(momo["MOMO NAMES"] || "")
      });
      addedCount++;
    }
  }

  const updatedRows = Array.from(bankMap.values()).sort((a, b) => {
    return clean(a["COMPANY NAME"]).localeCompare(clean(b["COMPANY NAME"]));
  });

  updatedRows.forEach((r) => {
    r["COMPANY NAME"] = upper(r["COMPANY NAME"]);
    r["BANK"] = upper(r["BANK"]);
    r["MOMO NAMES"] = upper(r["MOMO NAMES"]);
  });

  currentBankingRows = updatedRows.slice();
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

loadBankingDetails();
