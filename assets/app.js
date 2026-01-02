// assets/app.js
// MiLCo Supplier Sales Dashboard (Client / Read-only)
// Adds filters: Year, Month, Mode + chart: Bank vs MoMo over time

const BANKING_DETAILS_URL = "data/banking-details.csv";
const LEDGER_URL = "data/transactions.csv";

const $ = (id) => document.getElementById(id);

const MONTHS = [
  "January","February","March","April","May","June",
  "July","August","September","October","November","December"
];

function normName(s) {
  return (s ?? "").toString().trim().replace(/\s+/g, " ").toUpperCase();
}
function clean(s) { return (s ?? "").toString().trim(); }

function monthIndex(monthName) {
  const m = clean(monthName);
  const idx = MONTHS.findIndex(x => x.toUpperCase() === m.toUpperCase());
  return idx; // -1 if not found
}

function parseMonthYearFromText(text) {
  const t = clean(text);
  const re = new RegExp(`\\b(${MONTHS.join("|")})\\b\\s+(\\d{4})`, "i");
  const m = t.match(re);
  if (!m) return null;
  return { month: m[1], year: Number(m[2]), period: `${m[1]} ${m[2]}` };
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

function fmtMoney(n) {
  const x = Number(n);
  return Number.isFinite(x)
    ? x.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    : "0.00";
}

// --- Charts instances (so we can update instead of stacking)
let salesChartInstance = null;
let modeChartInstance = null;
let modeOverTimeChartInstance = null;

function normalizeLedgerRows(ledgerRows) {
  const out = [];

  for (const r of (ledgerRows || [])) {
    const company = clean(r["COMPANY NAME"] ?? r["COMPANY"] ?? r["NAME"]);
    const amount = Number(r["AMOUNT"]);
    const mode = clean(r["MODE"]).toUpperCase();
    const reference = clean(r["REFERENCE"] ?? r["COMMENT"] ?? r["PERIOD"] ?? "");

    let month = clean(r["MONTH"]);
    let year = Number(r["YEAR"]);
    let period = clean(r["PERIOD"]);

    // Upgrade from old formats if MONTH/YEAR missing
    if ((!month || !Number.isFinite(year)) && period) {
      const parsed = parseMonthYearFromText(period);
      if (parsed) {
        month = month || parsed.month;
        year = Number.isFinite(year) ? year : parsed.year;
        period = period || parsed.period;
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

    out.push({
      MONTH: month || "",
      YEAR: Number.isFinite(year) ? year : "",
      PERIOD: period || "",
      "COMPANY NAME": company,
      AMOUNT: amount,
      MODE: mode || "",
      REFERENCE: reference || ""
    });
  }

  // Sort stable by year, then month index
  out.sort((a, b) => {
    const ya = Number(a.YEAR) || 0, yb = Number(b.YEAR) || 0;
    if (ya !== yb) return ya - yb;
    const ma = monthIndex(a.MONTH), mb = monthIndex(b.MONTH);
    return (ma === -1 ? 99 : ma) - (mb === -1 ? 99 : mb);
  });

  return out;
}

let banking = [];
let ledger = [];
let ledgerFiltered = [];

(async function init() {
  try {
    [banking, ledger] = await Promise.all([
      parseCsvUrl(BANKING_DETAILS_URL),
      parseCsvUrl(LEDGER_URL).catch(() => [])
    ]);

    ledger = normalizeLedgerRows(ledger);

    // Setup filters
    setupFilters(ledger);

    // Initial render
    applyFiltersAndRender();

    // Supplier search always uses full ledger (but you can make it respect filters if you want)
    setupSupplierSearch(banking, ledger);

  } catch (e) {
    console.error(e);
    const msg = "Failed to load data files. Run via a local server and confirm /data/*.csv exists.";
    if ($("supplierResult")) $("supplierResult").textContent = msg;
  }
})();

function setupFilters(ledgerRows) {
  const yearSel = $("filterYear");
  const monthSel = $("filterMonth");
  const modeSel = $("filterMode");

  if (!yearSel || !monthSel || !modeSel) return;

  // Years
  const years = [...new Set(ledgerRows.map(r => r.YEAR).filter(y => y !== ""))].sort((a,b)=>a-b);
  yearSel.innerHTML = `<option value="ALL">All</option>` + years.map(y => `<option value="${y}">${y}</option>`).join("");

  // Months
  monthSel.innerHTML =
    `<option value="ALL">All</option>` +
    MONTHS.map(m => `<option value="${m}">${m}</option>`).join("");

  // React to changes
  [yearSel, monthSel, modeSel].forEach(el => el.addEventListener("change", applyFiltersAndRender));
}

function applyFiltersAndRender() {
  const year = clean($("filterYear")?.value || "ALL");
  const month = clean($("filterMonth")?.value || "ALL");
  const mode = clean($("filterMode")?.value || "ALL").toUpperCase();

  ledgerFiltered = ledger.filter(r => {
    const matchYear = (year === "ALL") || String(r.YEAR) === String(year);
    const matchMonth = (month === "ALL") || clean(r.MONTH).toUpperCase() === month.toUpperCase();
    const matchMode = (mode === "ALL") || clean(r.MODE).toUpperCase() === mode;
    return matchYear && matchMonth && matchMode;
  });

  // Filter summary
  const summary = [];
  if (year !== "ALL") summary.push(`Year: ${year}`);
  if (month !== "ALL") summary.push(`Month: ${month}`);
  if (mode !== "ALL") summary.push(`Mode: ${mode}`);
  $("filterSummary").textContent = summary.length ? `Active filters → ${summary.join(" · ")}` : "No filters applied (showing everything).";

  // KPIs + charts + table
  renderKPIs(ledgerFiltered);
  renderSalesByPeriodChart(ledgerFiltered);
  renderModeChart(ledgerFiltered);

  // Mode-over-time chart: ignore Mode filter so we always show both BANK and MOMO
  const ledgerForModeOverTime = ledger.filter(r => {
    const matchYear = (year === "ALL") || String(r.YEAR) === String(year);
    const matchMonth = (month === "ALL") || clean(r.MONTH).toUpperCase() === month.toUpperCase();
    return matchYear && matchMonth;
  });
  renderModeOverTimeChart(ledgerForModeOverTime);

  renderLatestTable(ledgerFiltered);
}

function renderKPIs(rows) {
  const periods = [...new Set(rows.map(r => clean(r.PERIOD)).filter(Boolean))];
  const total = rows.reduce((a, r) => a + (Number(r.AMOUNT) || 0), 0);

  const modeTotals = rows.reduce((acc, r) => {
    const m = clean(r.MODE).toUpperCase() || "UNKNOWN";
    acc[m] = (acc[m] || 0) + (Number(r.AMOUNT) || 0);
    return acc;
  }, {});

  $("kpiMonths").textContent = String(periods.length);
  $("kpiSales").textContent = `M ${fmtMoney(total)}`;
  $("kpiPayouts").textContent = `M ${fmtMoney(total)}`;
  $("kpiModes").textContent =
    `Bank M ${fmtMoney(modeTotals.BANK || 0)} / MoMo M ${fmtMoney(modeTotals.MOMO || 0)}`;
}

function renderSalesByPeriodChart(rows) {
  const byPeriod = {};
  rows.forEach(r => {
    const p = clean(r.PERIOD);
    if (!p) return;
    byPeriod[p] = (byPeriod[p] || 0) + (Number(r.AMOUNT) || 0);
  });

  const labels = Object.keys(byPeriod);
  const data = labels.map(k => byPeriod[k]);

  const el = document.getElementById("salesChart");
  if (!el || typeof Chart === "undefined") return;

  if (salesChartInstance) salesChartInstance.destroy();
  salesChartInstance = new Chart(el, {
    type: "bar",
    data: { labels, datasets: [{ label: "Total payouts", data }] },
    options: { responsive: true, plugins: { legend: { display: true } } }
  });
}

function renderModeChart(rows) {
  const modeTotals = rows.reduce((acc, r) => {
    const m = clean(r.MODE).toUpperCase() || "UNKNOWN";
    acc[m] = (acc[m] || 0) + (Number(r.AMOUNT) || 0);
    return acc;
  }, {});

  const labels = Object.keys(modeTotals);
  const data = labels.map(k => modeTotals[k]);

  const el = document.getElementById("modeChart");
  if (!el || typeof Chart === "undefined") return;

  if (modeChartInstance) modeChartInstance.destroy();
  modeChartInstance = new Chart(el, {
    type: "doughnut",
    data: { labels, datasets: [{ label: "Payouts by mode", data }] },
    options: { responsive: true }
  });
}

function renderModeOverTimeChart(rows) {
  // Build period list in chronological order
  const map = new Map(); // period -> {bank, momo}
  for (const r of rows) {
    const period = clean(r.PERIOD) || `${clean(r.MONTH)} ${clean(r.YEAR)}`.trim();
    if (!period) continue;
    if (!map.has(period)) map.set(period, { bank: 0, momo: 0 });
    const rec = map.get(period);
    const amt = Number(r.AMOUNT) || 0;
    const mode = clean(r.MODE).toUpperCase();
    if (mode === "BANK") rec.bank += amt;
    if (mode === "MOMO") rec.momo += amt;
  }

  const labels = [...map.keys()];
  const bankData = labels.map(p => map.get(p).bank);
  const momoData = labels.map(p => map.get(p).momo);

  const el = document.getElementById("modeOverTimeChart");
  if (!el || typeof Chart === "undefined") return;

  if (modeOverTimeChartInstance) modeOverTimeChartInstance.destroy();
  modeOverTimeChartInstance = new Chart(el, {
    type: "line",
    data: {
      labels,
      datasets: [
        { label: "Bank", data: bankData },
        { label: "MoMo", data: momoData }
      ]
    },
    options: {
      responsive: true,
      plugins: { legend: { display: true } }
    }
  });
}

function renderLatestTable(rows) {
  const tbody = document.querySelector("#txTable tbody");
  if (!tbody) return;

  const last = [...rows].slice(-30).reverse();
  tbody.innerHTML = "";

  last.forEach(r => {
    const tr = document.createElement("tr");
    const period = clean(r.PERIOD) || `${clean(r.MONTH)} ${clean(r.YEAR)}`.trim();
    const supplier = clean(r["COMPANY NAME"]);
    const amount = Number(r.AMOUNT) || 0;

    const mode = clean(r.MODE).toUpperCase();
    const badgeClass = mode === "BANK" ? "red" : (mode === "MOMO" ? "green" : "");
    const reference = clean(r.REFERENCE);

    tr.innerHTML = `
      <td>${period}</td>
      <td>${supplier}</td>
      <td>M ${fmtMoney(amount)}</td>
      <td><span class="badge ${badgeClass}">${mode || "—"}</span></td>
      <td>${reference}</td>
    `;
    tbody.appendChild(tr);
  });
}

function setupSupplierSearch(bankingRows, ledgerRows) {
  const input = $("supplierSearch");
  const out = $("supplierResult");
  if (!input || !out) return;

  const supplierIndex = new Map();
  bankingRows.forEach(r => supplierIndex.set(normName(r["COMPANY NAME"]), r));

  const ledgerBySupplier = new Map();
  ledgerRows.forEach(r => {
    const k = normName(r["COMPANY NAME"]);
    if (!k) return;
    if (!ledgerBySupplier.has(k)) ledgerBySupplier.set(k, []);
    ledgerBySupplier.get(k).push(r);
  });

  input.addEventListener("input", () => {
    const q = normName(input.value);
    if (!q) { out.innerHTML = ""; return; }

    const keys = [...supplierIndex.keys()];
    const matchKey =
      keys.find(k => k === q) ||
      keys.find(k => k.startsWith(q)) ||
      keys.find(k => k.includes(q)) ||
      keys.find(k => q.includes(k));

    if (!matchKey) {
      out.innerHTML = "<span class='muted'>No match in supplier master list.</span>";
      return;
    }

    const s = supplierIndex.get(matchKey);
    const history = ledgerBySupplier.get(matchKey) || [];
    const total = history.reduce((a, r) => a + (Number(r.AMOUNT) || 0), 0);

    const supplierModeTotals = history.reduce((acc, r) => {
      const m = clean(r.MODE).toUpperCase() || "UNKNOWN";
      acc[m] = (acc[m] || 0) + (Number(r.AMOUNT) || 0);
      return acc;
    }, {});

    out.innerHTML = `
      <div><b>${clean(s["COMPANY NAME"])}</b></div>

      <div class="muted small" style="margin-top:6px;">
        <span class="badge red">Bank</span>
        ${clean(s["BANK"]) || "—"} · Acc: ${clean(s["ACCOUNT"]) || "—"} · Branch: ${clean(s["BRANCH"]) || "—"}
      </div>

      <div class="muted small" style="margin-top:6px;">
        <span class="badge green">MoMo</span>
        ${clean(s["MOMO"]) || "—"} · No: ${clean(s["MOMO NUMBER"]) || "—"} · Names: ${clean(s["MOMO NAMES"]) || "—"}
      </div>

      <div style="margin-top:10px;">
        <b>Total payouts:</b> M ${fmtMoney(total)}
        <span class="muted small">(${history.length} month entries)</span>
      </div>

      <div class="muted small" style="margin-top:6px;">
        Bank: M ${fmtMoney(supplierModeTotals.BANK || 0)} · MoMo: M ${fmtMoney(supplierModeTotals.MOMO || 0)}
      </div>
    `;
  });
}
