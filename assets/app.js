// assets/app.js
// MiLCo Supplier Sales Dashboard (Client / Read-only)

const BANKING_DETAILS_URL = "data/banking-details.csv";
const LEDGER_URL = "data/transactions.csv";

const $ = (id) => document.getElementById(id);

function normName(s) {
  return (s ?? "")
    .toString()
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
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

function safeText(s) {
  return (s ?? "").toString().trim();
}

let banking = [];
let ledger = [];

(async function init() {
  try {
    [banking, ledger] = await Promise.all([
      parseCsvUrl(BANKING_DETAILS_URL),
      parseCsvUrl(LEDGER_URL).catch(() => [])
    ]);

    // KPIs
    const periods = [...new Set(ledger.map(r => safeText(r.PERIOD)).filter(Boolean))];
    const totalSales = ledger.reduce((a, r) => a + (Number(r.AMOUNT) || 0), 0);

    const modeTotals = ledger.reduce((acc, r) => {
      const m = safeText(r.MODE).toUpperCase() || "UNKNOWN";
      acc[m] = (acc[m] || 0) + (Number(r.AMOUNT) || 0);
      return acc;
    }, {});

    $("kpiMonths").textContent = String(periods.length);
    $("kpiSales").textContent = `M ${fmtMoney(totalSales)}`;
    $("kpiPayouts").textContent = `M ${fmtMoney(totalSales)}`;
    $("kpiModes").textContent =
      `Bank M ${fmtMoney(modeTotals.BANK || 0)} / MoMo M ${fmtMoney(modeTotals.MOMO || 0)}`;

    // Charts
    renderSalesByMonthChart(ledger);
    renderModeChart(modeTotals);

    // Latest tx table
    renderLatestTable(ledger);

    // Supplier search
    setupSupplierSearch(banking, ledger);
  } catch (e) {
    console.error(e);
    // Fail softly in UI
    const msg = "Failed to load data files. Check that you are running a local server and that /data/*.csv exists.";
    if ($("supplierResult")) $("supplierResult").textContent = msg;
  }
})();

function renderSalesByMonthChart(ledgerRows) {
  const byPeriod = {};
  ledgerRows.forEach(r => {
    const p = safeText(r.PERIOD);
    if (!p) return;
    byPeriod[p] = (byPeriod[p] || 0) + (Number(r.AMOUNT) || 0);
  });

  const labels = Object.keys(byPeriod);
  const data = labels.map(k => byPeriod[k]);

  const el = document.getElementById("salesChart");
  if (!el || typeof Chart === "undefined") return;

  new Chart(el, {
    type: "bar",
    data: {
      labels,
      datasets: [{ label: "Total payouts", data }]
    },
    options: {
      responsive: true,
      plugins: { legend: { display: true } }
    }
  });
}

function renderModeChart(modeTotals) {
  const labels = Object.keys(modeTotals);
  const data = labels.map(k => modeTotals[k]);

  const el = document.getElementById("modeChart");
  if (!el || typeof Chart === "undefined") return;

  new Chart(el, {
    type: "doughnut",
    data: {
      labels,
      datasets: [{ label: "Payouts by mode", data }]
    },
    options: { responsive: true }
  });
}

function renderLatestTable(ledgerRows) {
  const tbody = document.querySelector("#txTable tbody");
  if (!tbody) return;

  // show last 30 entries (most recent at top)
  const rows = [...ledgerRows].slice(-30).reverse();

  tbody.innerHTML = "";
  rows.forEach(r => {
    const tr = document.createElement("tr");

    const period = safeText(r.PERIOD);
    const supplier = safeText(r["COMPANY NAME"]);
    const amount = Number(r.AMOUNT) || 0;

    const mode = safeText(r.MODE).toUpperCase();
    const badgeClass = mode === "BANK" ? "red" : (mode === "MOMO" ? "green" : "");

    const reference = safeText(r.REFERENCE);

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

  // Build supplier index
  const supplierIndex = new Map();
  bankingRows.forEach(r => supplierIndex.set(normName(r["COMPANY NAME"]), r));

  // Pre-compute ledger by supplier for speed
  const ledgerBySupplier = new Map();
  ledgerRows.forEach(r => {
    const k = normName(r["COMPANY NAME"]);
    if (!k) return;
    if (!ledgerBySupplier.has(k)) ledgerBySupplier.set(k, []);
    ledgerBySupplier.get(k).push(r);
  });

  input.addEventListener("input", () => {
    const qRaw = input.value;
    const q = normName(qRaw);

    if (!q) {
      out.innerHTML = "";
      return;
    }

    const keys = [...supplierIndex.keys()];

    // Simple best-match: exact, then startsWith, then contains
    let matchKey =
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

    // mode totals for supplier
    const supplierModeTotals = history.reduce((acc, r) => {
      const m = safeText(r.MODE).toUpperCase() || "UNKNOWN";
      acc[m] = (acc[m] || 0) + (Number(r.AMOUNT) || 0);
      return acc;
    }, {});

    out.innerHTML = `
      <div><b>${safeText(s["COMPANY NAME"])}</b></div>
      <div class="muted small" style="margin-top:6px;">
        <span class="badge red">Bank</span>
        ${safeText(s["BANK"]) || "—"} · Acc: ${safeText(s["ACCOUNT"]) || "—"} · Branch: ${safeText(s["BRANCH"]) || "—"}
      </div>
      <div class="muted small" style="margin-top:6px;">
        <span class="badge green">MoMo</span>
        ${safeText(s["MOMO"]) || "—"} · No: ${safeText(s["MOMO NUMBER"]) || "—"} · Names: ${safeText(s["MOMO NAMES"]) || "—"}
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
