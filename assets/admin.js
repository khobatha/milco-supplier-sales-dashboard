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

$("processBtn").addEventListener("click", async () => {
  $("status").textContent = "Loading banking details + ledger...";
  $("downloads").innerHTML = "";
  $("exceptions").textContent = "";

  const file = $("salesFile").files?.[0];
  if (!file) {
    $("status").textContent = "Please choose a monthly sales-report CSV first.";
    return;
  }

  const threshold = Number($("threshold").value ?? 400);

  const [banking, ledgerExisting, sales] = await Promise.all([
    parseCsvUrl(BANKING_DETAILS_URL),
    parseCsvUrl(LEDGER_URL).catch(() => []),
    parseCsvFile(file)
  ]);

  // Index supplier master data by normalized name
  const bankMap = new Map();
  banking.forEach(r => bankMap.set(normName(r["COMPANY NAME"]), r));

  // Process sales rows
  const bankBatch = [];
  const momoBatch = [];
  const exceptions = [];

  // ledger rows to append
  const ledgerNew = [];

  for (const r of sales) {
    const rawName = r["COMPANY NAME"];
    const nameKey = normName(rawName);
    const amount = Number(r["SUM of COST"]);
    const comment = (r["COMMENT"] ?? "").toString().trim();

    if (!nameKey || !Number.isFinite(amount)) continue;

    const supplier = bankMap.get(nameKey);

    const mode = (amount >= threshold) ? "BANK" : "MOMO";

    if (!supplier) {
      exceptions.push({
        "COMPANY NAME": rawName,
        "AMOUNT": money(amount),
        "MODE": mode,
        "ISSUE": "Supplier not found in banking-details.csv"
      });
      continue;
    }

    if (mode === "BANK") {
      const account = (supplier["ACCOUNT"] ?? "").toString().trim();
      const branch = (supplier["BRANCH"] ?? "").toString().trim();

      if (!account || !branch) {
        exceptions.push({
          "COMPANY NAME": supplier["COMPANY NAME"],
          "AMOUNT": money(amount),
          "MODE": "BANK",
          "ISSUE": "Missing ACCOUNT or BRANCH"
        });
      } else {
        bankBatch.push({
          "NAME": supplier["COMPANY NAME"],
          "ACCOUNT": account,
          "BRANCH": branch,
          "AMOUNT": money(amount),
          "COMMENT": comment
        });
      }
    } else {
      const momoProvider = (supplier["MOMO"] ?? "").toString().trim();
      const momoNumber = (supplier["MOMO NUMBER"] ?? "").toString().trim();
      const momoNames  = (supplier["MOMO NAMES"] ?? "").toString().trim();

      if (!momoProvider || !momoNumber) {
        exceptions.push({
          "COMPANY NAME": supplier["COMPANY NAME"],
          "AMOUNT": money(amount),
          "MODE": "MOMO",
          "ISSUE": "Missing MOMO or MOMO NUMBER"
        });
      } else {
        momoBatch.push({
          "NAME": supplier["COMPANY NAME"],
          "MOMO PROVIDER": momoProvider,
          "MOMO NUMBER": momoNumber,
          "MOMO NAMES": momoNames,
          "AMOUNT": money(amount),
          "COMMENT": comment
        });
      }
    }

    // Always write ledger entry (even if exception? you can decide)
    // Here: only record successful batches
    const ok = (mode === "BANK")
      ? bankBatch.some(x => normName(x.NAME) === nameKey && Number(x.AMOUNT) === Number(money(amount)))
      : momoBatch.some(x => normName(x.NAME) === nameKey && Number(x.AMOUNT) === Number(money(amount)));

    if (ok) {
      // Extract month label from comment like "MiLCo October 2025 Sales"
      // If not found, store comment as period
      const period = comment;

      ledgerNew.push({
        "PERIOD": period,
        "COMPANY NAME": supplier["COMPANY NAME"],
        "AMOUNT": money(amount),
        "MODE": mode,
        "REFERENCE": comment
      });
    }
  }

  // Build updated ledger
  const ledgerUpdated = [...ledgerExisting, ...ledgerNew];

  // Create downloadable files
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
    csv: toCsv(ledgerUpdated, ["PERIOD","COMPANY NAME","AMOUNT","MODE","REFERENCE"])
  });

  if (exceptions.length) {
    files.push({
      name: "exceptions.csv",
      csv: toCsv(exceptions, ["COMPANY NAME","AMOUNT","MODE","ISSUE"])
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
