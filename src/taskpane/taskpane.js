/* global console, document, Office, Excel */

// ── CONFIG ────────────────────────────────────────────────────
const START_FOLDER_PATH          = "BailBonds/EmployeeSubmissions";
const START_FOLDER_PATH_PAYMENTS = "BailBonds/Payments";

// Column indices for secretary daily ledger (Sheet1)
const LEDGER_CLIENT_COL   = 0;
const LEDGER_AMOUNT_COL   = 1;
const LEDGER_BONDSMAN_COL = 5;

// Column indices for PaymentsTable in owner workbook
const PT_CLIENT       = 0;
const PT_TTL_BOND     = 1;
const PT_AMT_CHARGED  = 2;
const PT_AMT_COLLECT  = 3;
const PT_EXPENSE      = 4;
const PT_PCT_PAID     = 5;
const PT_START_BAL    = 6;
const PT_AMT_PAYMENT  = 7;
const PT_END_BAL      = 8;
const PT_BAL_OWED     = 9;
const PT_ENDING_BAL   = 10;
const PT_PAYMENT      = 11;
const PT_REACH        = 12;

// ── STATE ─────────────────────────────────────────────────────
let accessToken        = null;
let selectedFileId     = null;
let selectedPaymentsFileId = null;
let pendingImportRows  = [];
let pendingPaymentRows = [];

// Each browser gets its own breadcrumb stack
const browsers = {
  submission: {
    stackId:    "breadcrumb",
    browserId:  "file-browser",
    stack:      [{ id: "root", name: "My Files" }],
    onSelect:   onSubmissionFileSelected,
  },
  payments: {
    stackId:    "payments-breadcrumb",
    browserId:  "payments-file-browser",
    stack:      [{ id: "root", name: "My Files" }],
    onSelect:   onPaymentsFileSelected,
  },
};

// ── OFFICE INIT ───────────────────────────────────────────────
Office.onReady(() => {
  document.getElementById("sign-in-btn").onclick        = signIn;
  document.getElementById("sign-out-btn").onclick       = signOut;
  document.getElementById("preview-btn").onclick        = previewImport;
  document.getElementById("confirm-import-btn").onclick = confirmImport;
  document.getElementById("cancel-import-btn").onclick  = cancelImport;
  document.getElementById("preview-payments-btn").onclick   = previewPaymentsImport;
  document.getElementById("confirm-payments-btn").onclick   = confirmPaymentsImport;
  document.getElementById("cancel-payments-btn").onclick    = cancelPaymentsImport;
});

// ── AUTH ─────────────────────────────────────────────────────
function signIn() {
  setStatus("Opening sign in window...");
  Office.context.ui.displayDialogAsync(
    "https://cordusbailey02.github.io/AddinTest/dialog.html",
    { height: 60, width: 30, promptBeforeOpen: false },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        setStatus("Failed to open dialog: " + result.error.message);
        return;
      }
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (msg) => {
        dialog.close();
        try {
          const data = JSON.parse(msg.message);
          if (data.status === "success") {
            accessToken = data.accessToken;
            showMainSection(data.userName);
            setStatus("");
            await navigateBrowserToFolder("submission", START_FOLDER_PATH);
            await navigateBrowserToFolder("payments",  START_FOLDER_PATH_PAYMENTS);
          } else {
            setStatus("Sign in error: " + data.message);
          }
        } catch (e) {
          setStatus("Error parsing auth response.");
        }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, (evt) => {
        if (evt.error === 12006) setStatus("Sign in was cancelled.");
      });
    }
  );
}

function signOut() {
  accessToken            = null;
  selectedFileId         = null;
  selectedPaymentsFileId = null;
  pendingImportRows      = [];
  pendingPaymentRows     = [];

  // Reset both browser stacks
  browsers.submission.stack = [{ id: "root", name: "My Files" }];
  browsers.payments.stack   = [{ id: "root", name: "My Files" }];

  document.getElementById("sign-in-section").style.display  = "block";
  document.getElementById("main-section").style.display     = "none";
  document.getElementById("preview-section").style.display  = "none";
  document.getElementById("preview-payments-section").style.display = "none";
  setStatus("");
  setPaymentsStatus("");
}

// ── GENERIC FILE BROWSER ──────────────────────────────────────

/**
 * Navigate a browser to a named OneDrive path on startup.
 * Falls back to root if the folder doesn't exist.
 * @param {"submission"|"payments"} browserKey
 * @param {string} folderPath  e.g. "BailBonds/EmployeeSubmissions"
 */
async function navigateBrowserToFolder(browserKey, folderPath) {
  const b = browsers[browserKey];
  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/root:/${folderPath}`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!response.ok) throw new Error("Folder not found");
    const folder = await response.json();
    b.stack = [
      { id: "root",      name: "My Files"  },
      { id: folder.id,   name: folder.name },
    ];
    await browseFolder(browserKey, folder.id);
  } catch {
    b.stack = [{ id: "root", name: "My Files" }];
    await browseFolder(browserKey, "root");
  }
}

/**
 * Fetch and render folder contents into a browser panel.
 * @param {"submission"|"payments"} browserKey
 * @param {string} folderId  Graph item ID or "root"
 */
async function browseFolder(browserKey, folderId) {
  const b       = browsers[browserKey];
  const browser = document.getElementById(b.browserId);

  browser.innerHTML = "<div style='padding:12px;color:gray;'>Loading...</div>";

  // Reset selection state for this browser
  if (browserKey === "submission") {
    selectedFileId = null;
    document.getElementById("selected-file").style.display    = "none";
    document.getElementById("preview-btn").style.display      = "none";
    document.getElementById("preview-section").style.display  = "none";
  } else {
    selectedPaymentsFileId = null;
    document.getElementById("selected-payments-file").style.display         = "none";
    document.getElementById("preview-payments-btn").style.display           = "none";
    document.getElementById("preview-payments-section").style.display       = "none";
  }

  try {
    const url = folderId === "root"
      ? "https://graph.microsoft.com/v1.0/me/drive/root/children"
      : `https://graph.microsoft.com/v1.0/me/drive/items/${folderId}/children`;

    const response = await fetch(
      url + "?$orderby=name&$select=id,name,folder,file,size",
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    if (!response.ok) {
      const err = await response.json();
      browser.innerHTML = `<div style='padding:12px;color:red;'>${err.error?.message || "Error loading files"}</div>`;
      return;
    }

    const { value: items } = await response.json();

    if (!items.length) {
      browser.innerHTML = "<div style='padding:12px;color:gray;'>This folder is empty.</div>";
      return;
    }

    browser.innerHTML = "";

    items.forEach((item) => {
      const isFolder = !!item.folder;
      const isExcel  = item.name?.endsWith(".xlsx") || item.name?.endsWith(".xls");
      if (!isFolder && !isExcel) return;

      const div = document.createElement("div");
      div.className = "browser-item";
      div.innerHTML = `
        <span class="icon">${isFolder ? "📁" : "📗"}</span>
        <span class="name">${item.name}</span>
        ${isFolder ? '<span style="color:#aaa;font-size:11px;">▶</span>' : ""}
      `;

      div.onclick = () => {
        if (isFolder) {
          b.stack.push({ id: item.id, name: item.name });
          renderBreadcrumb(browserKey);
          browseFolder(browserKey, item.id);
        } else {
          // Deselect others in this browser
          browser.querySelectorAll(".browser-item").forEach(el => el.classList.remove("selected"));
          div.classList.add("selected");
          b.onSelect(item);
        }
      };

      browser.appendChild(div);
    });

    renderBreadcrumb(browserKey);

  } catch (error) {
    browser.innerHTML = `<div style='padding:12px;color:red;'>Error: ${error.message}</div>`;
  }
}

/**
 * Render breadcrumb trail for a browser panel.
 * @param {"submission"|"payments"} browserKey
 */
function renderBreadcrumb(browserKey) {
  const b   = browsers[browserKey];
  const el  = document.getElementById(b.stackId);

  el.innerHTML = b.stack.map((crumb, index) => {
    if (index === b.stack.length - 1) return `📁 ${crumb.name}`;
    return `<span data-index="${index}">${crumb.name}</span> › `;
  }).join("");

  el.querySelectorAll("span[data-index]").forEach((span) => {
    span.onclick = () => {
      const index = parseInt(span.getAttribute("data-index"));
      b.stack = b.stack.slice(0, index + 1);
      browseFolder(browserKey, b.stack[b.stack.length - 1].id);
    };
  });
}

// ── FILE SELECTION CALLBACKS ──────────────────────────────────
function onSubmissionFileSelected(item) {
  selectedFileId = item.id;
  document.getElementById("selected-file-name").textContent = item.name;
  document.getElementById("selected-file").style.display    = "block";
  document.getElementById("preview-btn").style.display      = "inline-block";
  setStatus("");
}

function onPaymentsFileSelected(item) {
  selectedPaymentsFileId = item.id;
  document.getElementById("selected-payments-file-name").textContent = item.name;
  document.getElementById("selected-payments-file").style.display    = "block";
  document.getElementById("preview-payments-btn").style.display      = "inline-block";
  setPaymentsStatus("");
}

// ── GET EMPLOYEE INITIALS FROM OPEN WORKBOOK ──────────────────
async function getEmployeeInitials() {
  return await Excel.run(async (context) => {
    const sheet    = context.workbook.worksheets.getItem("PAYMENTS");
    const nameCell = sheet.getRange("B2");
    nameCell.load("values");
    await context.sync();

    const fullName = nameCell.values[0][0];
    if (!fullName) throw new Error("Could not read employee name from B2 on PAYMENTS sheet.");

    // "LASTNAME, FIRSTNAME" → "LF"
    const parts        = String(fullName).split(",").map(p => p.trim());
    const lastInitial  = parts[0]?.[0]?.toUpperCase() || "";
    const firstInitial = parts[1]?.[0]?.toUpperCase() || "";

    return {
      initials: lastInitial + firstInitial,
      fullName:  String(fullName).trim(),
    };
  });
}

// ── READ SUBMISSION FILE ──────────────────────────────────────
async function readSubmissionData(fileId) {
  const response = await graphFetch(
    `items/${fileId}/workbook/worksheets/Sheet1/usedRange`
  );
  const rows     = response.values;
  const dataRows = rows.slice(3); // skip title, dates, headers

  return dataRows.reduce((bonds, row) => {
    const lastName  = row[2];
    const firstName = row[3];
    const ttlBond   = row[4];
    const amtChgd   = row[6];
    const amtCol    = row[7];
    const expense   = row[8];

    if (!lastName || !ttlBond) return bonds;
    if (String(lastName).toUpperCase().includes("TOTAL")) return bonds;

    bonds.push({
      client:       `${String(lastName).toUpperCase()}, ${String(firstName).toUpperCase()}`,
      ttlBond:      Number(ttlBond) || 0,
      amtCharged:   Number(amtChgd) || 0,
      amtCollected: Number(amtCol)  || 0,
      expense:      Number(expense) || 0,
      startBalance: (Number(amtChgd) || 0) - (Number(amtCol) || 0),
    });

    return bonds;
  }, []);
}

// ── READ SECRETARY LEDGER FILE ────────────────────────────────
async function readSecretaryPayments(fileId, employeeInitials) {
  const response = await graphFetch(
    `items/${fileId}/workbook/worksheets/Sheet1/usedRange`
  );
  const rows     = response.values;
  const dataRows = rows.slice(2); // skip date row and header row

  return dataRows.reduce((matched, row) => {
    const client   = String(row[LEDGER_CLIENT_COL]   ?? "").trim();
    const amount   = row[LEDGER_AMOUNT_COL];
    const bondsman = String(row[LEDGER_BONDSMAN_COL] ?? "").trim().toUpperCase();

    if (!client) return matched;
    if (amount === null || amount === "" || amount === undefined) return matched;
    if (bondsman !== employeeInitials.toUpperCase()) return matched;

    matched.push({
      client:   client.toUpperCase(),
      amount:   Number(amount) || 0,
      bondsman,
    });

    return matched;
  }, []);
}

// ── GRAPH API HELPER ──────────────────────────────────────────
async function graphFetch(path) {
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/${path}`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  if (!response.ok) {
    const err = await response.json();
    throw new Error(err.error?.message || `Graph request failed: ${path}`);
  }
  return response.json();
}

// ── PREVIEW BOND IMPORT ───────────────────────────────────────
async function previewImport() {
  if (!selectedFileId) {
    setStatus("Please select a submission file first.");
    return;
  }
  setStatus("Reading submission...");
  document.getElementById("preview-section").style.display = "none";

  try {
    const bonds = await readSubmissionData(selectedFileId);
    if (!bonds.length) {
      setStatus("No bond entries found in the submission file.");
      return;
    }

    pendingImportRows = bonds;

    const tbody = document.getElementById("preview-tbody");
    tbody.innerHTML = bonds.map(bond => `
      <tr>
        <td>${bond.client}</td>
        <td>$${bond.ttlBond.toLocaleString()}</td>
        <td>$${bond.amtCharged.toLocaleString()}</td>
        <td>$${bond.amtCollected.toLocaleString()}</td>
        <td>$${bond.expense.toLocaleString()}</td>
        <td>$${bond.startBalance.toLocaleString()}</td>
      </tr>
    `).join("");

    document.getElementById("import-summary").style.display   = "none";
    document.getElementById("preview-section").style.display  = "block";
    document.getElementById("confirm-import-btn").disabled    = false;
    setStatus(`${bonds.length} bond${bonds.length > 1 ? "s" : ""} ready to import.`);

  } catch (error) {
    setStatus("Error reading submission: " + error.message);
    console.error(error);
  }
}

function cancelImport() {
  pendingImportRows = [];
  document.getElementById("preview-section").style.display = "none";
  setStatus("Import cancelled.");
}

// ── CONFIRM BOND IMPORT ───────────────────────────────────────
async function confirmImport() {
  if (!pendingImportRows.length) return;
  document.getElementById("confirm-import-btn").disabled = true;
  setStatus("Importing...");

  try {
    await Excel.run(async (context) => {
      const table     = context.workbook.tables.getItem("PaymentsTable");
      const bodyRange = table.getDataBodyRange();
      bodyRange.load("values, rowCount, columnCount");
      await context.sync();

      const colCount        = bodyRange.columnCount;
      const currentRowCount = bodyRange.rowCount;

      // Compact — remove empty rows
      const compacted = bodyRange.values.filter(row =>
        row[PT_CLIENT] !== null && row[PT_CLIENT] !== "" && row[PT_CLIENT] !== undefined
      );

      // Build new rows
      const newRows = pendingImportRows.map(bond => {
        const row = new Array(colCount).fill("");
        row[PT_CLIENT]      = bond.client;
        row[PT_TTL_BOND]    = bond.ttlBond;
        row[PT_AMT_CHARGED] = bond.amtCharged;
        row[PT_AMT_COLLECT] = bond.amtCollected;
        row[PT_EXPENSE]     = bond.expense;
        row[PT_START_BAL]   = bond.startBalance;
        return row;
      });

      const allRows        = [...compacted, ...newRows];
      const neededRowCount = allRows.length;

      // Grow table if needed
      if (neededRowCount > currentRowCount) {
        for (let i = 0; i < neededRowCount - currentRowCount; i++) {
          table.rows.add(null, [new Array(colCount).fill("")]);
        }
        await context.sync();
      }

      // Write all rows back
      const freshBody = table.getDataBodyRange();
      freshBody.load("rowCount");
      await context.sync();

      freshBody.getCell(0, 0)
        .getResizedRange(allRows.length - 1, colCount - 1)
        .values = allRows;

      // Write formulas into new rows only
      const newRowStart = compacted.length;
      for (let i = 0; i < pendingImportRows.length; i++) {
        const ri = newRowStart + i;
        freshBody.getCell(ri, PT_END_BAL).formulas    = [["=[@[START BALANCE]]-[@[AMT OF PAYMENT]]"]];
        freshBody.getCell(ri, PT_ENDING_BAL).formulas = [["=[@[BALANCE OWED]]-[@[PAYMENT]]"]];
        freshBody.getCell(ri, PT_PAYMENT).formulas    = [["=[@[AMT OF PAYMENT]]*[@[% PAID ON BOND]]"]];
      }

      // Clear leftover rows if table shrank
      if (freshBody.rowCount > neededRowCount) {
        freshBody.getCell(neededRowCount, 0)
          .getResizedRange(freshBody.rowCount - neededRowCount - 1, colCount - 1)
          .clear("Contents");
      }

      await context.sync();
    });

    showSummary("import-summary", `✔ Successfully imported ${pendingImportRows.length} bond${pendingImportRows.length > 1 ? "s" : ""} into PaymentsTable.`, false);
    pendingImportRows = [];
    document.getElementById("confirm-import-btn").disabled = false;
    setStatus("");

  } catch (error) {
    showSummary("import-summary", `✖ Import failed: ${error.message}`, true);
    document.getElementById("confirm-import-btn").disabled = false;
    console.error(error);
  }
}

// ── PREVIEW PAYMENTS IMPORT ───────────────────────────────────
async function previewPaymentsImport() {
  if (!selectedPaymentsFileId) {
    setPaymentsStatus("Please select a daily ledger file first.");
    return;
  }
  setPaymentsStatus("Reading daily ledger...");
  document.getElementById("preview-payments-section").style.display = "none";

  try {
    const { initials, fullName } = await getEmployeeInitials();
    document.getElementById("detected-employee").textContent    = `${fullName} (${initials})`;
    document.getElementById("employee-badge").style.display     = "block";

    const matched = await readSecretaryPayments(selectedPaymentsFileId, initials);
    if (!matched.length) {
      setPaymentsStatus(`No payments found for bondsman "${initials}" in this file.`);
      return;
    }

    // Load PaymentsTable rows for cross-referencing
    const tableRows = await Excel.run(async (context) => {
      const table     = context.workbook.tables.getItem("PaymentsTable");
      const bodyRange = table.getDataBodyRange();
      bodyRange.load("values");
      await context.sync();
      return bodyRange.values;
    });

    const tableClients = tableRows.map((row, idx) => ({
      rowIndex:    idx,
      client:      String(row[PT_CLIENT]    || "").toUpperCase().trim(),
      startBalance: row[PT_START_BAL],
    }));

    const previews   = [];
    const unmatched  = [];

    for (const payment of matched) {
      const match = tableClients.find(r => r.client === payment.client && r.client !== "");
      if (match) {
        previews.push({
          client:       payment.client,
          amount:       payment.amount,
          rowIndex:     match.rowIndex,
          startBalance: match.startBalance,
          projectedEnd: (Number(match.startBalance) || 0) - payment.amount,
        });
      } else {
        unmatched.push(payment.client);
      }
    }

    pendingPaymentRows = previews;

    const tbody = document.getElementById("preview-payments-tbody");
    tbody.innerHTML = previews.map(row => `
      <tr>
        <td>${row.client}</td>
        <td>$${Number(row.startBalance || 0).toLocaleString()}</td>
        <td>$${row.amount.toLocaleString()}</td>
        <td>$${row.projectedEnd.toLocaleString()}</td>
      </tr>
    `).join("");

    const warningEl = document.getElementById("payments-unmatched-warning");
    if (unmatched.length) {
      warningEl.textContent  = `⚠ ${unmatched.length} payment(s) had no match in PaymentsTable: ${unmatched.join(", ")}`;
      warningEl.style.display = "block";
    } else {
      warningEl.style.display = "none";
    }

    document.getElementById("import-payments-summary").style.display    = "none";
    document.getElementById("preview-payments-section").style.display   = "block";
    document.getElementById("confirm-payments-btn").disabled             = !previews.length;

    setPaymentsStatus(`${previews.length} payment${previews.length !== 1 ? "s" : ""} matched and ready to apply.`);

  } catch (error) {
    setPaymentsStatus("Error: " + error.message);
    console.error(error);
  }
}

function cancelPaymentsImport() {
  pendingPaymentRows = [];
  document.getElementById("preview-payments-section").style.display = "none";
  setPaymentsStatus("Payment import cancelled.");
}

// ── CONFIRM PAYMENTS IMPORT ───────────────────────────────────
async function confirmPaymentsImport() {
  if (!pendingPaymentRows.length) return;
  document.getElementById("confirm-payments-btn").disabled = true;
  setPaymentsStatus("Applying payments...");

  try {
    await Excel.run(async (context) => {
      const table     = context.workbook.tables.getItem("PaymentsTable");
      const bodyRange = table.getDataBodyRange();
      bodyRange.load("values, rowCount, columnCount");
      await context.sync();

      for (const payment of pendingPaymentRows) {
        const ri = payment.rowIndex;

        // Write payment amount into AMT OF PAYMENT
        bodyRange.getCell(ri, PT_AMT_PAYMENT).values = [[payment.amount]];
        await context.sync();

        // Read recalculated END BALANCE
        const endBalCell = bodyRange.getCell(ri, PT_END_BAL);
        endBalCell.load("values");
        await context.sync();

        const newEndBalance = endBalCell.values[0][0];

        // Write END BALANCE → START BALANCE, then clear AMT OF PAYMENT
        bodyRange.getCell(ri, PT_START_BAL).values   = [[newEndBalance]];
        bodyRange.getCell(ri, PT_AMT_PAYMENT).values = [[""]];
        await context.sync();
      }
    });

    showSummary(
      "import-payments-summary",
      `✔ Successfully applied ${pendingPaymentRows.length} payment${pendingPaymentRows.length !== 1 ? "s" : ""} and updated START BALANCE.`,
      false
    );
    pendingPaymentRows = [];
    document.getElementById("confirm-payments-btn").disabled = false;
    setPaymentsStatus("");

  } catch (error) {
    showSummary("import-payments-summary", `✖ Payment import failed: ${error.message}`, true);
    document.getElementById("confirm-payments-btn").disabled = false;
    console.error(error);
  }
}

// ── UI HELPERS ────────────────────────────────────────────────
function showMainSection(userName) {
  document.getElementById("sign-in-section").style.display = "none";
  document.getElementById("main-section").style.display    = "block";
  document.getElementById("user-name").textContent         = userName;
}

function setStatus(message) {
  document.getElementById("status").textContent = message;
}

function setPaymentsStatus(message) {
  document.getElementById("payments-status").textContent = message;
}

function showSummary(elementId, message, isError) {
  const el       = document.getElementById(elementId);
  el.className   = isError ? "import-summary error" : "import-summary";
  el.textContent = message;
  el.style.display = "block";
}