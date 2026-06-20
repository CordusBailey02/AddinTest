/* global console, document, Office, Excel */

// ── CONFIG ────────────────────────────────────────────────────
const START_FOLDER_PATH = "BailBonds/EmployeeSubmissions";
const START_FOLDER_PATH_PAYMENTS = "BailBonds/Payments";

let accessToken = null;
let selectedFileId = null;
let selectedPaymentsFileId = null;
let pendingImportRows = [];
let pendingPaymentRows = [];
let breadcrumbStack = [{ id: "root", name: "My Files" }];
let breadcrumbStackPayments = [{ id: "root", name: "My Files" }];

// ── OFFICE INIT ───────────────────────────────────────────────
Office.onReady(() => {
  document.getElementById("sign-in-btn").onclick = signIn;
  document.getElementById("sign-out-btn").onclick = signOut;

  // Bond import
  document.getElementById("preview-btn").onclick = previewImport;
  document.getElementById("confirm-import-btn").onclick = confirmImport;
  document.getElementById("cancel-import-btn").onclick = cancelImport;

  // Payment import
  document.getElementById("preview-payments-btn").onclick = previewPaymentsImport;
  document.getElementById("confirm-payments-btn").onclick = confirmPaymentsImport;
  document.getElementById("cancel-payments-btn").onclick = cancelPaymentsImport;
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
            await navigateToStartFolder();
            await navigateToStartFolderPayments();
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
  accessToken = null;
  selectedFileId = null;
  selectedPaymentsFileId = null;
  pendingImportRows = [];
  pendingPaymentRows = [];
  breadcrumbStack = [{ id: "root", name: "My Files" }];
  breadcrumbStackPayments = [{ id: "root", name: "My Files" }];
  document.getElementById("sign-in-section").style.display = "block";
  document.getElementById("main-section").style.display = "none";
  document.getElementById("preview-section").style.display = "none";
  document.getElementById("preview-payments-section").style.display = "none";
  setStatus("");
  setPaymentsStatus("");
}

// ── START FOLDERS ─────────────────────────────────────────────
async function navigateToStartFolder() {
  setStatus(`Opening submissions folder...`);
  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/root:/${START_FOLDER_PATH}`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!response.ok) {
      breadcrumbStack = [{ id: "root", name: "My Files" }];
      await browseFolder("root");
      return;
    }
    const folder = await response.json();
    breadcrumbStack = [
      { id: "root", name: "My Files" },
      { id: folder.id, name: folder.name },
    ];
    setStatus("");
    await browseFolder(folder.id);
  } catch (error) {
    breadcrumbStack = [{ id: "root", name: "My Files" }];
    await browseFolder("root");
  }
}

async function navigateToStartFolderPayments() {
  setPaymentsStatus(`Opening payments folder...`);
  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/root:/${START_FOLDER_PATH_PAYMENTS}`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!response.ok) {
      breadcrumbStackPayments = [{ id: "root", name: "My Files" }];
      await browsePaymentsFolder("root");
      return;
    }
    const folder = await response.json();
    breadcrumbStackPayments = [
      { id: "root", name: "My Files" },
      { id: folder.id, name: folder.name },
    ];
    setPaymentsStatus("");
    await browsePaymentsFolder(folder.id);
  } catch (error) {
    breadcrumbStackPayments = [{ id: "root", name: "My Files" }];
    await browsePaymentsFolder("root");
  }
}

// ── FILE BROWSERS ─────────────────────────────────────────────
async function browseFolder(folderId) {
  const browser = document.getElementById("file-browser");
  browser.innerHTML = "<div style='padding:12px;color:gray;'>Loading...</div>";
  selectedFileId = null;
  document.getElementById("selected-file").style.display = "none";
  document.getElementById("preview-btn").style.display = "none";
  document.getElementById("preview-section").style.display = "none";

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
    const data = await response.json();
    const items = data.value;
    if (items.length === 0) {
      browser.innerHTML = "<div style='padding:12px;color:gray;'>This folder is empty.</div>";
      return;
    }
    browser.innerHTML = "";
    items.forEach((item) => {
      const isFolder = !!item.folder;
      const isExcel = item.name?.endsWith(".xlsx") || item.name?.endsWith(".xls");
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
          breadcrumbStack.push({ id: item.id, name: item.name });
          updateBreadcrumb();
          browseFolder(item.id);
        } else {
          document.querySelectorAll("#file-browser .browser-item").forEach(el => el.classList.remove("selected"));
          div.classList.add("selected");
          selectedFileId = item.id;
          document.getElementById("selected-file-name").textContent = item.name;
          document.getElementById("selected-file").style.display = "block";
          document.getElementById("preview-btn").style.display = "inline-block";
          setStatus("");
        }
      };
      browser.appendChild(div);
    });
    updateBreadcrumb();
  } catch (error) {
    browser.innerHTML = `<div style='padding:12px;color:red;'>Error: ${error.message}</div>`;
  }
}

async function browsePaymentsFolder(folderId) {
  const browser = document.getElementById("payments-file-browser");
  browser.innerHTML = "<div style='padding:12px;color:gray;'>Loading...</div>";
  selectedPaymentsFileId = null;
  document.getElementById("selected-payments-file").style.display = "none";
  document.getElementById("preview-payments-btn").style.display = "none";
  document.getElementById("preview-payments-section").style.display = "none";

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
    const data = await response.json();
    const items = data.value;
    if (items.length === 0) {
      browser.innerHTML = "<div style='padding:12px;color:gray;'>This folder is empty.</div>";
      return;
    }
    browser.innerHTML = "";
    items.forEach((item) => {
      const isFolder = !!item.folder;
      const isExcel = item.name?.endsWith(".xlsx") || item.name?.endsWith(".xls");
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
          breadcrumbStackPayments.push({ id: item.id, name: item.name });
          updateBreadcrumbPayments();
          browsePaymentsFolder(item.id);
        } else {
          document.querySelectorAll("#payments-file-browser .browser-item").forEach(el => el.classList.remove("selected"));
          div.classList.add("selected");
          selectedPaymentsFileId = item.id;
          document.getElementById("selected-payments-file-name").textContent = item.name;
          document.getElementById("selected-payments-file").style.display = "block";
          document.getElementById("preview-payments-btn").style.display = "inline-block";
          setPaymentsStatus("");
        }
      };
      browser.appendChild(div);
    });
    updateBreadcrumbPayments();
  } catch (error) {
    browser.innerHTML = `<div style='padding:12px;color:red;'>Error: ${error.message}</div>`;
  }
}

// ── BREADCRUMBS ───────────────────────────────────────────────
function updateBreadcrumb() {
  const breadcrumb = document.getElementById("breadcrumb");
  breadcrumb.innerHTML = breadcrumbStack.map((crumb, index) => {
    if (index === breadcrumbStack.length - 1) return `📁 ${crumb.name}`;
    return `<span data-index="${index}">${crumb.name}</span> › `;
  }).join("");
  breadcrumb.querySelectorAll("span").forEach((span) => {
    span.onclick = () => {
      const index = parseInt(span.getAttribute("data-index"));
      breadcrumbStack = breadcrumbStack.slice(0, index + 1);
      browseFolder(breadcrumbStack[breadcrumbStack.length - 1].id);
    };
  });
}

function updateBreadcrumbPayments() {
  const breadcrumb = document.getElementById("payments-breadcrumb");
  breadcrumb.innerHTML = breadcrumbStackPayments.map((crumb, index) => {
    if (index === breadcrumbStackPayments.length - 1) return `📁 ${crumb.name}`;
    return `<span data-index="${index}">${crumb.name}</span> › `;
  }).join("");
  breadcrumb.querySelectorAll("span").forEach((span) => {
    span.onclick = () => {
      const index = parseInt(span.getAttribute("data-index"));
      breadcrumbStackPayments = breadcrumbStackPayments.slice(0, index + 1);
      browsePaymentsFolder(breadcrumbStackPayments[breadcrumbStackPayments.length - 1].id);
    };
  });
}

// ── GET EMPLOYEE INITIALS FROM OPEN WORKBOOK ──────────────────
async function getEmployeeInitials() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("PAYMENTS");
    // B2:D2 is merged — reading B2 gives the full merged cell value
    const nameCell = sheet.getRange("B2");
    nameCell.load("values");
    await context.sync();

    const fullName = nameCell.values[0][0]; // "LASTNAME, FIRSTNAME"
    if (!fullName) throw new Error("Could not read employee name from cell B2 on PAYMENTS sheet.");

    // Convert "LASTNAME, FIRSTNAME" → initials "LF"
    const parts = String(fullName).split(",").map(p => p.trim());
    const lastInitial  = parts[0]?.[0]?.toUpperCase() || "";
    const firstInitial = parts[1]?.[0]?.toUpperCase() || "";

    return {
      initials: lastInitial + firstInitial,
      fullName: String(fullName).trim(),
    };
  });
}

// ── READ SECRETARY PAYMENT FILE ───────────────────────────────
async function readSecretaryPayments(fileId, employeeInitials) {
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets/Sheet1/usedRange`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );

  if (!response.ok) {
    const err = await response.json();
    throw new Error(err.error?.message || "Failed to read secretary payment file");
  }

  const data = await response.json();
  const rows = data.values;

  // Row 1 = date, Row 2 = headers, Row 3+ = data
  // Headers: CLIENT | AMOUNT | FORM | RCPT# | CASHIER | BONDSMAN | BALANCE | ...
  const headerRow = rows[1] || [];

  // Find column indices dynamically from the header row
  const colIndex = (name) => headerRow.findIndex(h =>
    String(h || "").toUpperCase().trim() === name.toUpperCase()
  );

  const clientCol   = colIndex("CLIENT");
  const amountCol   = colIndex("AMOUNT");
  const bondsmanCol = colIndex("BONDSMAN");

  if (clientCol === -1 || amountCol === -1 || bondsmanCol === -1) {
    throw new Error(
      `Could not find required columns. Found headers: ${headerRow.filter(Boolean).join(", ")}`
    );
  }

  const dataRows = rows.slice(2); // skip date row and header row
  const matched = [];

  for (const row of dataRows) {
    const client   = String(row[clientCol]   || "").trim();
    const amount   = row[amountCol];
    const bondsman = String(row[bondsmanCol] || "").trim().toUpperCase();

    // Skip empty rows
    if (!client || amount === null || amount === "" || amount === undefined) continue;

    // Only include rows that match this employee's initials
    if (bondsman !== employeeInitials.toUpperCase()) continue;

    matched.push({
      client:   client.toUpperCase(),
      amount:   Number(amount) || 0,
      bondsman,
    });
  }

  return matched;
}

// ── READ SUBMISSION FILE ──────────────────────────────────────
async function readSubmissionData(fileId) {
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets/Sheet1/usedRange`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  if (!response.ok) {
    const err = await response.json();
    throw new Error(err.error?.message || "Failed to read submission file");
  }
  const data = await response.json();
  const rows = data.values;
  const dataRows = rows.slice(3);
  const bonds = [];
  for (const row of dataRows) {
    const lastName  = row[2];
    const firstName = row[3];
    const ttlBond   = row[4];
    const amtChgd   = row[6];
    const amtCol    = row[7];
    const expense   = row[8];
    if (!lastName || !ttlBond) continue;
    if (String(lastName).toUpperCase().includes("TOTAL")) continue;
    const client = `${String(lastName).toUpperCase()}, ${String(firstName).toUpperCase()}`;
    const startBalance = (Number(amtChgd) || 0) - (Number(amtCol) || 0);
    bonds.push({
      client,
      ttlBond:      Number(ttlBond) || 0,
      amtCharged:   Number(amtChgd) || 0,
      amtCollected: Number(amtCol)  || 0,
      expense:      Number(expense) || 0,
      startBalance,
    });
  }
  return bonds;
}

// ── PREVIEW PAYMENTS IMPORT ───────────────────────────────────
async function previewPaymentsImport() {
  if (!selectedPaymentsFileId) {
    setPaymentsStatus("Please select a daily ledger file first.");
    return;
  }

  setPaymentsStatus("Diagnosing...");

  try {
    // Check what we read from B2
    const { initials, fullName } = await getEmployeeInitials();
    
    // Check what's in the ledger BONDSMAN column
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/items/${selectedPaymentsFileId}/workbook/worksheets/Sheet1/usedRange`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const data = await response.json();
    const rows = data.values;
    const dataRows = rows.slice(2);
    
    const bondsmanValues = dataRows
      .map(r => r[5])
      .filter(v => v !== null && v !== "" && v !== undefined);

    setPaymentsStatus(
      `B2 value: "${fullName}" → initials: "${initials}" | ` +
      `Bondsman values in file: ${JSON.stringify(bondsmanValues)}`
    );

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
      const table = context.workbook.tables.getItem("PaymentsTable");
      const bodyRange = table.getDataBodyRange();
      bodyRange.load("values, rowCount, columnCount");
      await context.sync();

      for (const payment of pendingPaymentRows) {
        const rowIndex = payment.rowIndex;

        // ── Write AMOUNT into AMT OF PAYMENT (col H, index 7) ──
        const amtOfPaymentCell = bodyRange.getCell(rowIndex, 7);
        amtOfPaymentCell.values = [[payment.amount]];
        await context.sync();

        // ── Read the recalculated END BALANCE (col I, index 8) ─
        const endBalanceCell = bodyRange.getCell(rowIndex, 8);
        endBalanceCell.load("values");
        await context.sync();

        const newEndBalance = endBalanceCell.values[0][0];

        // ── Write END BALANCE value → START BALANCE (col G, index 6) ─
        const startBalanceCell = bodyRange.getCell(rowIndex, 6);
        startBalanceCell.values = [[newEndBalance]];

        // ── Clear AMT OF PAYMENT (col H, index 7) ──────────────
        const clearAmtCell = bodyRange.getCell(rowIndex, 7);
        clearAmtCell.values = [[""]];

        await context.sync();
      }
    });

    const summary = document.getElementById("import-payments-summary");
    summary.className = "import-summary";
    summary.textContent = `✔ Successfully applied ${pendingPaymentRows.length} payment${pendingPaymentRows.length !== 1 ? "s" : ""} and updated START BALANCE.`;
    summary.style.display = "block";
    pendingPaymentRows = [];
    document.getElementById("confirm-payments-btn").disabled = false;
    setPaymentsStatus("");

  } catch (error) {
    const summary = document.getElementById("import-payments-summary");
    summary.className = "import-summary error";
    summary.textContent = `✖ Payment import failed: ${error.message}`;
    summary.style.display = "block";
    document.getElementById("confirm-payments-btn").disabled = false;
    console.error(error);
    setPaymentsStatus("");
  }
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
    if (bonds.length === 0) {
      setStatus("No bond entries found in the submission file.");
      return;
    }
    pendingImportRows = bonds;
    const tbody = document.getElementById("preview-tbody");
    tbody.innerHTML = "";
    bonds.forEach((bond) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${bond.client}</td>
        <td>$${bond.ttlBond.toLocaleString()}</td>
        <td>$${bond.amtCharged.toLocaleString()}</td>
        <td>$${bond.amtCollected.toLocaleString()}</td>
        <td>$${bond.expense.toLocaleString()}</td>
        <td>$${bond.startBalance.toLocaleString()}</td>
      `;
      tbody.appendChild(tr);
    });
    document.getElementById("import-summary").style.display = "none";
    document.getElementById("preview-section").style.display = "block";
    document.getElementById("confirm-import-btn").disabled = false;
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
      const table = context.workbook.tables.getItem("PaymentsTable");
      const bodyRange = table.getDataBodyRange();
      bodyRange.load("values, rowCount, columnCount");
      await context.sync();

      const colCount = bodyRange.columnCount;
      const currentRowCount = bodyRange.rowCount;

      const compacted = bodyRange.values.filter(row => {
        const clientVal = row[0];
        return clientVal !== null && clientVal !== "" && clientVal !== undefined;
      });

      const newRows = pendingImportRows.map(bond => {
        const row = new Array(colCount).fill("");
        row[0]  = bond.client;
        row[1]  = bond.ttlBond;
        row[2]  = bond.amtCharged;
        row[3]  = bond.amtCollected;
        row[4]  = bond.expense;
        row[6]  = bond.startBalance;
        return row;
      });

      const allRows = [...compacted, ...newRows];
      const neededRowCount = allRows.length;

      if (neededRowCount > currentRowCount) {
        const rowsToAdd = neededRowCount - currentRowCount;
        for (let i = 0; i < rowsToAdd; i++) {
          table.rows.add(null, [new Array(colCount).fill("")]);
        }
        await context.sync();
      }

      const freshBody = table.getDataBodyRange();
      freshBody.load("rowCount");
      await context.sync();

      const writeRange = freshBody.getCell(0, 0)
        .getResizedRange(allRows.length - 1, colCount - 1);
      writeRange.values = allRows;

      const newRowStartIndex = compacted.length;
      for (let i = 0; i < pendingImportRows.length; i++) {
        const rowIndex = newRowStartIndex + i;
        freshBody.getCell(rowIndex, 8).formulas  = [["=[@[START BALANCE]]-[@[AMT OF PAYMENT]]"]];
        freshBody.getCell(rowIndex, 10).formulas = [["=[@[BALANCE OWED]]-[@[PAYMENT]]"]];
        freshBody.getCell(rowIndex, 11).formulas = [["=[@[AMT OF PAYMENT]]*[@[% PAID ON BOND]]"]];
      }

      if (freshBody.rowCount > neededRowCount) {
        const leftoverCount = freshBody.rowCount - neededRowCount;
        const clearRange = freshBody.getCell(neededRowCount, 0)
          .getResizedRange(leftoverCount - 1, colCount - 1);
        clearRange.clear("Contents");
      }

      await context.sync();
    });

    const summary = document.getElementById("import-summary");
    summary.className = "import-summary";
    summary.textContent = `✔ Successfully imported ${pendingImportRows.length} bond${pendingImportRows.length > 1 ? "s" : ""} into PaymentsTable.`;
    summary.style.display = "block";
    pendingImportRows = [];
    document.getElementById("confirm-import-btn").disabled = false;
    setStatus("");
  } catch (error) {
    const summary = document.getElementById("import-summary");
    summary.className = "import-summary error";
    summary.textContent = `✖ Import failed: ${error.message}`;
    summary.style.display = "block";
    document.getElementById("confirm-import-btn").disabled = false;
    console.error(error);
    setStatus("");
  }
}

// ── UI HELPERS ────────────────────────────────────────────────
function showMainSection(userName) {
  document.getElementById("sign-in-section").style.display = "none";
  document.getElementById("main-section").style.display = "block";
  document.getElementById("user-name").textContent = userName;
}

function setStatus(message) {
  document.getElementById("status").textContent = message;
}

function setPaymentsStatus(message) {
  document.getElementById("payments-status").textContent = message;
}