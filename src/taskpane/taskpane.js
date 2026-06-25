/* global console, document, Office, Excel */

// ── CONFIG ────────────────────────────────────────────────────
const START_FOLDER_PATH          = "BailBonds/EmployeeSubmissions";
const START_FOLDER_PATH_PAYMENTS = "BailBonds/Payments";
const SUBMISSION_SHEET_NAME      = "Sheet1";   // fallback to index 0 if not found
const ARCHIVE_DATE_FORMAT        = "YYYY-MM-DD"; // change to alter archive naming

// Row offsets in source files (0-based index into .values array)
const SUBMISSION_HEADER_ROW  = 2; // row 3 in Excel (0-indexed)
const SUBMISSION_DATA_START  = 3; // row 4 in Excel (0-indexed)
const LEDGER_HEADER_ROW      = 1; // row 2 in Excel (0-indexed)
const LEDGER_DATA_START      = 2; // row 3 in Excel (0-indexed)

// Secretary daily ledger column indices
const LEDGER_CLIENT_COL   = 0;
const LEDGER_AMOUNT_COL   = 1;
const LEDGER_BONDSMAN_COL = 5;

// PaymentsTable column indices
const PT_CLIENT      = 0;
const PT_TTL_BOND    = 1;
const PT_AMT_CHARGED = 2;
const PT_AMT_COLLECT = 3;
const PT_EXPENSE     = 4;
const PT_PCT_PAID    = 5;
const PT_START_BAL   = 6;
const PT_AMT_PAYMENT = 7;
const PT_END_BAL     = 8;
const PT_BAL_OWED    = 9;
const PT_ENDING_BAL  = 10;
const PT_PAYMENT     = 11;
const PT_REACH       = 12;

// Highlight color for processed rows in source files
const PROCESSED_ROW_COLOR = "#90EE90"; // light green

// ── STATE ─────────────────────────────────────────────────────
let accessToken            = null;
let msalInstance           = null;
let selectedFileId         = null;
let selectedFileParentId   = null; // parent folder ID for archiving
let selectedFileName       = null;
let selectedPaymentsFileId = null;
let selectedPaymentsFileParentId = null;
let selectedPaymentsFileName     = null;
let pendingImportRows      = [];
let pendingPaymentRows     = [];
let submissionImportDone   = false; // tracks if archive button should be enabled
let paymentsImportDone     = false;

// Source row indices for marking after import
let submissionSourceRows   = []; // [{rowIndex, excelRow}]
let ledgerSourceRows       = []; // [{rowIndex, excelRow}]

// Browser config
const browsers = {
  submission: {
    stackId:   "breadcrumb",
    browserId: "file-browser",
    stack:     [{ id: "root", name: "My Files" }],
    onSelect:  onSubmissionFileSelected,
  },
  payments: {
    stackId:   "payments-breadcrumb",
    browserId: "payments-file-browser",
    stack:     [{ id: "root", name: "My Files" }],
    onSelect:  onPaymentsFileSelected,
  },
};

// ── OFFICE INIT ───────────────────────────────────────────────
Office.onReady(() => {
  document.getElementById("sign-in-btn").onclick            = signIn;
  document.getElementById("sign-out-btn").onclick           = signOut;
  document.getElementById("refresh-token-btn").onclick      = refreshTokenViaDialog;
  document.getElementById("preview-btn").onclick            = previewImport;
  document.getElementById("confirm-import-btn").onclick     = confirmImport;
  document.getElementById("cancel-import-btn").onclick      = cancelImport;
  document.getElementById("archive-submission-btn").onclick = archiveSubmissionFile;
  document.getElementById("preview-payments-btn").onclick   = previewPaymentsImport;
  document.getElementById("confirm-payments-btn").onclick   = confirmPaymentsImport;
  document.getElementById("cancel-payments-btn").onclick    = cancelPaymentsImport;
  document.getElementById("archive-payments-btn").onclick   = archivePaymentsFile;
});

// ── AUTH ─────────────────────────────────────────────────────
function signIn() {
  setStatus("Opening sign in window...");
  openAuthDialog(async (token, userName) => {
    accessToken = token;
    showMainSection(userName);
    setStatus("");
    await navigateBrowserToFolder("submission", START_FOLDER_PATH);
    await navigateBrowserToFolder("payments",   START_FOLDER_PATH_PAYMENTS);
  });
}

function openAuthDialog(onSuccess) {
  Office.context.ui.displayDialogAsync(
    "https://cordusbailey02.github.io/AddinTest/dialog.html",
    { height: 60, width: 30, promptBeforeOpen: false },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        setStatus("Failed to open sign in dialog: " + result.error.message);
        return;
      }
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (msg) => {
        dialog.close();
        try {
          const data = JSON.parse(msg.message);
          if (data.status === "success") {
            await onSuccess(data.accessToken, data.userName);
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

async function refreshTokenViaDialog() {
  showLoading("Refreshing session...");
  openAuthDialog(async (token) => {
    accessToken = token;
    hideLoading();
    showNotification("Session refreshed successfully.", "success");
    document.getElementById("session-expired-banner").style.display = "none";
  });
}

function signOut() {
  accessToken              = null;
  selectedFileId           = null;
  selectedFileParentId     = null;
  selectedFileName         = null;
  selectedPaymentsFileId   = null;
  selectedPaymentsFileParentId = null;
  selectedPaymentsFileName = null;
  pendingImportRows        = [];
  pendingPaymentRows       = [];
  submissionImportDone     = false;
  paymentsImportDone       = false;
  submissionSourceRows     = [];
  ledgerSourceRows         = [];

  browsers.submission.stack = [{ id: "root", name: "My Files" }];
  browsers.payments.stack   = [{ id: "root", name: "My Files" }];

  document.getElementById("sign-in-section").style.display            = "block";
  document.getElementById("main-section").style.display               = "none";
  document.getElementById("preview-section").style.display            = "none";
  document.getElementById("preview-payments-section").style.display   = "none";
  document.getElementById("session-expired-banner").style.display     = "none";
  setStatus("");
  setPaymentsStatus("");
}

// ── TOKEN MANAGEMENT ──────────────────────────────────────────
async function graphFetch(path, options = {}) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/${path}`;
  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });

  if (response.status === 401) {
    // Token expired — show banner and throw
    document.getElementById("session-expired-banner").style.display = "block";
    throw new Error("Session expired. Please refresh your session using the button above.");
  }

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err.error?.message || `Graph request failed: ${path}`);
  }

  // Some operations return no body (204 No Content)
  const text = await response.text();
  return text ? JSON.parse(text) : null;
}

// ── GENERIC FILE BROWSER ──────────────────────────────────────
async function navigateBrowserToFolder(browserKey, folderPath) {
  const b = browsers[browserKey];
  try {
    const folder = await graphFetch(`root:/${folderPath}`);
    b.stack = [
      { id: "root",    name: "My Files"  },
      { id: folder.id, name: folder.name },
    ];
    await browseFolder(browserKey, folder.id);
  } catch {
    b.stack = [{ id: "root", name: "My Files" }];
    await browseFolder(browserKey, "root");
  }
}

async function browseFolder(browserKey, folderId) {
  const b       = browsers[browserKey];
  const browser = document.getElementById(b.browserId);
  browser.innerHTML = "<div style='padding:12px;color:gray;'>Loading...</div>";

  if (browserKey === "submission") {
    selectedFileId = null;
    document.getElementById("selected-file").style.display           = "none";
    document.getElementById("preview-btn").style.display             = "none";
    document.getElementById("preview-section").style.display         = "none";
    document.getElementById("archive-submission-btn").style.display  = "none";
    submissionImportDone = false;
  } else {
    selectedPaymentsFileId = null;
    document.getElementById("selected-payments-file").style.display        = "none";
    document.getElementById("preview-payments-btn").style.display          = "none";
    document.getElementById("preview-payments-section").style.display      = "none";
    document.getElementById("archive-payments-btn").style.display          = "none";
    paymentsImportDone = false;
  }

  try {
    const url = folderId === "root"
      ? "root/children"
      : `items/${folderId}/children`;

    const data = await graphFetch(
      url + "?$orderby=name&$select=id,name,folder,file,size,parentReference"
    );

    const items = data.value;
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

function renderBreadcrumb(browserKey) {
  const b  = browsers[browserKey];
  const el = document.getElementById(b.stackId);
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

function onSubmissionFileSelected(item) {
  selectedFileId       = item.id;
  selectedFileName     = item.name;
  selectedFileParentId = item.parentReference?.id || null;
  document.getElementById("selected-file-name").textContent = item.name;
  document.getElementById("selected-file").style.display    = "block";
  document.getElementById("preview-btn").style.display      = "inline-block";
  document.getElementById("archive-submission-btn").style.display = "none";
  submissionImportDone = false;
  setStatus("");
}

function onPaymentsFileSelected(item) {
  selectedPaymentsFileId       = item.id;
  selectedPaymentsFileName     = item.name;
  selectedPaymentsFileParentId = item.parentReference?.id || null;
  document.getElementById("selected-payments-file-name").textContent = item.name;
  document.getElementById("selected-payments-file").style.display    = "block";
  document.getElementById("preview-payments-btn").style.display      = "inline-block";
  document.getElementById("archive-payments-btn").style.display      = "none";
  paymentsImportDone = false;
  setPaymentsStatus("");
}

// ── SHEET NAME RESOLUTION ─────────────────────────────────────
async function resolveSheetName(fileId, preferredName) {
  // Try preferred name first, fall back to first sheet
  try {
    const data = await graphFetch(
      `items/${fileId}/workbook/worksheets`
    );
    const sheets = data.value;
    const match  = sheets.find(
      s => s.name.toLowerCase() === preferredName.toLowerCase()
    );
    return match ? match.name : sheets[0]?.name;
  } catch {
    return preferredName; // best effort
  }
}

// ── EMPLOYEE INITIALS ─────────────────────────────────────────
async function getEmployeeInitials() {
  return await Excel.run(async (context) => {
    const sheet    = context.workbook.worksheets.getItem("PAYMENTS");
    const nameCell = sheet.getRange("B2");
    nameCell.load("values");
    await context.sync();

    const fullName = nameCell.values[0][0];
    if (!fullName) throw new Error("Could not read employee name from B2 on PAYMENTS sheet.");

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
  const sheetName = await resolveSheetName(fileId, SUBMISSION_SHEET_NAME);
  const data      = await graphFetch(
    `items/${fileId}/workbook/worksheets('${encodeURIComponent(sheetName)}')/usedRange`
  );
  const rows      = data.values;
  const headerRow = rows[SUBMISSION_HEADER_ROW] || [];
  const dataRows  = rows.slice(SUBMISSION_DATA_START);

  const bonds      = [];
  const sourceRows = [];

  dataRows.forEach((row, idx) => {
    // ── All raw submission columns ────────────────────────────
    const date           = row[0];   // Col A - DATE
    const jail           = row[1];   // Col B - JAIL
    const lastName       = row[2];   // Col C - LAST NAME
    const firstName      = row[3];   // Col D - FIRST NAME
    const ttlBond        = row[4];   // Col E - TTL BOND AMOUNT
    const bondNum        = row[5];   // Col F - #
    const amtChgd        = row[6];   // Col G - TOTAL AMT CHGD
    const amtCol         = row[7];   // Col H - TOTAL AMT COL
    const expense        = row[8];   // Col I - TTL EXPENSE
    const travel         = row[9];   // Col J - TRAVEL
    const amtInEnv       = row[10];  // Col K - AMT IN ENV
    const pymtType       = row[11];  // Col L - PYMT TYPE
    const balanceOwedRaw = row[12];  // Col M - BALANCE OWED
    const administration = row[13];  // Col N - ADMINISTRATION

    // Skip empty rows and totals row
    if (!lastName || !ttlBond) return;
    if (String(lastName).toUpperCase().includes("TOTAL")) return;

    const excelRow = SUBMISSION_DATA_START + idx + 1;

    // ── Numeric conversions ───────────────────────────────────
    const ttlBondNum   = Number(ttlBond) || 0;
    const amtChgdNum   = Number(amtChgd) || 0;
    const amtColNum    = Number(amtCol)  || 0;
    const expenseNum   = Number(expense) || 0;
    const travelNum    = Number(travel)  || 0;

    // ── COMPANY COST: Total cost company has incurred ─────────
    // COMPANY_COST = (TTL BOND × 2.7%) + TTL EXPENSE + TRAVEL
    const companyCost = (ttlBondNum * 0.027) + expenseNum + travelNum;

    // ── NET: Amount collected minus company cost ──────────────
    const net = amtColNum - companyCost;

    // ── START BALANCE: What client still owes company ─────────
    const startBalance = amtChgdNum - amtColNum;

    // ── BALANCE OWED & % PAID ON BOND ─────────────────────────
    let balanceOwed    = 0;
    let pctPaidOnBond  = 0;
    let initialPayment = 0;
    let netIsNegative  = false;

    if (net < 0) {
      netIsNegative  = true;
      balanceOwed    = 0;
      pctPaidOnBond  = 0;
      initialPayment = 0;
    } else {
      const rawDownPayment = net * 0.25;
      initialPayment       = rawDownPayment < 25 ? 25 : rawDownPayment;
      balanceOwed          = Math.max(0, net - initialPayment);
      pctPaidOnBond        = startBalance > 0
        ? parseFloat((balanceOwed / startBalance).toFixed(4))
        : 0;
    }

    // Format date for display
    let displayDate = "";
    if (date instanceof Date || (typeof date === "string" && date)) {
      displayDate = date instanceof Date
        ? date.toLocaleDateString()
        : String(date);
    } else if (typeof date === "number") {
      // Excel serial date number
      const jsDate = new Date((date - 25569) * 86400 * 1000);
      displayDate  = jsDate.toLocaleDateString();
    }

    bonds.push({
      // ── Every raw column from submission ──────────────────
      rawDate:        displayDate,
      rawJail:        jail         ?? "",
      lastName:       String(lastName).toUpperCase(),
      firstName:      String(firstName).toUpperCase(),
      rawTtlBond:     ttlBondNum,
      rawBondNum:     bondNum      ?? "",
      rawAmtChgd:     amtChgdNum,
      rawAmtCol:      amtColNum,
      rawExpense:     expenseNum,
      rawTravel:      travelNum,
      rawAmtInEnv:    amtInEnv     ?? "",
      rawPymtType:    pymtType     ?? "",
      rawBalanceOwed: balanceOwedRaw ?? "",
      rawAdmin:       administration ?? "",

      // ── Derived/combined fields ───────────────────────────
      client:         `${String(lastName).toUpperCase()}, ${String(firstName).toUpperCase()}`,
      ttlBond:        ttlBondNum,
      amtCharged:     amtChgdNum,
      amtCollected:   amtColNum,
      expense:        expenseNum,
      travel:         travelNum,

      // ── Calculated PAYMENTS fields ────────────────────────
      companyCost,
      net,
      startBalance,
      balanceOwed,
      pctPaidOnBond,
      initialPayment,
      netIsNegative,

      // ── Source tracking ───────────────────────────────────
      sourceRowIndex: idx,
      excelRow,
    });

    sourceRows.push({ rowIndex: idx, excelRow });
  });

  return { bonds, headerRow, sourceRows, sheetName };
}

// ── READ SECRETARY LEDGER ─────────────────────────────────────
async function readSecretaryPayments(fileId, employeeInitials) {
  const sheetName = await resolveSheetName(fileId, SUBMISSION_SHEET_NAME);
  const data      = await graphFetch(
    `items/${fileId}/workbook/worksheets('${encodeURIComponent(sheetName)}')/usedRange`
  );
  const rows      = data.values;
  const headerRow = rows[LEDGER_HEADER_ROW] || [];
  const dataRows  = rows.slice(LEDGER_DATA_START);

  const matched    = [];
  const sourceRows = [];

  dataRows.forEach((row, idx) => {
    const client   = String(row[LEDGER_CLIENT_COL]   ?? "").trim();
    const amount   = row[LEDGER_AMOUNT_COL];
    const bondsman = String(row[LEDGER_BONDSMAN_COL] ?? "").trim().toUpperCase();

    if (!client) return;
    if (amount === null || amount === "" || amount === undefined) return;
    if (bondsman !== employeeInitials.toUpperCase()) return;

    const excelRow = LEDGER_DATA_START + idx + 1; // 1-based

    matched.push({
      client:   client.toUpperCase(),
      amount:   Number(amount) || 0,
      bondsman,
      excelRow,
    });

    sourceRows.push({ rowIndex: idx, excelRow });
  });

  return { matched, headerRow, sourceRows, sheetName };
}

// ── MARK ROWS GREEN IN SOURCE FILE ───────────────────────────
async function markRowsProcessed(fileId, sheetName, excelRows, stampColumnIndex = null) {
  const encodedSheet = encodeURIComponent(sheetName);
  const timestamp    = new Date().toLocaleString();

  for (const excelRow of excelRows) {
    // Get the used range to know how many columns to highlight
    const rangeData = await graphFetch(
      `items/${fileId}/workbook/worksheets('${encodedSheet}')/usedRange?$select=columnCount`
    );
    const colCount = rangeData.columnCount || 14;

    // Convert column count to letter (e.g. 14 → N)
    const lastCol  = columnIndexToLetter(colCount - 1);
    const address  = `A${excelRow}:${lastCol}${excelRow}`;

    // Highlight the row green
    await graphFetch(
      `items/${fileId}/workbook/worksheets('${encodedSheet}')/range(address='${address}')/format/fill`,
      {
        method: "PATCH",
        body: JSON.stringify({ color: PROCESSED_ROW_COLOR }),
      }
    );

    // Add timestamp in the specified column if provided
    if (stampColumnIndex !== null) {
      const stampCol     = columnIndexToLetter(stampColumnIndex);
      const stampAddress = `${stampCol}${excelRow}`;
      await graphFetch(
        `items/${fileId}/workbook/worksheets('${encodedSheet}')/range(address='${stampAddress}')`,
        {
          method:  "PATCH",
          body: JSON.stringify({ values: [[timestamp]] }),
        }
      );
    }
  }
}

function columnIndexToLetter(index) {
  let letter = "";
  let n      = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter    = String.fromCharCode(65 + rem) + letter;
    n         = Math.floor((n - 1) / 26);
  }
  return letter;
}

// ── ARCHIVE FILE ──────────────────────────────────────────────
async function archiveFile(fileId, fileName, parentFolderId) {
  showLoading("Archiving file...");
  try {
    // Step 1: Get or create Archive subfolder
    const archiveFolderId = await getOrCreateArchiveFolder(parentFolderId);

    // Step 2: Build archive filename with today's date
    const dateStr      = formatDate(new Date(), ARCHIVE_DATE_FORMAT);
    const nameParts    = fileName.replace(/\.xlsx?$/i, "");
    const archiveName  = `${nameParts}_${dateStr}.xlsx`;

    // Step 3: Copy file into Archive folder with new name
    await graphFetch(`items/${fileId}/copy`, {
      method: "POST",
      body: JSON.stringify({
        parentReference: { id: archiveFolderId },
        name:            archiveName,
      }),
    });

    // Step 4: Delete the original
    await graphFetch(`items/${fileId}`, { method: "DELETE" });

    hideLoading();
    return archiveName;
  } catch (error) {
    hideLoading();
    throw error;
  }
}

async function getOrCreateArchiveFolder(parentFolderId) {
  // Look for folder named 'WeeklyBondsArchive' — change this constant to rename
  const ARCHIVE_FOLDER_NAME = "WeeklyBondsArchive";

  try {
    const data  = await graphFetch(`items/${parentFolderId}/children?$select=id,name,folder`);
    const found = data.value.find(
      item => item.folder && item.name.toLowerCase() === ARCHIVE_FOLDER_NAME.toLowerCase()
    );
    if (found) return found.id;
  } catch { /* not found, will create */ }

  // Create the archive folder
  const created = await graphFetch(`items/${parentFolderId}/children`, {
    method: "POST",
    body: JSON.stringify({
      name:   ARCHIVE_FOLDER_NAME,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename",
    }),
  });
  return created.id;
}

function formatDate(date, format) {
  const y  = date.getFullYear();
  const m  = String(date.getMonth() + 1).padStart(2, "0");
  const d  = String(date.getDate()).padStart(2, "0");
  return format
    .replace("YYYY", y)
    .replace("MM",   m)
    .replace("DD",   d);
}

async function archiveSubmissionFile() {
  if (!selectedFileId || !selectedFileParentId) {
    showNotification("No submission file selected to archive.", "error");
    return;
  }
  try {
    const archiveName = await archiveFile(
      selectedFileId, selectedFileName, selectedFileParentId
    );
    showNotification(`✔ Archived as "${archiveName}"`, "success");

    // Reset submission section state
    selectedFileId       = null;
    selectedFileName     = null;
    selectedFileParentId = null;
    submissionImportDone = false;
    submissionSourceRows = [];
    pendingImportRows    = [];

    // Hide everything and refresh the browser back to start folder
    document.getElementById("selected-file").style.display          = "none";
    document.getElementById("preview-btn").style.display            = "none";
    document.getElementById("preview-section").style.display        = "none";
    document.getElementById("archive-submission-btn").style.display = "none";
    document.getElementById("import-summary").style.display         = "none";
    setStatus("");

    // Refresh browser to show updated folder contents
    await navigateBrowserToFolder("submission", START_FOLDER_PATH);

  } catch (error) {
    showNotification(`✖ Archive failed: ${error.message}`, "error");
  }
}

async function archivePaymentsFile() {
  if (!selectedPaymentsFileId || !selectedPaymentsFileParentId) {
    showNotification("No payments file selected to archive.", "error");
    return;
  }
  try {
    const archiveName = await archiveFile(
      selectedPaymentsFileId, selectedPaymentsFileName, selectedPaymentsFileParentId
    );
    showNotification(`✔ Archived as "${archiveName}"`, "success");

    // Reset payments section state
    selectedPaymentsFileId           = null;
    selectedPaymentsFileName         = null;
    selectedPaymentsFileParentId     = null;
    paymentsImportDone               = false;
    ledgerSourceRows                 = [];
    pendingPaymentRows               = [];

    // Hide everything and refresh the browser back to start folder
    document.getElementById("selected-payments-file").style.display        = "none";
    document.getElementById("preview-payments-btn").style.display          = "none";
    document.getElementById("preview-payments-section").style.display      = "none";
    document.getElementById("archive-payments-btn").style.display          = "none";
    document.getElementById("import-payments-summary").style.display       = "none";
    document.getElementById("employee-badge").style.display                = "none";
    document.getElementById("payments-unmatched-warning").style.display    = "none";
    setPaymentsStatus("");

    // Refresh browser to show updated folder contents
    await navigateBrowserToFolder("payments", START_FOLDER_PATH_PAYMENTS);

  } catch (error) {
    showNotification(`✖ Archive failed: ${error.message}`, "error");
  }
}

// ── DYNAMIC PREVIEW TABLE ─────────────────────────────────────
function buildPreviewTable(containerId, headers, rows, extraColumns = []) {
  const container = document.getElementById(containerId);

  const allHeaders = [...headers, ...extraColumns.map(c => c.label)];

  const thead = allHeaders
    .map(h => `<th>${h}</th>`)
    .join("");

  const tbody = rows.map(row => {
    const cells = headers.map((_, i) => {
      const val = row[i];
      const num = Number(val);
      if (!isNaN(num) && val !== "" && val !== null) {
        return `<td>$${num.toLocaleString()}</td>`;
      }
      return `<td>${val ?? ""}</td>`;
    });
    const extraCells = extraColumns.map(col => {
      const val = col.getValue(row);
      return `<td>${val}</td>`;
    });
    return `<tr>${[...cells, ...extraCells].join("")}</tr>`;
  }).join("");

  container.innerHTML = `
    <div style="overflow-x:auto;">
      <table class="preview-table">
        <thead><tr>${thead}</tr></thead>
        <tbody>${tbody}</tbody>
      </table>
    </div>
  `;
}

// ── PRE-FLIGHT CHECKS ─────────────────────────────────────────
async function runPreflightChecks() {
  return await Excel.run(async (context) => {
    const errors = [];

    // Check PAYMENTS sheet exists
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    const tables = context.workbook.tables;
    tables.load("items/name");
    await context.sync();

    const sheetNames = sheets.items.map(s => s.name);
    const tableNames = tables.items.map(t => t.name);

    if (!sheetNames.includes("PAYMENTS")) {
      errors.push("PAYMENTS sheet not found. Please rename the sheet to 'PAYMENTS'.");
    }
    if (!tableNames.includes("PaymentsTable")) {
      errors.push("PaymentsTable not found. Please convert the PAYMENTS data range to a table named 'PaymentsTable'.");
    }

    // Check B2 has employee name
    if (sheetNames.includes("PAYMENTS")) {
      const sheet    = context.workbook.worksheets.getItem("PAYMENTS");
      const nameCell = sheet.getRange("B2");
      nameCell.load("values");
      await context.sync();
      if (!nameCell.values[0][0]) {
        errors.push("Cell B2 on PAYMENTS sheet is empty. It should contain the employee name (LASTNAME, FIRSTNAME).");
      }
    }

    return errors;
  });
}

// ── PREVIEW BOND IMPORT ───────────────────────────────────────
async function previewImport() {
  if (!selectedFileId) {
    setStatus("Please select a submission file first.");
    return;
  }

  showLoading("Reading submission file...");
  document.getElementById("preview-section").style.display = "none";

  try {
    const errors = await runPreflightChecks();
    if (errors.length) {
      hideLoading();
      showNotification("⚠ Setup issues found:\n• " + errors.join("\n• "), "error");
      return;
    }

    const { bonds, sourceRows } = await readSubmissionData(selectedFileId);

    if (!bonds.length) {
      hideLoading();
      setStatus("No bond entries found in the submission file.");
      return;
    }

    pendingImportRows    = bonds;
    submissionSourceRows = sourceRows;

    // ── Section 1: ALL raw submission columns ─────────────────
    // Show every column exactly as submitted — nothing omitted
    const submissionHeaders = [
      "DATE", "JAIL", "LAST NAME", "FIRST NAME",
      "TTL BOND", "#", "AMT CHGD", "AMT COL",
      "EXPENSE", "TRAVEL", "AMT IN ENV",
      "PYMT TYPE", "BALANCE OWED", "ADMINISTRATION",
    ];
    const submissionRows = bonds.map(b => [
      b.rawDate,
      b.rawJail,
      b.lastName,
      b.firstName,
      `$${b.rawTtlBond.toLocaleString()}`,
      b.rawBondNum,
      `$${b.rawAmtChgd.toLocaleString()}`,
      `$${b.rawAmtCol.toLocaleString()}`,
      `$${b.rawExpense.toLocaleString()}`,
      `$${b.rawTravel.toLocaleString()}`,
      b.rawAmtInEnv !== "" ? `$${Number(b.rawAmtInEnv).toLocaleString()}` : "—",
      b.rawPymtType  || "—",
      b.rawBalanceOwed !== "" ? `$${Number(b.rawBalanceOwed).toLocaleString()}` : "—",
      b.rawAdmin     || "—",
    ]);
    buildPreviewTable(
      "preview-submission-table",
      submissionHeaders,
      submissionRows
    );

    // ── Section 2: Calculated values going into PAYMENTS ─────
    const paymentsHeaders = [
      "CLIENT", "COMPANY COST", "NET",
      "START BALANCE", "BALANCE OWED",
      "% PAID ON BOND", "INITIAL PAYMENT", "NOTE",
    ];
    const paymentsRows = bonds.map(b => [
      b.client,
      `$${b.companyCost.toFixed(2)}`,
      b.netIsNegative
        ? `($${Math.abs(b.net).toFixed(2)})`
        : `$${b.net.toFixed(2)}`,
      `$${b.startBalance.toLocaleString()}`,
      b.netIsNegative ? "TBD"                    : `$${b.balanceOwed.toFixed(2)}`,
      b.netIsNegative ? "TBD"                    : `${(b.pctPaidOnBond * 100).toFixed(2)}%`,
      b.netIsNegative ? "—"                      : `$${b.initialPayment.toFixed(2)}`,
      b.netIsNegative
        ? "⚠ Net negative — BALANCE OWED set to 0"
        : b.initialPayment === 25
          ? "Min. $25 applied"
          : "",
    ]);
    buildPreviewTable(
      "preview-payments-calc-table",
      paymentsHeaders,
      paymentsRows
    );

    document.getElementById("import-summary").style.display  = "none";
    document.getElementById("preview-section").style.display = "block";
    document.getElementById("confirm-import-btn").disabled   = false;
    hideLoading();
    setStatus(
      `${bonds.length} bond${bonds.length > 1 ? "s" : ""} ready to import.`
    );

  } catch (error) {
    hideLoading();
    setStatus("Error reading submission: " + error.message);
    console.error(error);
  }
}

function cancelImport() {
  pendingImportRows    = [];
  submissionSourceRows = [];
  document.getElementById("preview-section").style.display = "none";
  setStatus("Import cancelled.");
}

// ── CONFIRM BOND IMPORT ───────────────────────────────────────
async function confirmImport() {
  if (!pendingImportRows.length) return;
  document.getElementById("confirm-import-btn").disabled = true;
  showLoading("Importing bonds...");

  const succeeded = [];
  const failed    = [];

  try {
    await Excel.run(async (context) => {
      const table     = context.workbook.tables.getItem("PaymentsTable");
      const bodyRange = table.getDataBodyRange();
      bodyRange.load("values, formulas, rowCount, columnCount");
      await context.sync();

      const colCount        = bodyRange.columnCount;
      const currentRowCount = bodyRange.rowCount;

      // ── STEP 1: Compact — preserve values AND formulas ───────
      const compactedValues   = [];
      const compactedFormulas = [];

      for (let i = 0; i < currentRowCount; i++) {
        const clientVal = bodyRange.values[i][PT_CLIENT];
        if (clientVal !== null && clientVal !== "" && clientVal !== undefined) {
          compactedValues.push(bodyRange.values[i]);
          compactedFormulas.push(bodyRange.formulas[i]);
        }
      }

      // ── STEP 2: Build new rows from pending bonds ─────────────
      const newRows = [];
      for (const bond of pendingImportRows) {
        try {
          const row = new Array(colCount).fill("");
          row[PT_CLIENT]      = bond.client;
          row[PT_TTL_BOND]    = bond.ttlBond;
          row[PT_AMT_CHARGED] = bond.amtCharged;
          row[PT_AMT_COLLECT] = bond.amtCollected;
          row[PT_EXPENSE]     = bond.expense;
          row[PT_PCT_PAID]    = bond.pctPaidOnBond;   // auto-calculated
          row[PT_START_BAL]   = bond.startBalance;
          row[PT_BAL_OWED]    = bond.balanceOwed;     // auto-calculated
          // PT_AMT_PAYMENT (H), PT_END_BAL (I), PT_ENDING_BAL (K),
          // PT_PAYMENT (L), PT_REACH (M) — left blank, formulas handle them
          newRows.push({ values: row });
          succeeded.push(bond);
        } catch (err) {
          failed.push({ bond, reason: err.message });
        }
      }

      const neededRowCount = compactedValues.length + newRows.length;

      // ── STEP 3: Grow table if needed ──────────────────────────
      if (neededRowCount > currentRowCount) {
        for (let i = 0; i < neededRowCount - currentRowCount; i++) {
          table.rows.add(null, [new Array(colCount).fill("")]);
        }
        await context.sync();
      }

      const freshBody = table.getDataBodyRange();
      freshBody.load("rowCount");
      await context.sync();

      // ── STEP 4: Write compacted rows preserving formulas ─────
      for (let i = 0; i < compactedValues.length; i++) {
        const rowRange = freshBody.getCell(i, 0)
          .getResizedRange(0, colCount - 1);
        rowRange.values = [compactedValues[i]];

        // Restore cells that had formulas
        const formulaRow = compactedFormulas[i];
        for (let col = 0; col < colCount; col++) {
          const formula = formulaRow[col];
          if (typeof formula === "string" && formula.startsWith("=")) {
            freshBody.getCell(i, col).formulas = [[formula]];
          }
        }
      }

      // ── STEP 5: Write new rows with structured formulas ───────
      const newRowStart = compactedValues.length;
      for (let i = 0; i < newRows.length; i++) {
        const ri = newRowStart + i;
        freshBody.getCell(ri, 0)
          .getResizedRange(0, colCount - 1)
          .values = [newRows[i].values];

        freshBody.getCell(ri, PT_END_BAL).formulas    =
          [["=[@[START BALANCE]]-[@[AMT OF PAYMENT]]"]];
        freshBody.getCell(ri, PT_ENDING_BAL).formulas =
          [["=[@[BALANCE OWED]]-[@[PAYMENT]]"]];
        freshBody.getCell(ri, PT_PAYMENT).formulas    =
          [["=[@[AMT OF PAYMENT]]*[@[% PAID ON BOND]]"]];
      }

      // ── STEP 6: Clear leftover rows if table shrank ───────────
      if (freshBody.rowCount > neededRowCount) {
        freshBody.getCell(neededRowCount, 0)
          .getResizedRange(
            freshBody.rowCount - neededRowCount - 1,
            colCount - 1
          )
          .clear("Contents");
      }

      await context.sync();
    });

    // Mark processed rows green in source submission file
    if (succeeded.length && selectedFileId) {
      showLoading("Marking rows as processed...");
      const successExcelRows = submissionSourceRows
        .filter((_, i) => i < succeeded.length)
        .map(r => r.excelRow);
      const { sheetName } = await readSubmissionData(selectedFileId);
      await markRowsProcessed(
        selectedFileId, sheetName, successExcelRows, null
      );
    }

    let message = `✔ Successfully imported ${succeeded.length} bond${succeeded.length !== 1 ? "s" : ""}.`;
    if (failed.length) {
      message += `\n⚠ ${failed.length} row${failed.length !== 1 ? "s" : ""} skipped:\n`;
      message += failed.map(f => `• ${f.bond.client}: ${f.reason}`).join("\n");
    }

    showSummary("import-summary", message, failed.length > 0 && !succeeded.length);
    pendingImportRows    = [];
    document.getElementById("confirm-import-btn").disabled          = false;
    submissionImportDone = true;
    document.getElementById("archive-submission-btn").style.display = "inline-block";
    document.getElementById("archive-submission-btn").disabled      = false;
    hideLoading();
    setStatus("");

  } catch (error) {
    hideLoading();
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

  showLoading("Reading daily ledger...");
  document.getElementById("preview-payments-section").style.display = "none";

  try {
    const errors = await runPreflightChecks();
    if (errors.length) {
      hideLoading();
      showNotification("⚠ Setup issues:\n• " + errors.join("\n• "), "error");
      return;
    }

    const { initials, fullName } = await getEmployeeInitials();
    document.getElementById("detected-employee").textContent = `${fullName} (${initials})`;
    document.getElementById("employee-badge").style.display  = "block";

    const { matched, sourceRows, sheetName } = await readSecretaryPayments(
      selectedPaymentsFileId, initials
    );

    if (!matched.length) {
      hideLoading();
      setPaymentsStatus(`No payments found for bondsman "${initials}" in this file.`);
      return;
    }

    ledgerSourceRows = sourceRows;

    // Load PaymentsTable for cross-referencing
    const tableRows = await Excel.run(async (context) => {
      const table     = context.workbook.tables.getItem("PaymentsTable");
      const bodyRange = table.getDataBodyRange();
      bodyRange.load("values");
      await context.sync();
      return bodyRange.values;
    });

    const tableClients = tableRows.map((row, idx) => ({
      rowIndex:     idx,
      client:       String(row[PT_CLIENT]    || "").toUpperCase().trim(),
      startBalance: row[PT_START_BAL],
    }));

    const previews  = [];
    const unmatched = [];

    for (const payment of matched) {
      const match = tableClients.find(
        r => r.client === payment.client && r.client !== ""
      );
      if (match) {
        previews.push({
          client:       payment.client,
          amount:       payment.amount,
          rowIndex:     match.rowIndex,
          startBalance: match.startBalance,
          projectedEnd: (Number(match.startBalance) || 0) - payment.amount,
          excelRow:     payment.excelRow,
        });
      } else {
        unmatched.push(payment.client);
      }
    }

    pendingPaymentRows = previews;

    // Build dynamic preview table with projected end balance as extra column
    const previewHeaders = ["CLIENT", "START BALANCE", "PAYMENT AMT"];
    const previewRows    = previews.map(p => [
      p.client,
      `$${Number(p.startBalance || 0).toLocaleString()}`,
      `$${p.amount.toLocaleString()}`,
    ]);
    const extraColumns = [{
      label:    "PROJECTED END BAL",
      getValue: (row, idx) => `$${previews[idx]?.projectedEnd?.toLocaleString() ?? "—"}`,
    }];

    buildPreviewTable("preview-payments-table-container", previewHeaders, previewRows, extraColumns);

    // Show unmatched warning
    const warningEl = document.getElementById("payments-unmatched-warning");
    if (unmatched.length) {
      warningEl.textContent   = `⚠ ${unmatched.length} payment(s) had no match in PaymentsTable: ${unmatched.join(", ")}`;
      warningEl.style.display = "block";
    } else {
      warningEl.style.display = "none";
    }

    document.getElementById("import-payments-summary").style.display  = "none";
    document.getElementById("preview-payments-section").style.display = "block";
    document.getElementById("confirm-payments-btn").disabled          = !previews.length;

    hideLoading();
    setPaymentsStatus(
      `${previews.length} payment${previews.length !== 1 ? "s" : ""} matched and ready to apply.`
    );

  } catch (error) {
    hideLoading();
    setPaymentsStatus("Error: " + error.message);
    console.error(error);
  }
}

function cancelPaymentsImport() {
  pendingPaymentRows = [];
  ledgerSourceRows   = [];
  document.getElementById("preview-payments-section").style.display = "none";
  setPaymentsStatus("Payment import cancelled.");
}

// ── CONFIRM PAYMENTS IMPORT ───────────────────────────────────
async function confirmPaymentsImport() {
  if (!pendingPaymentRows.length) return;
  document.getElementById("confirm-payments-btn").disabled = true;
  showLoading("Applying payments...");

  const succeeded = [];
  const failed    = [];

  try {
    await Excel.run(async (context) => {
      const table     = context.workbook.tables.getItem("PaymentsTable");
      const bodyRange = table.getDataBodyRange();
      bodyRange.load("values, rowCount, columnCount");
      await context.sync();

      for (const payment of pendingPaymentRows) {
        try {
          const ri = payment.rowIndex;

          // ── 1. Write AMT OF PAYMENT ───────────────────────────
          // This triggers PAYMENT (col L) and ENDING BALANCE (col K) to recalc
          bodyRange.getCell(ri, PT_AMT_PAYMENT).values = [[payment.amount]];
          await context.sync();

          // ── 2. Read recalculated PAYMENT and ENDING BALANCE ───
          const paymentCell      = bodyRange.getCell(ri, PT_PAYMENT);
          const endingBalCell    = bodyRange.getCell(ri, PT_ENDING_BAL);
          const endBalCell       = bodyRange.getCell(ri, PT_END_BAL);
          paymentCell.load("values");
          endingBalCell.load("values");
          endBalCell.load("values");
          await context.sync();

          const newEndingBalance = endingBalCell.values[0][0]; // col K
          const newEndBalance    = endBalCell.values[0][0];    // col I

          // ── 3. Write ENDING BALANCE → BALANCE OWED (col J) ───
          bodyRange.getCell(ri, PT_BAL_OWED).values = [[newEndingBalance]];

          // ── 4. Write END BALANCE → START BALANCE (col G) ─────
          bodyRange.getCell(ri, PT_START_BAL).values = [[newEndBalance]];

          // ── 5. Clear AMT OF PAYMENT (col H) ──────────────────
          bodyRange.getCell(ri, PT_AMT_PAYMENT).values = [[""]];

          await context.sync();
          succeeded.push(payment);

        } catch (err) {
          failed.push({ payment, reason: err.message });
        }
      }
    });

    // Mark rows green + timestamp in source ledger
    if (succeeded.length && selectedPaymentsFileId) {
      showLoading("Marking rows as processed...");
      const rangeData     = await graphFetch(
        `items/${selectedPaymentsFileId}/workbook/worksheets/Sheet1/usedRange?$select=columnCount`
      );
      const stampColIndex = rangeData.columnCount || 12;
      const { initials }  = await getEmployeeInitials();
      const { sheetName } = await readSecretaryPayments(selectedPaymentsFileId, initials);
      const successExcelRows = succeeded.map(p => p.excelRow);
      await markRowsProcessed(
        selectedPaymentsFileId, sheetName, successExcelRows, stampColIndex
      );
    }

    let message = `✔ Applied ${succeeded.length} payment${succeeded.length !== 1 ? "s" : ""} and updated balances.`;
    if (failed.length) {
      message += `\n⚠ ${failed.length} skipped:\n`;
      message += failed.map(f => `• ${f.payment.client}: ${f.reason}`).join("\n");
    }

    showSummary("import-payments-summary", message, failed.length > 0 && !succeeded.length);
    pendingPaymentRows = [];
    document.getElementById("confirm-payments-btn").disabled = false;
    paymentsImportDone = true;
    document.getElementById("archive-payments-btn").style.display = "inline-block";
    document.getElementById("archive-payments-btn").disabled      = false;
    hideLoading();
    setPaymentsStatus("");

  } catch (error) {
    hideLoading();
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
  const el         = document.getElementById(elementId);
  el.className     = isError ? "import-summary error" : "import-summary";
  el.textContent   = message;
  el.style.display = "block";
}

function showLoading(message = "Working...") {
  document.getElementById("loading-text").textContent   = message;
  document.getElementById("loading-overlay").style.display = "flex";
}

function hideLoading() {
  document.getElementById("loading-overlay").style.display = "none";
}

function showNotification(message, type = "info") {
  const el       = document.getElementById("notification-bar");
  el.textContent = message;
  el.className   = `notification-bar ${type}`;
  el.style.display = "block";
  setTimeout(() => { el.style.display = "none"; }, 6000);
}