/* global console, document, Office, Excel */

// ── CONFIG ────────────────────────────────────────────────────
const START_FOLDER_PATH = "BailBonds";

let accessToken = null;
let selectedFileId = null;
let pendingImportRows = [];
let breadcrumbStack = [{ id: "root", name: "My Files" }];

// ── OFFICE INIT ───────────────────────────────────────────────
Office.onReady(() => {
  document.getElementById("sign-in-btn").onclick = signIn;
  document.getElementById("sign-out-btn").onclick = signOut;
  document.getElementById("preview-btn").onclick = previewImport;
  document.getElementById("confirm-import-btn").onclick = confirmImport;
  document.getElementById("cancel-import-btn").onclick = cancelImport;
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
  pendingImportRows = [];
  breadcrumbStack = [{ id: "root", name: "My Files" }];
  document.getElementById("sign-in-section").style.display = "block";
  document.getElementById("main-section").style.display = "none";
  document.getElementById("preview-section").style.display = "none";
  setStatus("");
}

// ── START FOLDER ──────────────────────────────────────────────
async function navigateToStartFolder() {
  setStatus(`Opening ${START_FOLDER_PATH}...`);
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

// ── FILE BROWSER ──────────────────────────────────────────────
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

  // Row 1 = title, Row 2 = dates/agent, Row 3 = headers, Row 4+ = data
  const dataRows = rows.slice(3);

  const bonds = [];
  for (const row of dataRows) {
    const lastName  = row[2];  // Col C
    const firstName = row[3];  // Col D
    const ttlBond   = row[4];  // Col E
    const amtChgd   = row[6];  // Col G
    const amtCol    = row[7];  // Col H
    const expense   = row[8];  // Col I

    // Skip empty rows and totals row
    if (!lastName || !ttlBond) continue;
    if (String(lastName).toUpperCase().includes("TOTAL")) continue;

    const client = `${String(lastName).toUpperCase()}, ${String(firstName).toUpperCase()}`;
    const startBalance = (Number(amtChgd) || 0) - (Number(amtCol) || 0);

    bonds.push({
      client,
      ttlBond:       Number(ttlBond) || 0,
      amtCharged:    Number(amtChgd) || 0,
      amtCollected:  Number(amtCol)  || 0,
      expense:       Number(expense) || 0,
      startBalance,
    });
  }

  return bonds;
}

// ── PREVIEW IMPORT ────────────────────────────────────────────
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

    // Build preview table
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

// ── CONFIRM IMPORT — writes to currently open workbook ────────
async function confirmImport() {
  if (!pendingImportRows.length) return;

  document.getElementById("confirm-import-btn").disabled = true;
  setStatus("Importing...");

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("PAYMENTS");
      const usedRange = sheet.getUsedRange();
      usedRange.load("rowCount");
      await context.sync();

      // Used range includes header rows 1-4 and the totals row
      // Data rows start at row 5 (index 4). Totals is always the last used row.
      // New data goes before the totals row.
      const totalsRowIndex = usedRange.rowCount - 1; // 0-based index of TOTALS row
      let nextRowIndex = totalsRowIndex; // insert before TOTALS

      for (const bond of pendingImportRows) {
        // Insert a blank row before the TOTALS row so formulas in TOTALS expand
        const totalsRange = sheet.getRangeByIndexes(nextRowIndex, 0, 1, 13);
        totalsRange.insert(Excel.InsertShiftDirection.down);
        await context.sync();

        // Now write values to the newly inserted row
        const newRow = sheet.getRangeByIndexes(nextRowIndex, 0, 1, 13);
        newRow.values = [[
          bond.client,       // A - CLIENT
          bond.ttlBond,      // B - TTL BOND
          bond.amtCharged,   // C - AMT CHARGED
          bond.amtCollected, // D - AMT COLLECTED
          bond.expense,      // E - EXPENSE
          "",                // F - % PAID ON BOND (owner fills manually)
          bond.startBalance, // G - START BALANCE
          "",                // H - AMT OF PAYMENT (secretary later)
          "",                // I - END BALANCE (formula below)
          "",                // J - BALANCE OWED
          "",                // K - ENDING BALANCE (formula below)
          "",                // L - PAYMENT (formula below)
          "",                // M - REACH
        ]];

        // Write formulas — use Excel row number (1-based = index + 1)
        const excelRow = nextRowIndex + 1;
        const endBalCell    = sheet.getRangeByIndexes(nextRowIndex, 8, 1, 1);  // Col I
        const endingBalCell = sheet.getRangeByIndexes(nextRowIndex, 10, 1, 1); // Col K
        const paymentCell   = sheet.getRangeByIndexes(nextRowIndex, 11, 1, 1); // Col L

        endBalCell.formulas    = [[`=PAYMENTS!$G${excelRow}-PAYMENTS!$H${excelRow}`]];
        endingBalCell.formulas = [[`=PAYMENTS!$J${excelRow}-PAYMENTS!$L${excelRow}`]];
        paymentCell.formulas   = [[`=PAYMENTS!$H${excelRow}*PAYMENTS!$F${excelRow}`]];

        await context.sync();
        nextRowIndex++;
      }
    });

    const summary = document.getElementById("import-summary");
    summary.className = "import-summary";
    summary.textContent = `✔ Successfully imported ${pendingImportRows.length} bond${pendingImportRows.length > 1 ? "s" : ""} into PAYMENTS.`;
    summary.style.display = "block";
    pendingImportRows = [];
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