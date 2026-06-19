/* global console, document, Office */

let accessToken = null;
let selectedFileId = null;
let breadcrumbStack = [{ id: "root", name: "My Files" }];

Office.onReady(() => {
  document.getElementById("sign-in-btn").onclick = signIn;
  document.getElementById("sign-out-btn").onclick = signOut;
  document.getElementById("read-file-btn").onclick = readFile;
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
            await browseFolder("root");
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
  breadcrumbStack = [{ id: "root", name: "My Files" }];
  document.getElementById("sign-in-section").style.display = "block";
  document.getElementById("main-section").style.display = "none";
  setStatus("");
}

// ── FILE BROWSER ──────────────────────────────────────────────
async function browseFolder(folderId) {
  const browser = document.getElementById("file-browser");
  browser.innerHTML = "<div style='padding:12px;color:gray;'>Loading...</div>";

  // Clear selected file when navigating
  selectedFileId = null;
  document.getElementById("selected-file").style.display = "none";
  document.getElementById("read-file-btn").style.display = "none";
  document.getElementById("output-section").style.display = "none";

  try {
    const url = folderId === "root"
      ? "https://graph.microsoft.com/v1.0/me/drive/root/children"
      : `https://graph.microsoft.com/v1.0/me/drive/items/${folderId}/children`;

    const response = await fetch(url + "?$orderby=name&$select=id,name,folder,file,size", {
      headers: { Authorization: `Bearer ${accessToken}` },
    });

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

      // Only show folders and Excel files
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
          // Select this file
          document.querySelectorAll(".browser-item").forEach(el => el.classList.remove("selected"));
          div.classList.add("selected");
          selectedFileId = item.id;
          document.getElementById("selected-file-name").textContent = item.name;
          document.getElementById("selected-file").style.display = "block";
          document.getElementById("read-file-btn").style.display = "inline-block";
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
    if (index === breadcrumbStack.length - 1) {
      return `📁 ${crumb.name}`;
    }
    return `<span data-index="${index}">${crumb.name}</span> › `;
  }).join("");

  // Add click handlers for breadcrumb navigation
  breadcrumb.querySelectorAll("span").forEach((span) => {
    span.onclick = () => {
      const index = parseInt(span.getAttribute("data-index"));
      breadcrumbStack = breadcrumbStack.slice(0, index + 1);
      browseFolder(breadcrumbStack[breadcrumbStack.length - 1].id);
    };
  });
}

// ── READ FILE ─────────────────────────────────────────────────
async function readFile() {
  if (!selectedFileId) {
    setStatus("Please select a file first.");
    return;
  }

  setStatus("Reading file...");
  document.getElementById("output-section").style.display = "none";

  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/items/${selectedFileId}/workbook/worksheets/Sheet1/usedRange`,
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      }
    );

    if (!response.ok) {
      const err = await response.json();
      setStatus(`Error: ${err.error?.message || response.statusText}`);
      return;
    }

    const data = await response.json();
    displayTable(data.values);
    setStatus("Data loaded successfully.");
  } catch (error) {
    setStatus("Failed to read file: " + error.message);
    console.error(error);
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

function displayTable(values) {
  if (!values || values.length === 0) {
    document.getElementById("output").innerHTML = "<p>No data found.</p>";
    document.getElementById("output-section").style.display = "block";
    return;
  }

  const table = document.createElement("table");
  values.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    row.forEach((cell) => {
      const td = document.createElement(rowIndex === 0 ? "th" : "td");
      td.textContent = cell;
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });

  document.getElementById("output").innerHTML = "";
  document.getElementById("output").appendChild(table);
  document.getElementById("output-section").style.display = "block";
}