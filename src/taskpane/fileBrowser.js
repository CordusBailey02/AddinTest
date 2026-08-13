import { appState } from "./appState";
import { graphFetch } from "./sharePoint";
import {
  setStatus,
  setPaymentsStatus
} from "./uiHelpers";

// Browser config
const browsers = {
  submission: {
    get driveId() {
      return appState.weeklyBondsId;
    },
    stackId: "breadcrumb",
    browserId: "file-browser",
    stack: [{ id: "root", name: "Weekly Bonds" }],
    onSelect: onEmployeeSubmissionFileSelected,
  },

  payments: {
    get driveId() {
      return appState.dailyLedgersId;
    },
    stackId: "payments-breadcrumb",
    browserId: "payments-file-browser",
    stack: [{ id: "root", name: "Daily Ledgers" }],
    onSelect: onPaymentsFileSelected,
  },
};

async function browseFolder(browserKey, folderId) {
  const b = browsers[browserKey];
  const browser = document.getElementById(b.browserId);

  browser.innerHTML = "<div style='padding:12px;color:gray;'>Loading...</div>";

  if (browserKey === "submission") {
    appState.selectedFileId = null;
    document.getElementById("selected-file").style.display = "none";
    document.getElementById("preview-btn").style.display = "none";
    document.getElementById("preview-section").style.display = "none";
    document.getElementById("archive-submission-btn").style.display = "none";
    appState.submissionImportDone = false;
  } else {
    appState.selectedPaymentsFileId = null;
    document.getElementById("selected-payments-file").style.display = "none";
    document.getElementById("preview-payments-btn").style.display = "none";
    document.getElementById("preview-payments-section").style.display = "none";
    document.getElementById("archive-payments-btn").style.display = "none";
    appState.paymentsImportDone = false;
  }

  try {
    const url = folderId === "root"
      ? "root/children"
      : `items/${folderId}/children`;

    const data = await graphFetch(
      url + "?$orderby=name&$select=id,name,folder,file,size,parentReference",
      b.driveId
    );

    const items = data.value;

    if (!items.length) {
      browser.innerHTML =
        "<div style='padding:12px;color:gray;'>This library is empty.</div>";
      return;
    }

    browser.innerHTML = "";

    items.forEach((item) => {
      const isFolder = !!item.folder;
      const isExcel =
        item.name?.endsWith(".xlsx") ||
        item.name?.endsWith(".xls");

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
          b.stack.push({
            id: item.id,
            name: item.name
          });

          renderBreadcrumb(browserKey);
          browseFolder(browserKey, item.id);

        } else {
          browser
            .querySelectorAll(".browser-item")
            .forEach(el => el.classList.remove("selected"));

          div.classList.add("selected");
          b.onSelect(item);
        }
      };

      browser.appendChild(div);
    });

    renderBreadcrumb(browserKey);

  } catch (error) {
    browser.innerHTML =
      `<div style='padding:12px;color:red;'>Error: ${error.message}</div>`;
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

// Function to handle when an employee submitted file is selected 
function onEmployeeSubmissionFileSelected(item) {
  appState.selectedFileId       = item.id;
  appState.selectedFileName     = item.name;
  appState.selectedFileParentId = item.parentReference?.id || null;
  document.getElementById("selected-file-name").textContent = item.name;
  document.getElementById("selected-file").style.display    = "block";
  document.getElementById("preview-btn").style.display      = "inline-block";
  document.getElementById("archive-submission-btn").style.display = "none";
  appState.submissionImportDone = false;
  setStatus("");
}

// Function to handle when a daily ledger file is selected
function onPaymentsFileSelected(item) {
  appState.selectedPaymentsFileId       = item.id;
  appState.selectedPaymentsFileName     = item.name;
  appState.selectedPaymentsFileParentId = item.parentReference?.id || null;
  document.getElementById("selected-payments-file-name").textContent = item.name;
  document.getElementById("selected-payments-file").style.display    = "block";
  document.getElementById("preview-payments-btn").style.display      = "inline-block";
  document.getElementById("archive-payments-btn").style.display      = "none";
  appState.paymentsImportDone = false;
  setPaymentsStatus("");
}

export {
  browsers,
  browseFolder
};