/* global console, document, Office */
import { PublicClientApplication } from "@azure/msal-browser";

// ── CONFIG ────────────────────────────────────────────────────
const msalConfig = {
  auth: {
    clientId: "0c472194-49de-476f-ad1f-8cd689bc60e9",
    authority: "https://login.microsoftonline.com/f285c292-6d27-443d-a1f0-dd70bcc4de16",
    redirectUri: "https://cordusbailey02.github.io/AddinTest/taskpane.html",
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
};

const loginRequest = {
  scopes: ["User.Read", "Files.Read", "Files.Read.All", "Sites.Read.All"],
};

const msalInstance = new PublicClientApplication(msalConfig);

// ── OFFICE INIT ───────────────────────────────────────────────
Office.onReady(async () => {
  await msalInstance.initialize();

  // Handle redirect response if returning from login
  await msalInstance.handleRedirectPromise();

  document.getElementById("sign-in-btn").onclick = signIn;
  document.getElementById("sign-out-btn").onclick = signOut;
  document.getElementById("read-file-btn").onclick = readFile;

  // If already signed in, go straight to main section
  const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  if (account) {
    msalInstance.setActiveAccount(account);
    showMainSection(account.name || account.username);
  }
});

// ── AUTH ──────────────────────────────────────────────────────
async function signIn() {
  try {
    setStatus("Signing in...");
    const result = await msalInstance.loginPopup(loginRequest);
    msalInstance.setActiveAccount(result.account);
    showMainSection(result.account.name || result.account.username);
    setStatus("");
  } catch (error) {
    setStatus("Sign in failed: " + error.message);
    console.error(error);
  }
}

function signOut() {
  msalInstance.logoutPopup();
  document.getElementById("sign-in-section").style.display = "block";
  document.getElementById("main-section").style.display = "none";
}

async function getToken() {
  const account = msalInstance.getActiveAccount();
  try {
    // Try silent first
    const result = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
    return result.accessToken;
  } catch {
    // Fall back to popup
    const result = await msalInstance.acquireTokenPopup(loginRequest);
    return result.accessToken;
  }
}

// ── READ SHAREPOINT FILE ──────────────────────────────────────
async function readFile() {
  const filePath = document.getElementById("file-path").value.trim();

  if (!filePath) {
    setStatus("Please enter a file path.");
    return;
  }

  setStatus("Fetching file data...");
  document.getElementById("output-section").style.display = "none";

  try {
    const token = await getToken();

    // Get the default SharePoint drive for the tenant root
    // File path should be like: /sites/SiteName/Shared Documents/File.xlsx
    const encodedPath = encodeURIComponent(filePath);

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/root:${filePath}:/workbook/worksheets/Sheet1/usedRange`,
      {
        headers: { Authorization: `Bearer ${token}` },
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
  table.style.borderCollapse = "collapse";
  table.style.width = "100%";
  table.style.fontSize = "12px";

  values.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    row.forEach((cell) => {
      const td = document.createElement(rowIndex === 0 ? "th" : "td");
      td.textContent = cell;
      td.style.border = "1px solid #ccc";
      td.style.padding = "4px 8px";
      if (rowIndex === 0) td.style.backgroundColor = "#f3f3f3";
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });

  document.getElementById("output").innerHTML = "";
  document.getElementById("output").appendChild(table);
  document.getElementById("output-section").style.display = "block";
}