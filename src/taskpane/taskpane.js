/* global console, document, Office */

let accessToken = null;

Office.onReady(() => {
  document.getElementById("sign-in-btn").onclick = signIn;
  document.getElementById("sign-out-btn").onclick = signOut;
  document.getElementById("read-file-btn").onclick = readFile;
});

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

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (msg) => {
        dialog.close();

        try {
          const data = JSON.parse(msg.message);

          if (data.status === "success") {
            accessToken = data.accessToken;
            showMainSection(data.userName);
            setStatus("");
          } else {
            setStatus("Sign in error: " + data.message);
          }
        } catch (e) {
          setStatus("Error parsing auth response.");
        }
      });

      dialog.addEventHandler(Office.EventType.DialogEventReceived, (evt) => {
        if (evt.error === 12006) {
          setStatus("Sign in was cancelled.");
        }
      });
    }
  );
}

function signOut() {
  accessToken = null;
  document.getElementById("sign-in-section").style.display = "block";
  document.getElementById("main-section").style.display = "none";
  setStatus("");
}

async function readFile() {
  const filePath = document.getElementById("file-path").value.trim();

  if (!filePath) {
    setStatus("Please enter a file path.");
    return;
  }

  if (!accessToken) {
    setStatus("Please sign in first.");
    return;
  }

  setStatus("Fetching file data...");
  document.getElementById("output-section").style.display = "none";

  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/root:${filePath}:/workbook/worksheets/Sheet1/usedRange`,
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