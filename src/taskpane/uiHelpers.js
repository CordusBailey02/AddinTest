

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


export {
  showMainSection,
  setStatus,
  setPaymentsStatus,
  showSummary,
  showLoading,
  hideLoading,
  showNotification,
  buildPreviewTable
};