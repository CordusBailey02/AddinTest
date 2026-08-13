import {
  PT_AMT_PAYMENT,
  PT_PAYMENT,
  PT_ENDING_BAL,
  PT_END_BAL,
  PT_BAL_OWED,
  PT_START_BAL
} from "./employeeSubmission";

import { appState } from "./appState";

import { getEmployeeInitials } from "./employeeSettings";
import { graphFetch } from "./sharePoint";
import { 
  setPaymentsStatus,
  showSummary,
  showLoading,
  hideLoading,
  showNotification,
  buildPreviewTable
} from "./uiHelpers";

// Row offsets in source files (0-based index into .values array)

const LEDGER_HEADER_ROW      = 1; // row 2 in Excel (0-indexed)
const LEDGER_DATA_START      = 2; // row 3 in Excel (0-indexed)

// Secretary daily ledger column indices
const LEDGER_CLIENT_COL   = 0;
const LEDGER_AMOUNT_COL   = 1;
const LEDGER_BONDSMAN_COL = 5;

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


async function readSecretaryPayments(fileId, employeeInitials) {
  const sheetName = await resolveSheetName(fileId, SUBMISSION_SHEET_NAME);
  const data      = await graphFetch(
    `items/${fileId}/workbook/worksheets('${encodeURIComponent(sheetName)}')/usedRange`, appState.dailyLedgersId
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

async function previewPaymentsImport() {
  if (!appState.selectedPaymentsFileId) {
    setPaymentsStatus("Please select a daily ledger file first.");
    return;
  }

  showLoading("Reading daily ledger...");
  document.getElementById("preview-payments-section").style.display = "none";

  try {
    //const errors = await runPreflightChecks();
    //if (errors.length) {
    //  hideLoading();
    //  showNotification("⚠ Setup issues:\n• " + errors.join("\n• "), "error");
    //  return;
    //}

    const { initials, fullName } = await getEmployeeInitials();
    document.getElementById("detected-employee").textContent = `${fullName} (${initials})`;
    document.getElementById("employee-badge").style.display  = "block";

    const { matched, sourceRows, sheetName } = await readSecretaryPayments(
      appState.selectedPaymentsFileId, initials
    );

    if (!matched.length) {
      hideLoading();
      setPaymentsStatus(`No payments found for bondsman "${initials}" in this file.`);
      return;
    }

    appState.ledgerSourceRows = sourceRows;

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

    appState.pendingPaymentRows = previews;

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

      for (const payment of appState.pendingPaymentRows) {
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

    let message = `✔ Applied ${succeeded.length} payment${succeeded.length !== 1 ? "s" : ""} and updated balances.`;
    if (failed.length) {
      message += `\n⚠ ${failed.length} skipped:\n`;
      message += failed.map(f => `• ${f.payment.client}: ${f.reason}`).join("\n");
    }

    showSummary("import-payments-summary", message, failed.length > 0 && !succeeded.length);
    appState.pendingPaymentRows = [];
    document.getElementById("confirm-payments-btn").disabled = false;
    appState.paymentsImportDone = true;
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

function cancelPaymentsImport() {
  appState.pendingPaymentRows = [];
  appState.accessTokenledgerSourceRows   = [];
  document.getElementById("preview-payments-section").style.display = "none";
  setPaymentsStatus("Payment import cancelled.");
}

export {
    previewPaymentsImport,
    confirmPaymentsImport,
    cancelPaymentsImport
};