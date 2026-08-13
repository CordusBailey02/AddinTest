
// GOOD
import {
  signIn,
  signOut,
  refreshTokenViaDialog
} from "./authentication";

// GOOD
import {
  previewEmployeeSubmissionImport,
  confirmEmployeeSubmissionImport,
  cancelEmployeeSubmissionImport
} from "./employeeSubmission";

// GOOD
import {
  previewPaymentsImport,
  confirmPaymentsImport,
  cancelPaymentsImport
} from "./dailyLedger";

// ── OFFICE INIT ───────────────────────────────────────────────
Office.onReady(() => {
  document.getElementById("sign-in-btn").onclick            = signIn;
  document.getElementById("sign-out-btn").onclick           = signOut;
  document.getElementById("refresh-token-btn").onclick      = refreshTokenViaDialog;

  document.getElementById("preview-btn").onclick            = previewEmployeeSubmissionImport;
  document.getElementById("confirm-import-btn").onclick     = confirmEmployeeSubmissionImport;
  document.getElementById("cancel-import-btn").onclick      = cancelEmployeeSubmissionImport;
  //document.getElementById("archive-submission-btn").onclick = archiveSubmissionFile;

  document.getElementById("preview-payments-btn").onclick   = previewPaymentsImport;
  document.getElementById("confirm-payments-btn").onclick   = confirmPaymentsImport;
  document.getElementById("cancel-payments-btn").onclick    = cancelPaymentsImport;
  //document.getElementById("archive-payments-btn").onclick   = archivePaymentsFile;

});


/*
async function runPreflightChecks() {
  return await Excel.run(async (context) => {
    const errors = [];

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
    if (!sheetNames.includes("EMPLOYEE SETTINGS")) {
      errors.push("EMPLOYEE SETTINGS sheet not found. Please add it to this workbook.");
    }
    if (!tableNames.includes("PaymentsTable")) {
      errors.push("PaymentsTable not found. Please convert the PAYMENTS data range to a table named 'PaymentsTable'.");
    }

    if (sheetNames.includes("PAYMENTS")) {
      const sheet    = context.workbook.worksheets.getItem("PAYMENTS");
      const nameCell = sheet.getRange("B2");
      nameCell.load("values");
      await context.sync();
      if (!nameCell.values[0][0]) {
        errors.push("Cell B2 on PAYMENTS sheet is empty. Check that EMPLOYEE SETTINGS!B3 has the employee name.");
      }
    }

    return errors;
  });
}
*/