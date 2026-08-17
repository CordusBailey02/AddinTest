import {
  PAYMENT_PLAN_THRESHOLD,
  COMMISSION_PLAN,
  PAYMENT_PLAN,
  determineCommisionOrPaymentPlan,
  calcFullPayCommissionGTE5,
  calcFullPayCommissionLT5,
  calcPaymentPlanBond,
} from "./payrollCalculations";

import {
  getEmployeeSettings
} from "./employeeSettings";

import { appState } from "./appState";
import { graphFetch } from "./sharePoint";
import { 
  setStatus,
  showSummary,
  showLoading,
  hideLoading,
  buildPreviewTable 
} from "./uiHelpers.js";

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

const SUBMISSION_HEADER_ROW  = 2; // row 3 in Excel (0-indexed)
const SUBMISSION_DATA_START  = 3; // row 4 in Excel (0-indexed)


// ########################### READ EMPLOYEE SUBMITTED FILE FUNCTIONS ###########################

async function readEmployeeSubmissionData(file_id) {
  const settings = await getEmployeeSettings();
  const commission_rate = settings.commissionRate;
  // Fetch data from proper sheet in workbook
  const data      = await graphFetch(
    `items/${file_id}/workbook/worksheets('${encodeURIComponent('Sheet1')}')/usedRange`, appState.weeklyBondsId
  );

  const rows = data.values;
  const header_row = rows[SUBMISSION_HEADER_ROW] || [];
  const data_rows = rows.slice(SUBMISSION_DATA_START);

  const bonds = [];

  console.log("Printing raw row data from excel sheet:")
  console.log(data_rows)

  data_rows.forEach(row => {
    const date            = row[0]; // Column A - DATE
    const jail            = row[1]; // Column B - JAIL
    const last_name       = row[2]; // Column C - LAST NAME
    const first_name      = row[3]; // Column D - FIRST NAME
    const ttl_bond_amt    = row[4]; // Column E - TTL BOND AMOUNT
    const powers          = row[5]; // Column F - #
    const ttl_amt_chgd    = row[6]; // Column G - TOTAL AMT CHGD
    const ttl_amt_col     = row[7]; // Column H - TOTAL AMT COL
    const ttl_expense     = row[8]; // Column I - TTL EXPENSE
    const travel          = row[9]; // Column J - TRAVEL
    const amt_in_env      = row[10]; // Column K - AMT IN ENV
    const pymt_type       = row[11]; // Column L - PYMT TYPE
    const balance_owed    = row[12]; // Column M - BALANCE OWED
    const administration  = row[13]; // Column N - ADMINISTRATION

    // Skip rows that are empty
    if(!last_name || !first_name || !ttl_bond_amt) {
      return;
    }

    // Convert numeric values to numbers
    const ttl_bond_amt_num  = Number(ttl_bond_amt);
    const powers_num        = Number(powers);
    const ttl_amt_chgd_num  = Number(ttl_amt_chgd);
    const ttl_amt_col_num   = Number(ttl_amt_col);
    const ttl_expense_num   = Number(ttl_expense);
    const travel_num        = Number(travel);
    const amt_in_env_num    = Number(amt_in_env);
    const balance_owed_num  = Number(balance_owed);

    // Format date for display
    let displayDate = "";
    if (date instanceof Date) {
      displayDate = date.toLocaleDateString();
    } else if (typeof date === "number") {
      displayDate = new Date((date - 25569) * 86400 * 1000).toLocaleDateString();
    } else if (date) {
      displayDate = String(date);
    }

    const bond_path = determineCommisionOrPaymentPlan(ttl_bond_amt_num, balance_owed_num);

    let bond_record = {
      // Raw submission fields
      raw_date:           displayDate,
      raw_jail:           jail,
      raw_last_name:      String(last_name).toUpperCase(),
      raw_first_name:     String(first_name).toUpperCase(),
      raw_ttl_bond_amt:   ttl_bond_amt_num,
      raw_powers:         powers_num,
      raw_ttl_amt_chgd:   ttl_amt_chgd_num,
      raw_ttl_amt_col:    ttl_amt_col_num,
      raw_ttl_expense:    ttl_expense_num,
      raw_travel:         travel_num,
      raw_amt_in_env:     amt_in_env_num,
      raw_pymt_type:      pymt_type,
      raw_balance_owed:   balance_owed_num,
      raw_administration: administration,

      // Derived for payments sheet
      client_company:         `${String(last_name).toUpperCase()}, ${String(first_name).toUpperCase()}`,
      ttl_bond_company:       ttl_bond_amt_num,
      amt_charged_company:    ttl_amt_chgd_num,
      amt_collected_company:  ttl_amt_col_num,
      expense_company:        ttl_expense_num,
      //percent on bond_company (internal calculation)
      start_balance_company:  balance_owed_num,
      //amt of payment_company (from daily ledger)
      //end balance_company (FORMULA)
      //balance_owed_employee (internal calculation)
      //ending balance_employee (FORMULA)
      // payment_employee (FORMULA)           

      // Path determines where this bond goes
      bond_path, // "commission_plan" | "payment_plan"
    };

    if (bond_path === COMMISSION_PLAN) {
      let full_pay_amount = 0

      if(ttl_bond_amt_num >= 5000) {
        full_pay_amount = calcFullPayCommissionGTE5(ttl_bond_amt_num, commission_rate);
      }
      else {
        full_pay_amount = calcFullPayCommissionLT5(ttl_bond_amt_num, commission_rate);
      }
      // ── Path 1: Full pay — goes to TIMESHEET, not PAYMENTS ──
      bond_record = {
        ...bond_record,
        full_pay_amount,
        // No payment plan fields needed
        percent_paid_on_bond:   null,
        reach:                  null,
        employee_owed_balance:  null,
      };
    } else {
      // ── Path 2: Payment plan — goes into PAYMENTS sheet ──────
      const calc = calcPaymentPlanBond(
        ttl_bond_amt_num,
        ttl_expense_num,
        travel_num,
        ttl_amt_col_num,
        balance_owed_num,
        powers_num,
        commission_rate
      );
      bond_record = {
        ...bond_record,
        ...calc,
      };
    }

    bonds.push(bond_record);
  });

  return { bonds, header_row };
}

async function previewEmployeeSubmissionImport() {
  // Validate employee submission file was selected from file navigation
  if(!appState.selectedFileId) {
    setStatus("Please select a submision file for employee first.");
    return;
  }

  console.log(appState.selectedFileId)

  // Show loading before displaying preview section with data
  showLoading("Reading submission file...")
  document.getElementById("preview-section").style.display = "none";

  // ##### DO SOME PRE CHECKS TO VERIFY DOCUMENTS #####
  // pre-flight checks

  const { bonds, header_row } = await readEmployeeSubmissionData(appState.selectedFileId);

  // Check if any bonds data was pulled from the excel sheet
  if (!bonds.length) {
    hideLoading();
    setStatus("No bond entries found in the submission file.");
    return;
  }

  // Separate bonds by path
  const payment_plan_bonds = bonds.filter(b => b.bond_path === PAYMENT_PLAN);
  const full_pay_bonds     = bonds.filter(b => b.bond_path === COMMISSION_PLAN);

  // Only payment plan bonds go into pendingImportRows
  appState.pendingImportRows    = payment_plan_bonds;

  console.log("Printing the bonds read:")
  console.log(bonds)

  // ── Section 1: Raw submission data (all bonds) ────────────
  const submission_rows = bonds.map(b => [
    b.raw_date,
    b.raw_jail,
    b.raw_last_name,
    b.raw_first_name,
    `$${b.raw_ttl_bond_amt.toLocaleString()}`,
    b.raw_powers.toLocaleString(),
    `$${b.raw_ttl_amt_chgd.toLocaleString()}`,
    `$${b.raw_ttl_amt_col.toLocaleString()}`,
    `$${b.raw_ttl_expense.toLocaleString()}`,
    `$${b.raw_travel.toLocaleString()}`,
    `$${b.raw_amt_in_env.toLocaleString()}`,
    b.raw_pymt_type || "—",
    `$${b.raw_balance_owed.toLocaleString()}`,
    b.raw_administration || "—",
  ]);
  buildPreviewTable("preview-submission-table", header_row, submission_rows);

  // ── Section 2: Full pay bonds (Path 1) ───────────────────
  if (full_pay_bonds.length) {
    const fullPayHeaders = ["CLIENT", "TTL BOND", "COMMISSION", "NOTE"];
    const fullPayRows    = full_pay_bonds.map(b => [
      b.client_company,
      `$${b.ttl_bond_company.toLocaleString()}`,
      `$${b.full_pay_amount}`,
      b.ttl_bond_company < PAYMENT_PLAN_THRESHOLD
        ? "Under $5,000 — paid in full"
        : "Fully paid — paid in full",
    ]);
    buildPreviewTable("preview-fullpay-table", fullPayHeaders, fullPayRows);
    document.getElementById("fullpay-section").style.display = "block";
  } else {
    document.getElementById("fullpay-section").style.display = "none";
  }

  // ── Section 3: Payment plan bonds (Path 2) ───────────────
  if (payment_plan_bonds.length) {
    const paymentsHeaders = [
      "CLIENT", "TTL BOND", "AMT CHARGED", "AMT COLLECTED", "EXPENSE", "% PAID ON BOND", "START BALANCE", "AMT OF PAYMENT",
      "END BALANCE", "BALANCE OWED", "ENDING BALANCE", "PAYMENT", "REACH", "DOWN PAYMENT",
    ];
    const paymentsRows = payment_plan_bonds.map(b => [
      b.client_company,
      b.ttl_bond_company,
      b.amt_charged_company,
      b.amt_collected_company,
      b.expense_company,
      `${(b.percent_paid_on_bond * 100)}%`,
      `$${b.start_balance_company.toLocaleString()}`,
      b.amt_of_payment || "",
      `$${b.start_balance_company.toLocaleString()}`, // END BALANCE IS START BALANCE WHEN RECORD ENTERED
      `$${b.employee_owed_balance}`,
      `$${b.employee_owed_balance}`, // EMPLOYEE OWED BALANCE IS ENDING BALANCE WHEN RECORD ENTERED
      b.payment || "",
      b.reach_payment_goal > 0 ? `$${b.reach_payment_goal.toFixed(2)}` : "—",
      b.down_payment,
    ]);
    buildPreviewTable("preview-payments-calc-table", paymentsHeaders, paymentsRows);
    document.getElementById("payment-plan-section").style.display = "block";
  } else {
    document.getElementById("payment-plan-section").style.display = "none";
  }

  document.getElementById("import-summary").style.display  = "none";
  document.getElementById("preview-section").style.display = "block";
  document.getElementById("confirm-import-btn").disabled   = !payment_plan_bonds.length;

  // SHow the preview section and stop loading screen
  hideLoading();

  setStatus(
    `${bonds.length} bond${bonds.length > 1 ? "s" : ""} read — ` +
    `${payment_plan_bonds.length} payment plan, ${full_pay_bonds.length} full pay.`
  );
}

async function confirmEmployeeSubmissionImport() {
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
      for (const bond of appState.pendingImportRows) {
        try {
          const row = new Array(colCount).fill("");
          row[PT_CLIENT]      = bond.client_company;
          row[PT_TTL_BOND]    = bond.ttl_bond_company;
          row[PT_AMT_CHARGED] = bond.amt_charged_company;
          row[PT_AMT_COLLECT] = bond.amt_collected_company;
          row[PT_EXPENSE]     = bond.expense_company;
          row[PT_PCT_PAID]    = bond.percent_paid_on_bond;   // auto-calculated
          row[PT_START_BAL]   = bond.start_balance_company;
          row[PT_BAL_OWED]    = bond.employee_owed_balance;     // auto-calculated
          row[PT_REACH]       = bond.reach_payment_goal !== 0 ? bond.reach_payment_goal : ""
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

    let message = `✔ Successfully imported ${succeeded.length} bond${succeeded.length !== 1 ? "s" : ""}.`;
    if (failed.length) {
      message += `\n⚠ ${failed.length} row${failed.length !== 1 ? "s" : ""} skipped:\n`;
      message += failed.map(f => `• ${f.bond.client}: ${f.reason}`).join("\n");
    }

    showSummary("import-summary", message, failed.length > 0 && !succeeded.length);
    appState.pendingImportRows    = [];
    document.getElementById("confirm-import-btn").disabled          = false;
    appState.submissionImportDone = true;
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

function cancelEmployeeSubmissionImport() {
  appState.pendingImportRows    = [];
  appState.submissionSourceRows = [];
  document.getElementById("preview-section").style.display = "none";
  setStatus("Import cancelled.");
}

export {
    previewEmployeeSubmissionImport,
    confirmEmployeeSubmissionImport,
    cancelEmployeeSubmissionImport,
    PT_START_BAL,
    PT_AMT_PAYMENT,
    PT_END_BAL,
    PT_BAL_OWED,
    PT_ENDING_BAL,
    PT_PAYMENT,
    PT_CLIENT
};