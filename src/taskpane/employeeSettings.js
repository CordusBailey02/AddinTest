

// ── READ EMPLOYEE SETTINGS ────────────────────────────────────
async function getEmployeeSettings() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("EMPLOYEE SETTINGS");

    // Read B3:B10 — all value cells in one call
    const valueRange = sheet.getRange("B3:B10");
    valueRange.load("values");
    await context.sync();

    const v = valueRange.values;

    return {
      name:           String(v[0][0] || ""),   // B3
      employeeNumber: String(v[1][0] || ""),   // B4
      department:     String(v[2][0] || ""),   // B5
      title:          String(v[3][0] || ""),   // B6
      payRate:        Number(v[4][0]) || 0,    // B7
      supervisor:     String(v[5][0] || ""),   // B8
      email:          String(v[6][0] || ""),   // B9
      commissionRate: Number(v[7][0]) || 0,    // B10
    };
  });
}

async function getEmployeeInitials() {
  return await Excel.run(async (context) => {
    const sheet    = context.workbook.worksheets.getItem("EMPLOYEE SETTINGS");
    const nameCell = sheet.getRange("B3");
    nameCell.load("values");
    await context.sync();

    const fullName = nameCell.values[0][0];
    if (!fullName) throw new Error("Could not read employee name from EMPLOYEE SETTINGS sheet.");

    const parts        = String(fullName).split(",").map(p => p.trim());
    const lastInitial  = parts[0]?.[0]?.toUpperCase() || "";
    const firstInitial = parts[1]?.[0]?.toUpperCase() || "";

    return {
      initials: lastInitial + firstInitial,
      fullName:  String(fullName).trim(),
    };
  });
}

export {
  getEmployeeSettings,
  getEmployeeInitials
};