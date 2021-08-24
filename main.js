// Constant variables for percentages, tax rates, etc
const mnTaxRate = 0.09;
const quartlyBonusFundPercentage = 0.05;
const foodCostPercentage = 0.475;
const employeeGratuityPayoutPercentage = 0.475;

function CalculatePayroll() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let formS = ss.getSheetByName("Form");
  let accountingS = ss.getSheetByName("Accounting History");
  let payrollPeriod = formS.getRange("C5").getValue();
  let grossTotal = formS.getRange("C3").getValue();
  let taxesPaid = (grossTotal * mnTaxRate).toFixed(2);
  let netTotal = (grossTotal - taxesPaid).toFixed(2);
  let employeeGratuityPayout = (
    netTotal * employeeGratuityPayoutPercentage
  ).toFixed(2);

  // Run function to calculate employee payout
  EmployeePayout(payrollPeriod, employeeGratuityPayout);

  // Clear inputs
  ClearCell();
}

// Function clears all cells in the form
function ClearCell() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let formS = ss.getSheetByName("Form");

  let rangesToClear = [
    "C3",
    "C5", // Inputs
    "C8",
    "C9",
    "C10",
    "C11",
    "C12",
    "C13",
    "C14",
    "C15",
    "C16",
    "C17",
    "C18", // Employee Hours Worked
  ];

  // Loop over Range and clear contents
  for (let i = 0; i < rangesToClear.length; i++) {
    formS.getRange(rangesToClear[i]).clearContent();
  }
}

function insertRow(sheet, rowData) {
  let lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    sheet
      .insertRowBefore(2)
      .getRange(2, 1, 1, rowData.length)
      .setValues([rowData]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function EmployeePayout(payrollPeriod, employeeGratuityPayout) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let formS = ss.getSheetByName("Form");
  let payoutS = ss.getSheetByName("Employee Payout History");
  let totalHoursWorked = 0;
  let employeeNameRange = [
    "B8",
    "B9",
    "B10",
    "B11",
    "B12",
    "B13",
    "B14",
    "B15",
    "B16",
    "B17",
    "B18",
  ];
  let employeeHoursRange = [
    "C8",
    "C9",
    "C10",
    "C11",
    "C12",
    "C13",
    "C14",
    "C15",
    "C16",
    "C17",
    "C18",
  ];

  // Loop over hours worked range to find total hours worked
  for (let i = 0; i < employeeHoursRange.length; i++) {
    let employeeHours = formS.getRange(employeeHoursRange[i]).getValue();
    totalHoursWorked += employeeHours;
  }

  // Calculate hourly gratuity based on total hours worked and payout %
  let hourlyGratuity = (employeeGratuityPayout / totalHoursWorked).toFixed(2);

  // Loop over employee range, append each employee record to Payout History table
  for (let i = 0; i < employeeNameRange.length; i++) {
    let employeeName = formS.getRange(employeeNameRange[i]).getValue();
    let employeeHours = formS.getRange(employeeHoursRange[i]).getValue();
    let hourlyGratuityPayout = hourlyGratuity * employeeHours;

    // Validate inputs, append rows
    if (employeeName !== "" && employeeHours !== "") {
      let rowData = [
        payrollPeriod,
        employeeName,
        employeeHours,
        hourlyGratuityPayout,
      ];
      insertRow(payoutS, rowData);
    }
  }
}
