/**
 * This script generates a "Monthly" report summarizing financial data for each worker.
 * It calculates "Deserves" (amount owed), "Got" (amount paid), and "Difference" (balance) by month and year.
 * Workers marked as hidden in the "Parameters" sheet are excluded from the report.
 * The "Difference" rows are formatted in bold for easy identification.
 */
function generateMonthlyReport() {
  try {
    // Access all relevant sheets
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const parametersSheet = spreadsheet.getSheetByName("Parameters");
    const dataSheet = spreadsheet.getSheetByName("Data");
    const paymentsSheet = spreadsheet.getSheetByName("Payments");
    const monthlySheet = spreadsheet.getSheetByName("Monthly");

    // Validate that all required sheets are present
    if (!parametersSheet || !dataSheet || !paymentsSheet || !monthlySheet) {
      throw new Error("One or more required sheets (Parameters, Data, Payments, Monthly) are missing.");
    }

    // Fetch data from sheets
    const parameters = parametersSheet.getDataRange().getValues();
    const data = dataSheet.getDataRange().getValues();
    const payments = paymentsSheet.getDataRange().getValues();

    // Get column indices from "Parameters" sheet
    const parametersHeaders = parameters[0];
    const ownerNameIndex = parametersHeaders.indexOf("Full Name");
    const hideMonthlyIndex = parametersHeaders.indexOf("Hide in Monthly Report");

    if (ownerNameIndex === -1 || hideMonthlyIndex === -1) {
      throw new Error("Required columns ('Full Name' or 'Hide in Monthly Report') are missing in 'Parameters' sheet.");
    }

    // Get column indices from "Data" sheet
    const dataHeaders = data[0];
    const fromIndex = dataHeaders.indexOf("From Timestamp");
    const dataOwnerIndex = dataHeaders.indexOf("Owner Name");
    const totalCostIndex = dataHeaders.indexOf("Total Cost");

    if (fromIndex === -1 || dataOwnerIndex === -1 || totalCostIndex === -1) {
      if (fromIndex === -1) console.error("Column 'From Timestamp' is missing in 'Data' sheet.");
      if (dataOwnerIndex === -1) console.error("Column 'Owner Name' is missing in 'Data' sheet.");
      if (totalCostIndex === -1) console.error("Column 'Total Cost' is missing in 'Data' sheet.");
      throw new Error("Required columns are missing in 'Data' sheet.");
    }

    // Get column indices from "Payments" sheet
    const paymentsHeaders = payments[0];
    const paymentNameIndex = paymentsHeaders.indexOf("Employee Name");
    const monthIndex = paymentsHeaders.indexOf("Payment Month");
    const yearIndex = paymentsHeaders.indexOf("Payment Year");
    const paymentAmountIndex = paymentsHeaders.indexOf("Amount");

    if (paymentNameIndex === -1 || monthIndex === -1 || yearIndex === -1 || paymentAmountIndex === -1) {
      throw new Error("Required columns are missing in 'Payments' sheet.");
    }

    // Step 1: Identify owners to exclude
    const hiddenOwners = new Set();
    parameters.slice(1).forEach(row => {
      if (row[hideMonthlyIndex] === true || row[hideMonthlyIndex] === "TRUE") {
        hiddenOwners.add(row[ownerNameIndex]);
      }
    });

    // Step 2: Extract unique months/years from "Data" sheet
    const monthYearSet = new Set();
    const ownerSet = new Set();

    data.slice(1).forEach(row => {
      const fromDate = row[fromIndex];
      const ownerName = row[dataOwnerIndex];
      if (fromDate && !hiddenOwners.has(ownerName)) {
        const date = new Date(fromDate);
        const monthYear = `${date.getMonth() + 1}/${date.getFullYear()}`;
        monthYearSet.add(monthYear);
      }
      if (ownerName && !hiddenOwners.has(ownerName)) {
        ownerSet.add(ownerName);
      }
    });

    const monthYears = Array.from(monthYearSet).sort((a, b) => {
      const [aMonth, aYear] = a.split("/").map(Number);
      const [bMonth, bYear] = b.split("/").map(Number);
      return aYear !== bYear ? aYear - bYear : aMonth - bMonth;
    });

    const owners = Array.from(ownerSet).sort();

    // Step 3: Prepare the output array
    const output = [];

    // Header row with month-year columns
    const headerRow = ["Owner Name", "Category", ...monthYears];
    output.push(headerRow);

    // Rows for each owner
    const differenceRows = [];
    owners.forEach(owner => {
      const deservesRow = [owner, "Deserves", ...monthYears.map(() => 0)];
      const gotRow = [owner, "Got", ...monthYears.map(() => 0)];
      const differenceRow = [owner, "Difference", ...monthYears.map(() => 0)];

      // Calculate "Deserves" amounts
      data.slice(1).forEach(row => {
        const fromDate = row[fromIndex];
        const totalCost = parseFloat(row[totalCostIndex]) || 0;
        const rowOwner = row[dataOwnerIndex];
        if (fromDate && rowOwner === owner) {
          const date = new Date(fromDate);
          const monthYear = `${date.getMonth() + 1}/${date.getFullYear()}`;
          const columnIndex = headerRow.indexOf(monthYear);
          if (columnIndex > -1) {
            deservesRow[columnIndex] += totalCost;
          }
        }
      });

      // Calculate "Got" amounts
      payments.slice(1).forEach(row => {
        const paymentName = row[paymentNameIndex];
        const paymentMonth = row[monthIndex];
        const paymentYear = row[yearIndex];
        const paymentAmount = parseFloat(row[paymentAmountIndex]) || 0;

        if (paymentName === owner) {
          const monthYear = `${paymentMonth}/${paymentYear}`;
          const columnIndex = headerRow.indexOf(monthYear);
          if (columnIndex > -1) {
            gotRow[columnIndex] += paymentAmount;
          }
        }
      });

      // Calculate "Difference" amounts
      for (let i = 2; i < headerRow.length; i++) {
        differenceRow[i] = deservesRow[i] - gotRow[i];
      }

      // Add rows to the output array
      output.push(deservesRow, gotRow, differenceRow);

      // Track "Difference" rows for formatting
      differenceRows.push(output.length - 1);
    });

    // Step 4: Write the report to the "Monthly" sheet
    monthlySheet.clear();
    monthlySheet.getRange(1, 1, output.length, output[0].length).setValues(output);

    // Step 5: Apply bold formatting to "Difference" rows
    differenceRows.forEach(rowIndex => {
      const range = monthlySheet.getRange(rowIndex + 1, 1, 1, output[0].length);
      range.setFontWeight("bold");
    });

    console.log("Monthly report generated successfully with bold formatting.");
  } catch (error) {
    console.error(`An error occurred: ${error.message}`);
  }
}
