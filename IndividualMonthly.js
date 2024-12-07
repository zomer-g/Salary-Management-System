/**
 * Generates a monthly financial report for each owner and writes it to their respective "Summary" sheet.
 * Aggregates total costs and the number of workdays per month for each owner.
 */
function generateIndividualMonthlyReports() {
  try {
    // Access the main spreadsheet and the "Data" sheet
    const mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const mainDataSheet = mainSpreadsheet.getSheetByName("Data");
    const allData = mainDataSheet.getDataRange().getValues();

    // Map to store monthly summaries for each owner
    const ownerReports = {};

    // Extract headers and data rows
    const headers = allData[0];
    const rows = allData.slice(1);

    // Identify column indices based on header names
    const ownerCol = headers.indexOf("Owner Name");
    const sourceSheetLinkCol = headers.indexOf("Source Sheet Link");
    const dateCol = headers.indexOf("Date");
    const totalCostCol = headers.indexOf("Total Cost");

    // Validate required columns
    if (ownerCol === -1 || sourceSheetLinkCol === -1 || dateCol === -1 || totalCostCol === -1) {
      console.error("Required columns are missing in the 'Data' sheet.");
      return;
    }

    // Aggregate data by owner and month
    rows.forEach(row => {
      const ownerName = row[ownerCol];
      const sourceSheetLink = row[sourceSheetLinkCol];
      const date = row[dateCol];
      const totalCost = parseFloat(row[totalCostCol]) || 0;

      // Skip rows with missing data
      if (!ownerName || !sourceSheetLink || !date) return;

      // Parse the date and extract the month and year
      const parsedDate = new Date(date);
      if (isNaN(parsedDate)) return;

      const monthYear = `${parsedDate.getMonth() + 1}/${parsedDate.getFullYear()}`;

      // Initialize owner data structure if not already present
      if (!ownerReports[ownerName]) {
        ownerReports[ownerName] = {
          sourceSheetLink,
          monthlyData: {},
        };
      }

      // Initialize month data if not already present
      if (!ownerReports[ownerName].monthlyData[monthYear]) {
        ownerReports[ownerName].monthlyData[monthYear] = {
          totalCost: 0,
          daysCount: new Set(),
        };
      }

      // Update monthly totals and add the day to the set
      ownerReports[ownerName].monthlyData[monthYear].totalCost += totalCost;
      ownerReports[ownerName].monthlyData[monthYear].daysCount.add(parsedDate.toDateString());
    });

    // Write the summary data to each owner's "Summary" sheet
    for (const [ownerName, reportData] of Object.entries(ownerReports)) {
      try {
        // Open the target spreadsheet using the source link
        const targetSpreadsheet = SpreadsheetApp.openByUrl(reportData.sourceSheetLink);

        // Access or create the "Summary" sheet
        let summarySheet = targetSpreadsheet.getSheetByName("Summary");
        if (!summarySheet) {
          summarySheet = targetSpreadsheet.insertSheet("Summary");
        } else {
          summarySheet.clear();
        }

        // Prepare the summary data for writing
        const summaryHeaders = ["Month and Year", "Total Cost", "Number of Days"];
        const summaryData = [summaryHeaders];

        for (const [monthYear, data] of Object.entries(reportData.monthlyData)) {
          summaryData.push([
            monthYear,
            data.totalCost,
            data.daysCount.size, // Count of unique days
          ]);
        }

        // Write the data to the "Summary" sheet
        summarySheet.getRange(1, 1, summaryData.length, summaryData[0].length).setValues(summaryData);

        console.log(`Monthly report written successfully for ${ownerName}.`);
      } catch (error) {
        console.error(`Error writing summary for ${ownerName}: ${error.message}`);
      }
    }

    console.log("All monthly reports generated successfully.");
  } catch (error) {
    console.error(`An error occurred while generating monthly reports: ${error.message}`);
  }
}
