/**
 * This script collects, processes, and sorts data from multiple target sheets,
 * and writes it to the main sheet in a single operation to minimize update time.
 * If the "from" and "to" times are missing, they default to "00:00:00".
 */
function collectAndSortWorkerData() {
  try {
    // Access the main spreadsheet and sheets
    const mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const parametersSheet = mainSpreadsheet.getSheetByName("Parameters");
    const mainDataSheet = mainSpreadsheet.getSheetByName("Data");
    const parameters = parametersSheet.getDataRange().getValues();

    // Create a map to store hourly cost from the "Parameters" sheet
    const costMap = {};
    parameters.slice(1).forEach(row => {
      const link = row[0]; // Target sheet link
      const hourlyCost = parseFloat(row[7]) || 0; // Hourly cost (column 8 in the Parameters sheet)
      if (link) {
        costMap[link] = hourlyCost;
      }
    });

    // Define headers for the consolidated data
    const headers = [
      "Source Sheet Link", "Owner Name",
      "Event Type", "Date", "Start Time", "End Time",
      "Global Amount", "Notes", "Travel Cost",
      "From Timestamp", "To Timestamp", "Time Difference (Hours)", "Work Cost", "Total Cost"
    ];

    // Collect all processed data rows
    const processedData = [headers];

    // Iterate over each row in the "Parameters" sheet
    for (let i = 1; i < parameters.length; i++) {
      const targetSheetLink = parameters[i][0]; // Link to the target Google Sheet
      const ownerName = parameters[i][2]; // Owner's name

      if (!targetSheetLink) {
        console.log(`Row ${i + 1}: Target sheet link is empty. Skipping.`);
        continue;
      }

      try {
        // Access the target spreadsheet and its "Data" sheet
        const targetSpreadsheet = SpreadsheetApp.openByUrl(targetSheetLink);
        const targetDataSheet = targetSpreadsheet.getSheetByName("Data");

        if (!targetDataSheet) {
          console.log(`Row ${i + 1}: No "Data" sheet found in the target spreadsheet. Skipping.`);
          continue;
        }

        const targetData = targetDataSheet.getDataRange().getValues();

        // Process each row in the target sheet's "Data"
        for (let rowIndex = 1; rowIndex < targetData.length; rowIndex++) {
          const row = targetData[rowIndex];

          // Extract relevant columns
          const [eventType, date, fromHour, toHour, globalAmount, notes, travel] = row;

          // Default missing time values
          const validFromHour = fromHour ? fromHour.toString() : "00:00:00";
          const validToHour = toHour ? toHour.toString() : "00:00:00";

          // Parse and calculate timestamps, time differences, and costs
          let fromTimestamp = null, toTimestamp = null, timeDifference = null, workCost = null, totalCost = null;
          if (date) {
            try {
              const parsedDate = new Date(date);

              // Parse "from" timestamp
              fromTimestamp = parseTimestamp(parsedDate, validFromHour);

              // Parse "to" timestamp
              toTimestamp = parseTimestamp(parsedDate, validToHour);

              // Calculate the time difference in hours
              timeDifference = (toTimestamp - fromTimestamp) / (1000 * 60 * 60); // Convert ms to hours
            } catch (error) {
              console.error(`Error parsing date/time for row ${rowIndex + 1}: ${error.message}`);
            }
          }

          // Calculate work and total costs
          workCost = calculateWorkCost(globalAmount, timeDifference, costMap[targetSheetLink]);
          const travelCost = parseFloat(travel) || 0;
          totalCost = (workCost || 0) + travelCost;

          // Append processed row
          const newRow = [
            targetSheetLink, ownerName,
            eventType || "", date || "", fromHour || "", toHour || "",
            globalAmount || "", notes || "", travel || "",
            fromTimestamp, toTimestamp, timeDifference || "", workCost || "", totalCost || ""
          ];
          processedData.push(newRow);
        }

        console.log(`Row ${i + 1}: Data from target sheet processed successfully.`);
      } catch (error) {
        console.error(`Row ${i + 1}: Error processing target sheet. ${error.message}`);
      }
    }

    // Sort data by "From Timestamp"
    processedData.slice(1).sort((a, b) => {
      if (a[9] && b[9]) return new Date(a[9]) - new Date(b[9]);
      if (!a[9]) return 1;
      if (!b[9]) return -1;
      return 0;
    });

    // Write sorted data to the main "Data" sheet
    mainDataSheet.clear();
    mainDataSheet.getRange(1, 1, processedData.length, processedData[0].length).setValues(processedData);

    console.log("All data successfully processed and written to the main sheet.");
  } catch (error) {
    console.error(`An error occurred during data processing: ${error.message}`);
  }
}

/**
 * Parses a timestamp from a given date and time string.
 * @param {Date} date - The base date.
 * @param {string} timeString - The time string to parse.
 * @returns {Date} - The parsed timestamp.
 */
function parseTimestamp(date, timeString) {
  const timestamp = new Date(date);
  const timeParts = timeString.match(/(\d+):(\d+):(\d+)\s?(AM|PM)?/);
  if (timeParts) {
    let hours = parseInt(timeParts[1], 10);
    const minutes = parseInt(timeParts[2], 10);
    const seconds = parseInt(timeParts[3], 10);
    if (timeParts[4]?.toUpperCase() === "PM" && hours < 12) hours += 12;
    if (timeParts[4]?.toUpperCase() === "AM" && hours === 12) hours = 0;
    timestamp.setHours(hours, minutes, seconds);
  }
  return timestamp;
}

/**
 * Calculates the work cost based on global amount, time difference, and hourly rate.
 * @param {number} globalAmount - Global cost for the event.
 * @param {number} timeDifference - Time difference in hours.
 * @param {number} hourlyRate - Hourly cost rate.
 * @returns {number} - The calculated work cost.
 */
function calculateWorkCost(globalAmount, timeDifference, hourlyRate) {
  if (globalAmount > 0) return globalAmount;
  if (timeDifference > 0 && hourlyRate) return timeDifference * hourlyRate;
  return 0;
}
