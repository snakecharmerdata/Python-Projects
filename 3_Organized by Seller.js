/**
 * Organizes data in sheet S2 by unique values in column A, starting from row 2.
 * This function:
 * 1. Preserves the header row (row 1)
 * 2. Sorts the data by column A values
 * 3. Groups rows with the same values in column A together
 */
function organizeS2ByUniqueNamesInColumnA() {
  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the S2 sheet
  var sheet = ss.getSheetByName("S2");
  
  // Check if sheet exists
  if (!sheet) {
    Logger.log("Sheet 'S2' not found.");
    SpreadsheetApp.getActiveSpreadsheet().toast("Sheet 'S2' not found", "Error", 5);
    return;
  }
  
  // Get the number of rows and columns with data
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  
  // If there are fewer than 2 rows, there's only a header or nothing at all
  if (lastRow < 2) {
    Logger.log("Not enough data to organize. Sheet has fewer than 2 rows.");
    SpreadsheetApp.getActiveSpreadsheet().toast("Not enough data to organize. Need at least 2 rows.", "Warning", 5);
    return;
  }
  
  // Get all data including headers (we'll separate them later)
  var allData = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  
  // Extract headers (row 1)
  var headers = allData[0];
  
  // Extract data rows (row 2 and beyond)
  var dataRows = allData.slice(1);
  
  // Create a map to store rows by their column A values
  var rowsByColumnA = {};
  
  // Group rows by their column A values
  for (var i = 0; i < dataRows.length; i++) {
    var row = dataRows[i];
    var columnAValue = row[0]; // Column A is index 0
    
    // Skip empty values in column A
    if (columnAValue === "" || columnAValue === null) {
      continue;
    }
    
    // Convert to string for consistent handling
    columnAValue = String(columnAValue);
    
    // Initialize array for this column A value if it doesn't exist
    if (!rowsByColumnA[columnAValue]) {
      rowsByColumnA[columnAValue] = [];
    }
    
    // Add this row to the appropriate group
    rowsByColumnA[columnAValue].push(row);
  }
  
  // Get unique column A values and sort them alphabetically
  var uniqueColumnAValues = Object.keys(rowsByColumnA).sort();
  
  // Create a new array for organized data
  var organizedData = [headers]; // Start with the headers
  
  // Add rows in order of sorted unique column A values
  for (var i = 0; i < uniqueColumnAValues.length; i++) {
    var columnAValue = uniqueColumnAValues[i];
    var rows = rowsByColumnA[columnAValue];
    
    // Add all rows for this column A value to the organized data
    for (var j = 0; j < rows.length; j++) {
      organizedData.push(rows[j]);
    }
  }
  
  // Clear the sheet data (except formulas and formatting)
  sheet.clearContents();
  
  // Write the organized data back to the sheet
  if (organizedData.length > 0) {
    sheet.getRange(1, 1, organizedData.length, organizedData[0].length).setValues(organizedData);
    Logger.log("Data in S2 has been organized by unique values in column A.");
    SpreadsheetApp.getActiveSpreadsheet().toast("Data in S2 has been organized by unique values in column A.", "Success", 5);
  } else {
    Logger.log("No data to organize.");
    SpreadsheetApp.getActiveSpreadsheet().toast("No data to organize.", "Warning", 5);
  }
  
  // Force the spreadsheet to update
  SpreadsheetApp.flush();
}