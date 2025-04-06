
/**
 * Fixed function that organizes data in sheet S2 by unique values in column A,
 * inserts empty rows between different groups, and stops at two consecutive empty rows.
 * This version handles ALL rows in the spreadsheet without arbitrary limits.
 */
function separateGroupsWithEmptyRowsFixed() {
  // Start timing the execution
  var startTime = new Date();
  
  // Get the active spreadsheet and the S2 sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("S2");
  
  // Check if sheet exists
  if (!sheet) {
    Logger.log("Sheet 'S2' not found.");
    SpreadsheetApp.getActiveSpreadsheet().toast("Sheet 'S2' not found", "Error", 5);
    return;
  }
  
  // Get the actual number of rows and columns with data - NO LIMITS
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  
  // Log for debugging
  Logger.log("Processing sheet with " + lastRow + " rows and " + lastColumn + " columns");
  
  // If there are fewer than 2 rows, there's only a header or nothing at all
  if (lastRow < 2) {
    Logger.log("Not enough data to organize. Sheet has fewer than 2 rows.");
    SpreadsheetApp.getActiveSpreadsheet().toast("Not enough data to organize. Need at least 2 rows.", "Warning", 5);
    return;
  }
  
  // Get data in batches if there are many rows to avoid memory issues
  var allData = [];
  var batchSize = 1000; // Process 1000 rows at a time, but don't limit total rows
  
  // Get header row separately
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  allData.push(headers);
  
  // Process data in batches for better memory management
  for (var i = 2; i <= lastRow; i += batchSize) {
    var rowsToGet = Math.min(batchSize, lastRow - i + 1);
    var batchData = sheet.getRange(i, 1, rowsToGet, lastColumn).getValues();
    allData = allData.concat(batchData);
    
    // Yield to prevent timeout on very large datasets
    if (rowsToGet === batchSize) {
      SpreadsheetApp.flush();
      Logger.log("Processed rows " + i + " to " + (i + rowsToGet - 1));
    }
  }
  
  // Check for two consecutive empty rows to find where to stop processing
  var stopAtIndex = -1;
  for (var i = 1; i < allData.length - 1; i++) {
    if (isEmptyRow(allData[i]) && isEmptyRow(allData[i+1])) {
      stopAtIndex = i;
      Logger.log("Found two consecutive empty rows at positions " + (i+1) + " and " + (i+2) + ". Will stop processing here.");
      break;
    }
  }
  
  // If we found two consecutive empty rows, limit our processing
  var processedData;
  if (stopAtIndex > 0) {
    processedData = allData.slice(1, stopAtIndex);
    Logger.log("Stopping at row " + (stopAtIndex + 1) + " due to two consecutive empty rows");
  } else {
    processedData = allData.slice(1);
  }
  
  // Group the data by column A values
  var groups = {};
  var uniqueValues = [];
  
  for (var i = 0; i < processedData.length; i++) {
    var row = processedData[i];
    
    // Skip empty rows
    if (isEmptyRow(row)) continue;
    
    var columnAValue = String(row[0] || "");
    
    // First time seeing this value?
    if (!groups[columnAValue]) {
      groups[columnAValue] = [];
      uniqueValues.push(columnAValue);
    }
    
    // Add this row to its group
    groups[columnAValue].push(row);
  }
  
  // Sort unique values alphabetically for consistent output
  uniqueValues.sort();
  
  // Create new worksheet data with headers and grouped data
  var newData = [headers]; // Start with headers
  
  // Add each group's data with empty rows between groups
  for (var i = 0; i < uniqueValues.length; i++) {
    var value = uniqueValues[i];
    var groupRows = groups[value];
    
    // Add all rows for this group
    for (var j = 0; j < groupRows.length; j++) {
      newData.push(groupRows[j]);
    }
    
    // Add an empty row after the group (except for the last group)
    if (i < uniqueValues.length - 1) {
      newData.push(Array(lastColumn).fill(""));
    }
  }
  
  // Apply the changes in batches to handle large datasets
  try {
    // Clear the sheet first
    sheet.clear();
    
    // Write the header row
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Write data in batches if there's a lot
    if (newData.length > 1) {
      var dataToWrite = newData.slice(1); // Skip header as we've already written it
      
      // Set maximum batch size for writing
      var writeBatchSize = 1000;
      
      for (var i = 0; i < dataToWrite.length; i += writeBatchSize) {
        var batchToWrite = dataToWrite.slice(i, i + writeBatchSize);
        if (batchToWrite.length > 0) {
          // row index is i+2 because we're 0-indexed, and already wrote row 1 (header)
          sheet.getRange(i + 2, 1, batchToWrite.length, lastColumn).setValues(batchToWrite);
          SpreadsheetApp.flush(); // Force update to avoid timeout
          Logger.log("Wrote rows " + (i + 2) + " to " + (i + batchToWrite.length + 1));
        }
      }
    }
    
    // Success message
    var executionTime = (new Date() - startTime) / 1000;
    var rowsProcessed = stopAtIndex > 0 ? stopAtIndex : processedData.length;
    
    Logger.log("Data organization completed in " + executionTime + " seconds. Processed " + rowsProcessed + " rows.");
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Data organized with empty rows between groups. Processed " + rowsProcessed + " rows in " + executionTime + " seconds", 
      "Success", 
      8
    );
  } catch (e) {
    Logger.log("Error: " + e.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Error organizing data: " + e.toString(), 
      "Error", 
      10
    );
  }
}

/**
 * Faster helper function to check if a row is empty
 */
function isEmptyRow(row) {
  for (var i = 0; i < row.length; i++) {
    // Check if cell has any content
    if (row[i] !== "" && row[i] !== null) {
      return false;
    }
  }
  return true;
}