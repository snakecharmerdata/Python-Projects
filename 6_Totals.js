function replaceXWithSumOfValuesAbove() {
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
  
  // Get the number of rows with data
  var lastRow = sheet.getLastRow();
  
  // If there are fewer than 2 rows, there's only a header or nothing at all
  if (lastRow < 2) {
    Logger.log("Not enough data. Sheet has fewer than 2 rows.");
    SpreadsheetApp.getActiveSpreadsheet().toast("Not enough data. Need at least 2 rows.", "Warning", 5);
    return;
  }
  
  // Column H is the 8th column
  var columnHIndex = 8;
  
  try {
    // Get all values from column H (excluding header row)
    var columnHRange = sheet.getRange(2, columnHIndex, lastRow - 1, 1);
    var columnHValues = columnHRange.getValues();
    
    // Debug log to check if we're getting data correctly
    Logger.log("Retrieved " + columnHValues.length + " values from column H");
    
    // Track cells to update and their indices
    var cellsToUpdate = [];
    var cellIndices = [];
    
    // Track running sum of numeric values
    var runningSum = 0;
    
    // Process each cell in column H
    for (var i = 0; i < columnHValues.length; i++) {
      var currentValue = columnHValues[i][0];
      
      // Debug log for special values to help troubleshoot
      if (currentValue === 'x') {
        Logger.log("Found 'x' at row " + (i + 2) + ", current running sum: " + runningSum);
      }
      
      // If current cell contains an 'x' (ensuring case sensitivity and exact match)
      if (currentValue === 'x') {
        // Add this cell to the update list with the current running sum
        cellsToUpdate.push([runningSum]);
        cellIndices.push(i + 2); // +2 because we're starting from row 2 and array is 0-indexed
      } 
      // Otherwise, if it's a number, add it to our running sum
      else if (typeof currentValue === 'number') {
        runningSum += currentValue;
        Logger.log("Added number: " + currentValue + ", new running sum: " + runningSum);
      } 
      // Try to convert strings to numbers (in case they're formatted as text)
      else if (typeof currentValue === 'string') {
        // Remove any commas, currency symbols, etc.
        var cleanValue = currentValue.replace(/[$,]/g, '');
        var numValue = parseFloat(cleanValue);
        
        if (!isNaN(numValue)) {
          runningSum += numValue;
          Logger.log("Converted string '" + currentValue + "' to number: " + numValue + ", new running sum: " + runningSum);
        }
      }
    }
    
    // Log summary before updating
    Logger.log("Found " + cellsToUpdate.length + " 'x' values to replace");
    
    // Update all cells with 'x' to their respective sums - using batch update for efficiency
    if (cellsToUpdate.length > 0) {
      // Update cells in batches of 100 for better performance
      var batchSize = 100;
      
      for (var i = 0; i < cellIndices.length; i += batchSize) {
        var batchIndicesLength = Math.min(batchSize, cellIndices.length - i);
        var batchValues = cellsToUpdate.slice(i, i + batchIndicesLength);
        
        // Update each cell in this batch individually as they may not be contiguous
        for (var j = 0; j < batchIndicesLength; j++) {
          var rowIndex = cellIndices[i + j];
          var cellRange = sheet.getRange(rowIndex, columnHIndex);
          
          // Set the value
          cellRange.setValue(batchValues[j][0]);
          
          // Apply bold red formatting to the cell
          cellRange.setFontWeight("bold");
          cellRange.setFontColor("#FF0000");
        }
        
        // Yield occasionally to prevent timeout
        if (i > 0) {
          SpreadsheetApp.flush();
          Logger.log("Processed " + Math.min(i + batchIndicesLength, cellIndices.length) + " replacements so far");
        }
      }
    }
    
    // Success message
    var executionTime = (new Date() - startTime) / 1000;
    
    Logger.log("Replaced " + cellsToUpdate.length + " 'x' values with bold red sums in column H, completed in " + executionTime + " seconds.");
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Replaced " + cellsToUpdate.length + " 'x' values with bold red sums in column H in " + executionTime + " seconds", 
      "Success", 
      5
    );
  } catch (e) {
    Logger.log("Error: " + e.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Error replacing 'x' values: " + e.toString(), 
      "Error", 
      10
    );
  }
}
