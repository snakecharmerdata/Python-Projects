
/**
 * Fills blank cells in column H with an 'x' character in sheet S2,
 * starting from row 2. This function preserves all other data.
 */
function fillBlankCellsInColumnH() {
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
  
  // Column H is the 8th column (index 7 in 0-based arrays)
  var columnHIndex = 8;
  
  // Process data in batches to avoid memory issues
  var batchSize = 1000; // Process 1000 rows at a time
  var cellsModified = 0;
  
  try {
    // Process the sheet in batches
    for (var startRow = 2; startRow <= lastRow; startRow += batchSize) {
      // Calculate how many rows to process in this batch
      var numRows = Math.min(batchSize, lastRow - startRow + 1);
      
      // Get the data for column H in this batch
      var range = sheet.getRange(startRow, columnHIndex, numRows, 1);
      var values = range.getValues();
      
      // Track which cells need to be modified
      var cellsToModify = [];
      
      // Check each cell in column H
      for (var i = 0; i < values.length; i++) {
        // If the cell is blank (empty or null)
        if (values[i][0] === "" || values[i][0] === null) {
          // Change the value to 'x'
          values[i][0] = "x";
          
          // Add to the list of cells that were modified
          cellsToModify.push(startRow + i);
          cellsModified++;
        }
      }
      
      // Write the updated values back to the sheet
      range.setValues(values);
      
      // Log progress
      if (cellsToModify.length > 0) {
        Logger.log("Modified " + cellsToModify.length + " cells in rows " + startRow + " to " + (startRow + numRows - 1));
      }
      
      // Yield to prevent timeout on very large datasets
      if (numRows === batchSize) {
        SpreadsheetApp.flush();
      }
    }
    
    // Success message
    var executionTime = (new Date() - startTime) / 1000;
    
    Logger.log("Added 'x' to " + cellsModified + " blank cells in column H, completed in " + executionTime + " seconds.");
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Added 'x' to " + cellsModified + " blank cells in column H in " + executionTime + " seconds", 
      "Success", 
      5
    );
  } catch (e) {
    Logger.log("Error: " + e.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Error filling blank cells: " + e.toString(), 
      "Error", 
      10
    );
  }
}