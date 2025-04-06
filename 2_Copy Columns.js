
/**
 * Transfers data from S1 to S2 based on an alternative set of column mappings:
 * - Column A (S1) -> Column F (S2)
 * - Column B (S1) -> Column B (S2)
 * - Column C (S1) -> Column C (S2)
 * - Column D (S1) -> Column E (S2)
 * - Column E (S1) -> Column D (S2)
 * - Column J (S1) -> Column G (S2)
 * - Column M (S1) -> Column A (S2)
 * - Column O (S1) -> Column H (S2)
 */
function transferDataFromS1ToS2AltMapping() {
  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the source and destination sheets
  var sourceSheet = ss.getSheetByName("S1");
  var destSheet = ss.getSheetByName("S2");
  
  // Check if both sheets exist
  if (!sourceSheet) {
    Logger.log("Source sheet 'S1' not found.");
    SpreadsheetApp.getActiveSpreadsheet().toast("Source sheet 'S1' not found", "Error", 5);
    return;
  }
  
  if (!destSheet) {
    Logger.log("Destination sheet 'S2' not found.");
    SpreadsheetApp.getActiveSpreadsheet().toast("Destination sheet 'S2' not found", "Error", 5);
    return;
  }
  
  // Get all data from S1
  var sourceData = sourceSheet.getDataRange().getValues();
  
  // If S1 is empty, exit the function
  if (sourceData.length <= 0) {
    Logger.log("Source sheet is empty. Nothing to transfer.");
    SpreadsheetApp.getActiveSpreadsheet().toast("Source sheet is empty. Nothing to transfer.", "Warning", 5);
    return;
  }
  
  // Create array to hold the mapped data for S2
  var destData = [];
  
  // Loop through each row in the source data
  for (var i = 0; i < sourceData.length; i++) {
    // Create a new row for the destination data with the new mapping
    // The order here determines which column in S2 the data goes to
    var newRow = [
      sourceData[i][12], // Column M (index 12) from S1 -> Column A in S2
      sourceData[i][1],  // Column B (index 1) from S1 -> Column B in S2
      sourceData[i][2],  // Column C (index 2) from S1 -> Column C in S2
      sourceData[i][4],  // Column E (index 4) from S1 -> Column D in S2
      sourceData[i][3],  // Column D (index 3) from S1 -> Column E in S2
      sourceData[i][0],  // Column A (index 0) from S1 -> Column F in S2
      sourceData[i][9],  // Column J (index 9) from S1 -> Column G in S2
      sourceData[i][14]  // Column O (index 14) from S1 -> Column H in S2
    ];
    
    // Add the new row to the destination data
    destData.push(newRow);
  }
  
  // Clear existing data in the destination sheet (optional - remove if you want to append)
  destSheet.clear();
  
  // Write the mapped data to S2
  if (destData.length > 0) {
    destSheet.getRange(1, 1, destData.length, destData[0].length).setValues(destData);
    Logger.log("Data transfer from S1 to S2 with alternative mapping completed successfully!");
    SpreadsheetApp.getActiveSpreadsheet().toast("Data transfer with alternative mapping completed successfully!", "Success", 5);
  } else {
    Logger.log("No data to transfer.");
    SpreadsheetApp.getActiveSpreadsheet().toast("No data to transfer.", "Warning", 5);
  }
  
  // Force the spreadsheet to update
  SpreadsheetApp.flush();
}
