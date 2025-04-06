/**
 * Capitalizes the first letter of every word in row 1 on the specified sheet
 * while maintaining the column order.
 * 
 * @param {string} sheetName - The name of the sheet to modify (default: "Sheet1")
 */
function capitalizeFirstLettersInRow1(sheetName) {
  // If no sheet name is provided, use Sheet1 as default
  sheetName = sheetName || "S1";
  
  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the specified sheet
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log("Sheet '" + sheetName + "' not found.");
    return;
  }
  
  // Get the number of columns in the sheet
  var lastColumn = sheet.getLastColumn();
  
  if (lastColumn <= 0) {
    Logger.log("Sheet is empty. No data to modify.");
    return;
  }
  
  // Get row 1 data
  var row1Data = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  Logger.log("Original row 1 data: " + JSON.stringify(row1Data));
  
  // Process each cell in row 1
  var modified = false;
  
  for (var i = 0; i < row1Data.length; i++) {
    // Convert to string if it's not already
    var cellValue = String(row1Data[i]);
    
    if (cellValue && cellValue.trim() !== '') {
      var originalValue = cellValue;
      
      // More robust word splitting - handles multiple spaces, tabs, etc.
      var words = cellValue.split(/\s+/);
      
      // Capitalize the first letter of each word
      for (var j = 0; j < words.length; j++) {
        if (words[j].length > 0) {
          // Extract first character and the rest of the word
          var firstChar = words[j].charAt(0);
          var restOfWord = words[j].substring(1);
          
          // Convert first char to uppercase and combine with rest of word
          words[j] = firstChar.toUpperCase() + restOfWord;
        }
      }
      
      // Join the words back together with a single space between them
      var newValue = words.join(' ');
      
      // Update the row data if there's been a change
      if (newValue !== originalValue) {
        row1Data[i] = newValue;
        modified = true;
        Logger.log("Cell " + (i+1) + " changed from '" + originalValue + "' to '" + newValue + "'");
      }
    }
  }
  
  // Write the modified data back to row 1 only if changes were made
  if (modified) {
    sheet.getRange(1, 1, 1, row1Data.length).setValues([row1Data]);
    Logger.log("First letter of each word in row 1 has been capitalized in " + sheetName + ".");
  } else {
    Logger.log("No changes were made to row 1 in " + sheetName + ".");
  }
  
  // Force the spreadsheet to update
  SpreadsheetApp.flush();
  
  return modified;
}

/**
 * Helper function to call capitalizeFirstLettersInRow1 for Sheet2
 */
function capitalizeFirstLettersInSheet2() {
  var result = capitalizeFirstLettersInRow1("Sheet2");
  if (result) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Successfully capitalized row 1 in Sheet2", "Success", 5);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("No changes were needed in row 1 of Sheet2", "Information", 5);
  }
}

/**
 * Helper function to call capitalizeFirstLettersInRow1 for Sheet1
 */
function capitalizeFirstLettersInSheet1() {
  var result = capitalizeFirstLettersInRow1("Sheet1");
  if (result) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Successfully capitalized row 1 in Sheet1", "Success", 5);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("No changes were needed in row 1 of Sheet1", "Information", 5);
  }
}