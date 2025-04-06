function executeAllScripts() {
  var functionOrder = [
    'capitalizeFirstLettersInRow1(sheetName)',
    'transferDataFromS1ToS2AltMapping',
    'organizeS2ByUniqueNamesInColumnA',
    'separateGroupsWithEmptyRowsFixed',
    'fillBlankCellsInColumnH',
    'replaceXWithSumOfValuesAbove'
  ];
  
  try {
    Logger.log('Starting script execution: ' + new Date());
    
    for (var i = 0; i < functionOrder.length; i++) {
      var functionName = functionOrder[i];
      
      if (typeof this[functionName] === 'function') {
        Logger.log('Executing: ' + functionName);
        
        // Handle special case for first function which requires a parameter
        if (functionName === 'capitalizeFirstLettersInRow1') {
          this[functionName]('S2'); // Assuming 'S2' is the target sheet name
        } else {
          this[functionName]();
        }
        
        Logger.log(functionName + ' completed');
      } else {
        Logger.log('Warning: Function ' + functionName + ' not found');
      }
    }
    
    Logger.log('All scripts completed successfully: ' + new Date());
    SpreadsheetApp.getActiveSpreadsheet().toast('All scripts executed successfully!');
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast('Error: ' + error.toString(), 'Script Error', 10);
  } 
}

/**
 * Creates a custom menu when the spreadsheet is opened
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Execute Scripts')
    .addItem('Run All Scripts', 'executeAllScripts')
    .addToUi();
}