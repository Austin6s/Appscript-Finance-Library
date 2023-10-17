// Global variables for the active spreadsheet and folder
var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var activeFolder = DriveApp.getFileById(activeSpreadsheet.getId()).getParents().next();
var countries = ["Taiwan", "Philippines"];
var preservationConfig = {
  "Income": {
    "columnsToPreserve": [2, 7, 12, 17],
    "rowsToPreserve": [1, 2, 4],
    "cellsToPreserve": []
  },
  "Expenses": {
    "columnsToPreserve": [2],
    "rowsToPreserve": [1, 2, 4],
    "cellsToPreserve": []
  },
  "Sales Tracker": {
    "columnsToPreserve": [],
    "rowsToPreserve": [1, 2],
    "cellsToPreserve": []
  },
  "Expenses Tracker": {
    "columnsToPreserve": [14, 18],
    "rowsToPreserve": [1, 2],
    "cellsToPreserve": []
  }
  // Add more sheet configurations as needed
};

function clearDataWithRowAndColumnPreservation(targetSheet) {
  //If targetSheet is not provided, default to the active sheet.
  if (!targetSheet) {
    var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  }

  var activeSheetName = targetSheet.getName();
  var clearedCellAddresses = []; // Initialize an array to store cleared cell addresses
  
  // Check if the sheet name exists in the preservationConfig
  if (!preservationConfig.hasOwnProperty(activeSheetName)) {
    Logger.log("Sheet " + activeSheetName + " is not in the preservation configuration. Nothing to clear.");
    return;
  }
  
  var dataRange = targetSheet.getDataRange();
  var values = dataRange.getValues();
  var formulas = dataRange.getFormulas();
  
  var config = preservationConfig[activeSheetName];
  var columnsToPreserve = config.columnsToPreserve || [];
  var cellsToPreserve = config.cellsToPreserve || [];
  var rowsToPreserve = config.rowsToPreserve || [];
  
  Logger.log("Starting clearDataWithRowAndColumnPreservation for sheet: " + activeSheetName);
  
  for (var row = 0; row < values.length; row++) {
    for (var col = 0; col < values[row].length; col++) {
      var cellAddress = targetSheet.getRange(row + 1, col + 1).getA1Notation();
      var shouldClear = !columnsToPreserve.includes(col + 1) &&
                        !cellsToPreserve.includes(cellAddress) &&
                        !rowsToPreserve.includes(row + 1);
      
      //Logger.log("Checking row: " + (row + 1) + ", col: " + (col + 1) + ", shouldClear: " + shouldClear);
      
      if (shouldClear && formulas[row][col] === "") {
        var cellAddress = targetSheet.getRange(row + 1, col + 1).getA1Notation();
        clearedCellAddresses.push(cellAddress); // Store the cleared cell address
        targetSheet.getRange(row + 1, col + 1).clearContent();
        //Logger.log("Cleared cell at row: " + (row + 1) + ", col: " + (col + 1));
      }
    }
  }
  if (clearedCellAddresses.length > 0) {
    var startCellAddress = clearedCellAddresses[0];
    var endCellAddress = clearedCellAddresses[clearedCellAddresses.length - 1];
    var clearedRangeAddress = startCellAddress + ":" + endCellAddress;
    Logger.log("Cleared range address: " + clearedRangeAddress);
    // Now you have the A1 notation of the cleared range for non-empty cells
  }

  Logger.log("Finished clearDataWithRowAndColumnPreservation for sheet: " + activeSheetName);
  Logger.log(clearedCellAddresses)
  return clearedCellAddresses;
}


// Function to extract the month from the active spreadsheet's name
function extractMonth() {
  var spreadsheetName = activeSpreadsheet.getName();
  // Split the name by spaces and take the first word as the month
  var nameParts = spreadsheetName.split(" ");
  var month = nameParts[0];
  return month;
}


// Function to extract the year from the folder's name
function extractYear() {
  var year = activeFolder.getName();
  return year;
}


function getSourceData(month, year, country, clearedCellAddresses, sheetName) {
  // Get the grandparent folder (Financials folder)
  var grandparentFolder = activeFolder.getParents().next().getParents().next();
  Logger.log("Grandparent folder is: " + grandparentFolder);
  
  // Get the country folder (Taiwan Financials or Philippines Financials)
  Logger.log("country: " + country);
  var countryFolders = grandparentFolder.getFoldersByName(country);
  while (countryFolders.hasNext()) {
    var countryFolder = countryFolders.next();
    Logger.log("Folder name: " + countryFolder.getName());
  }

  // Get the parent folder (year folder)
  var yearFolder = countryFolder.getFoldersByName(year).next();
  
  // Get the source spreadsheet (September Financials)
  var sourceSpreadsheet = SpreadsheetApp.open(yearFolder.getFilesByName(month).next());
  
  // Access the "Income" sheet within the source spreadsheet
  var sourceSheet = sourceSpreadsheet.getSheetByName(sheetName); // Replace with the actual sheet name

  // Initialize an array to store the values from the specified cells
  var sourceDataValues = [];

  // Loop through the clearedCellAddresses and get values from each cell
  for (var i = 0; i < clearedCellAddresses.length; i++) {
    var cellAddress = clearedCellAddresses[i];
    var cellValue = sourceSheet.getRange(cellAddress).getValue();
    sourceDataValues.push([cellValue]); // Push the value as a single-element array
  }

  // Now, sourceDataValues contains the data from the specified cells in the source sheet
  Logger.log(sourceDataValues);

  // Determine the target sheet where you want to set the values
  var targetSheet = activeSpreadsheet.getSheetByName(sheetName);

  // Loop through the clearedCellAddresses and set values cell by cell
  for (var i = 0; i < clearedCellAddresses.length; i++) {
    var cellAddress = clearedCellAddresses[i];
    var cellValue = sourceDataValues[i][0]; // Get the value from sourceDataValues
    var targetCell = targetSheet.getRange(cellAddress);
    targetCell.setValue(cellValue); // Set the value in the corresponding cell
  }

}


function importData(){
  var month = extractMonth();
  var year = extractYear();  
  for (var key in preservationConfig){
    var sheet = activeSpreadsheet.getSheetByName(key);   
    var clearedCellAddresses = clearDataWithRowAndColumnPreservation(sheet);
    getSourceData(month, year, countries[1], clearedCellAddresses, key);
  }
  // for (var i = 0; i = countries.length; i++){
  //   getSourceData(month, year, countries[i], clearedCellAddresses);
  // }
}
