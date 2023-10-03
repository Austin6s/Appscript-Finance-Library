function getSubTypeRange(selectedType) {
  var dataListSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Lists");
  
  var ranges = {
    "Marketing": "Marketing",
    "Manufacturing": "Manufacturing",
    "Contracting": "Contracting",
    "Administrative": "Administrative",
    "Personnel": "Personnel",
    "Miscellaneous": "Miscellaneous",
    "Tournament": "Tournament",
    "Commission": "Commission",
    "Inventory": "Inventory"
  };
  Logger.log(selectedType);
  if (ranges[selectedType]) {
    var subTypeRange = dataListSheet.getRange(ranges[selectedType]);
    return subTypeRange;
  }
  
  return null;
}

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

function clearDataWithRowAndColumnPreservation() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheetName = sheet.getName();
  
  // Check if the sheet name exists in the preservationConfig
  if (!preservationConfig.hasOwnProperty(sheetName)) {
    Logger.log("Sheet " + sheetName + " is not in the preservation configuration. Nothing to clear.");
    return;
  }
  
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var formulas = dataRange.getFormulas();
  
  var config = preservationConfig[sheetName];
  var columnsToPreserve = config.columnsToPreserve || [];
  var cellsToPreserve = config.cellsToPreserve || [];
  var rowsToPreserve = config.rowsToPreserve || [];
  
  Logger.log("Starting clearDataWithRowAndColumnPreservation for sheet: " + sheetName);
  
  for (var row = 0; row < values.length; row++) {
    for (var col = 0; col < values[row].length; col++) {
      var cellAddress = sheet.getRange(row + 1, col + 1).getA1Notation();
      var shouldClear = !columnsToPreserve.includes(col + 1) &&
                        !cellsToPreserve.includes(cellAddress) &&
                        !rowsToPreserve.includes(row + 1);
      
      Logger.log("Checking row: " + (row + 1) + ", col: " + (col + 1) + ", shouldClear: " + shouldClear);
      
      if (shouldClear && formulas[row][col] === "") {
        sheet.getRange(row + 1, col + 1).clearContent();
        Logger.log("Cleared cell at row: " + (row + 1) + ", col: " + (col + 1));
      }
    }
  }
  
  Logger.log("Finished clearDataWithRowAndColumnPreservation for sheet: " + sheetName);
}

function onOpen() {
  createCustomMenu();
}

function createCustomMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Reset Data', 'clearDataWithRowAndColumnPreservation')
    .addToUi();
}


function setCurrencyFormatForRange() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Summary'); // Set the "Summary" sheet as the active sheet.
  var summaryRange = sheet.getRange('C5:E9')
  var grossRevRange = sheet.getRange('Income!H5:J20'); // Adjust the range as needed.
  var grossCogsRange = sheet.getRange('Income!M5:O20');
  var grossProfitRange = sheet.getRange('Income!R5:T20');
  var expensesRange = sheet.getRange('Expenses!C5:E55');
  var salesTrackerRange1 = sheet.getRange('Sales Tracker!H3:I');
  var salesTrackerRange2 = sheet.getRange('Sales Tracker!K3:R');
  var salesTrackerRunningTotalRange = sheet.getRange('Sales Tracker!T2:V2');
  var expensesTrackerRunningTotalRange = sheet.getRange('Expenses Tracker!K1:K2');
  var expensesTrackerCommissionRange = sheet.getRange('Expenses Tracker!N3:O');
  var expensesTrackerRange = sheet.getRange('Expenses Tracker!G3:H');
  // Get the currency code value from cell A1.
  var currencyCode = sheet.getRange('F1').getValue().toUpperCase(); // Convert to uppercase for case-insensitivity.

  // Define number formats for different currencies.
  var numberFormats = {
  'PHP': '₱#,##0.00', // Philippine Peso
  'TWD': 'NT$#,##0.00', // New Taiwan Dollar
  'JPY': '¥#,##0', // Japanese Yen
  'HKD': 'HK$#,##0.00', // Hong Kong Dollar
  'IDR': 'Rp #,##0', // Indonesian Rupiah
  // Add more currency formats as needed
};

  // Check if the currency code exists in the numberFormats object.
    if (numberFormats.hasOwnProperty(currencyCode)) {
    // Set the number format for each range based on the currency code.
    summaryRange.setNumberFormat(numberFormats[currencyCode]);
    grossRevRange.setNumberFormat(numberFormats[currencyCode]);
    grossCogsRange.setNumberFormat(numberFormats[currencyCode]);
    grossProfitRange.setNumberFormat(numberFormats[currencyCode]);
    expensesRange.setNumberFormat(numberFormats[currencyCode]);
    salesTrackerRange1.setNumberFormat(numberFormats[currencyCode]);
    salesTrackerRange2.setNumberFormat(numberFormats[currencyCode]);
    salesTrackerRunningTotalRange.setNumberFormat(numberFormats[currencyCode]);
    expensesTrackerRunningTotalRange.setNumberFormat(numberFormats[currencyCode]);
    expensesTrackerCommissionRange.setNumberFormat(numberFormats[currencyCode]);
    expensesTrackerRange.setNumberFormat(numberFormats[currencyCode]);
    // Add more ranges if needed.
  } else {
    // If the currency code is not recognized, set a default number format for each range.
    summaryRange.setNumberFormat('$#,##0.00');
    grossRevRange.setNumberFormat('$#,##0.00');
    grossCogsRange.setNumberFormat('$#,##0.00');
    grossProfitRange.setNumberFormat('$#,##0.00');
    expensesRange.setNumberFormat('$#,##0.00');
    salesTrackerRange1.setNumberFormat('$#,##0.00');
    salesTrackerRange2.setNumberFormat('$#,##0.00');
    salesTrackerRunningTotalRange.setNumberFormat('$#,##0.00');
    expensesTrackerRunningTotalRange.setNumberFormat('$#,##0.00');
    expensesTrackerCommissionRange.setNumberFormat('$#,##0.00');
    expensesTrackerRange.setNumberFormat('$#,##0.00');
    // Add more ranges if needed.
  }
}

function populateStartingInventory() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentMonth = activeSpreadsheet.getName(); // Get the name of the active spreadsheet
  if (currentMonth === "January") {
    Logger.log("Current month is January. No action needed.");
    return; // Exit the function without performing any actions
  }

  // Define a dictionary to map month names to their numerical values
  var months = {
    "February": 2,
    "March": 3,
    "April": 4,
    "May": 5,
    "June": 6,
    "July": 7,
    "August": 8,
    "September": 9,
    "October": 10,
    "November": 11,
    "December": 12
  };

  // Get the numerical value of the current month
  var currentMonthValue = months[currentMonth];

  // Handle the case of December wrapping around to January
  var previousMonthValue = currentMonthValue - 1;

  // Adjust for January (0) if necessary
  if (previousMonthValue === 0) {
    previousMonthValue = 12;
  }

  // Find the name of the previous month
  var previousMonth = Object.keys(months).find(key => months[key] === previousMonthValue);

  // Access the parent folder of the active spreadsheet
  var parentFolder = DriveApp.getFileById(activeSpreadsheet.getId()).getParents().next();

  // Access the previous month's spreadsheet (without ".gsheet" extension)
  var previousSpreadsheet = parentFolder.getFilesByName(previousMonth).next();

  // Access the "Inventory" sheet in the previous month's spreadsheet
  var previousSheet = SpreadsheetApp.open(previousSpreadsheet).getSheetByName("Inventory");

  // Get the data from the previous month's "Remaining Inventory" column
  var data = previousSheet.getRange("F2:F14").getValues();

  // Access the active sheet
  var activeSheet = activeSpreadsheet.getSheetByName("Inventory");

  // Set the "Starting Inventory" column in the active sheet
  activeSheet.getRange("C2:C14").setValues(data);
}
