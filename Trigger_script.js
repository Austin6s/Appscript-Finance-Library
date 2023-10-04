function onEdit(e) {
    var sheet = e.source.getActiveSheet();
    var editedCell = e.range;
  
    var expensesTrackerSheet = "Expenses Tracker";
    var expenseTypeColumn = 3;
    var expenseSubTypeColumn = 4;
    var inventoryDropdownColumn = 5; // Column index of the Inventory dropdown column
  
    // CURRENCY CONVERTER Check if the edited cell is F1 and the sheet is "Summary".
    if (sheet.getName() === 'Summary' && editedCell.getA1Notation() === 'F1') {
      ProcessesforMonthlysales.setCurrencyFormatForRange();
    }
  
    // DEPENDENT DROPDOWN Check if the sheet is "Expenses Tracker" and the edited cell is in the expenseTypeColumn.
    if (sheet.getName() === expensesTrackerSheet && editedCell.getColumn() === expenseTypeColumn) {
      var selectedType = editedCell.getValue();
      //Logger.log(selectedType);
      var subTypeCell = sheet.getRange(editedCell.getRow(), expenseSubTypeColumn);
  
      if (selectedType === "") {
        subTypeCell.clearDataValidations();
        subTypeCell.setValue("");
      } else {
        var subTypeRange = ProcessesforMonthlysales.getSubTypeRange(selectedType);
        //Logger.log(subTypeRange);
        if (subTypeRange) {
          var validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(subTypeRange).build();
          subTypeCell.setDataValidation(validationRule);
        } else {
          subTypeCell.clearDataValidations();
          subTypeCell.setValue("");
        }
      }
    }
   }
  
  function onOpen(){
    ProcessesforMonthlysales.populateStartingInventory();
  }
  