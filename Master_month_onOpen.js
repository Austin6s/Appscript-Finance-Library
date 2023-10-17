function onOpen() {
  createCustomMenu();
}

function createCustomMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Import Data', 'importData') // Add this line to include the importData function
    .addItem('Reset Data', 'clearDataWithRowAndColumnPreservation')
    .addToUi();
}