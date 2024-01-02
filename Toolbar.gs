function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('EasyForm')
    .addItem('Analyze Form', 'analyzeForm')
    .addItem('Send Data', 'sendData')
    .addItem('Clear Form Fields', 'clearFormFields')
    .addToUi();
}



function clearFormFields() {
  const formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Questions");
  formSheet.getRange("A2:E" + formSheet.getLastRow() + 1).clear({ contentsOnly: true, validationsOnly: true });

  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteAllProperties();
  } catch (err) {
    // TODO (developer) - Handle exception
    console.log('Failed with error %s', err.message);
  }
}
