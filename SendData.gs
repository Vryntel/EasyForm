function sendData() {

  const settingsSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = settingsSpreadsheet.getSheetByName("Settings");
  const formSheet = settingsSpreadsheet.getSheetByName("Form Questions");
  const sourceSpreadsheetName = settingsSheet.getRange("B3").getValue();
  const sourceWorksheetName = settingsSheet.getRange("B4").getValue();
  const formURL = settingsSheet.getRange("B9").getValue();
  var stopOnError = settingsSheet.getRange("B13").getValue();
  var errors = "";

  var dataSheet;
  var headers;
  var formFields;
  var sectionsCounter;

  // Retrive stored info about the form
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const savedFormURL = scriptProperties.getProperty('formURL');
    if (savedFormURL != formURL) {
      SpreadsheetApp.getUi().alert("Form URL not matching the fields. Try to analyze the form again");
      return;
    }
    formFields = JSON.parse(scriptProperties.getProperty('formFields'));
    headers = JSON.parse(scriptProperties.getProperty('headers'));
    sectionsCounter = scriptProperties.getProperty('sectionsCounter');

  } catch (err) {
    // TODO (developer) - Handle exception
    console.log('Failed with error %s', err.message);
  }

  // Check if parameters are empty
  if (sourceSpreadsheetName == "" || sourceWorksheetName == "" || formURL == "") {
    SpreadsheetApp.getUi().alert("Insert requested values: Spreadsheet url, Worksheet name, Form url");
    return;
  }


  // Check if spreadsheet url is the same spreadsheet
  if (sourceSpreadsheetName.split("d/")[1].split("/")[0] != SpreadsheetApp.getActiveSpreadsheet().getId()) {
    try {
      dataSheet = SpreadsheetApp.openByUrl(sourceSpreadsheetName).getSheetByName(sourceWorksheetName);
      if (dataSheet == null) {
        SpreadsheetApp.getUi().alert("Worksheet not valid");
        return;
      }
    }
    catch (e) {
      SpreadsheetApp.getUi().alert("Spreadsheet url not valid");
      return;
    }
  }
  else {
    dataSheet = settingsSpreadsheet.getSheetByName(sourceWorksheetName);
  }





  var formUrlSubmit;
  var pageHistory = "?&pageHistory=0";

  var splittedURL = formURL.split('/');
  var formId = splittedURL[splittedURL.length - 2];

  if (sectionsCounter == 1) {
    formUrlSubmit = "https://docs.google.com/forms/u/0/d/e/" + formId + "/formResponse";
  }
  else {
    for (let j = 1; j < sectionsCounter; j++) {
      pageHistory = pageHistory + "," + j;
    }
    formUrlSubmit = "https://docs.google.com/forms/u/0/d/e/" + formId + "/formResponse" + pageHistory;
  }


  var startRow = settingsSheet.getRange("B6").getValue();
  var endRow = settingsSheet.getRange("B7").getValue();

  var lastRow = formSheet.getLastRow();
  var valuesToUse = formSheet.getRange("B2:E" + lastRow).getValues();

  // Contains the column index / custom value to use as input to the form
  var dataToSend = [];

  var header;

  // Get the column index / custom value to use as input to the form
  // if element of array is number it's a column reference on the source sheet, otherwise if string it is custom value
  for (let i = 0; i < valuesToUse.length; i++) {

    // If column and custom values are empty, we need to check if question is required by using stored array infos formFields
    if (valuesToUse[i][0] == "" && valuesToUse[i][3] == "") {
      if (formFields[i][3] == true) {
        SpreadsheetApp.getUi().alert("Value is required for question: \n" + formFields[i][1] + " (Row " + (i + 2) + ")");
        return;
      }
      else {
        dataToSend.push("");
      }
    }
    // Check if Custom value checkbox is checked
    else if (valuesToUse[i][2] == true) {
      dataToSend.push(valuesToUse[i][3].toString());
    }
    else {
      // Find the dropdown column index in the source sheet
      header = headers.indexOf(valuesToUse[i][0]);
      if (header == -1) {
        SpreadsheetApp.getUi().alert("Column name not found");
        return;
      }
      dataToSend.push(header);
    }

  }

  // String that contains all values to send
  // I don't use an object because for checkbox you need multiple same key
  var payload = "";

  // Source data
  var rowsToSubmit = dataSheet.getRange(startRow, 1, (endRow - startRow) + 1, dataSheet.getLastColumn()).getValues();

  var questionType;
  var entryid;
  var answer;
  var options;


  // Itereate over the rows to submit
  for (let x = 0; x < rowsToSubmit.length; x++) {

    payload = "";

    for (let i = 0; i < formFields.length; i++) {

      questionType = formFields[i][0];
      entryid = "entry." + formFields[i][2];

      // If number it is a column index of the source sheet
      if (typeof dataToSend[i] == "number") {
        answer = rowsToSubmit[x][dataToSend[i]];
      }
      else {
        answer = dataToSend[i];
      }

      // If column index to use or custom value is empty skip iteration
      if (answer == "") {
        continue;
      }

      // If checkbox question I need to repeat the same entry id for all selected checkboxes
      if (questionType == 4) {
        if (answer.search("\u2009") != -1) {
          answer = answer.split("\u2009");
          for (let j = 0; j < answer.length; j++) {
            payload = payload + "&" + entryid + "=" + answer[j];
          }
        }
        else {
          payload = payload + "&" + entryid + "=" + answer;
        }
      }
      else {
        // If Date/Time question convert to date (for date I added also hour and minutes 
        // because in the form you can add the time to a date question)
        if (questionType == 9) {
          answer = Utilities.formatDate(new Date(answer), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd hh:mm');
        }

        if (questionType == 10) {
          answer = Utilities.formatDate(new Date(answer), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'hh:mm:ss');
        }

        payload = payload + "&" + entryid + "=" + answer;
      }
    }

    options = {
      "method": "post",
      "payload": payload
    };

    try {
      UrlFetchApp.fetch(formUrlSubmit, options);
    }
    catch (e) {
      if (stopOnError == "TRUE") {
        SpreadsheetApp.getUi().alert("There was an error in row " + (startRow + x) + " in sheet  " + sourceWorksheetName);
        return;
      }
      else {
        errors = errors + " Row " + (startRow + x);
      }
    }
  }

  if (errors != "") {
    SpreadsheetApp.getUi().alert("There were some errors in rows: \n" + errors);
    return;
  }
}
