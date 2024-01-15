function analyzeForm() {

  const settingsSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = settingsSpreadsheet.getSheetByName("Settings");
  const formSheet = settingsSpreadsheet.getSheetByName("Form Questions");
  const sourceSpreadsheetName = settingsSheet.getRange("B3").getValue();
  const sourceWorksheetName = settingsSheet.getRange("B4").getValue();
  const formURL = settingsSheet.getRange("B9").getValue();
  var scriptProperties = PropertiesService.getScriptProperties();

  var sectionsCounter = 1;
  var formPage;

  // Sheet that contains all the rows to submit
  var dataSheet;



  // Check if parameters are empty
  if (sourceSpreadsheetName == "" || sourceWorksheetName == "" || formURL == "") {
    SpreadsheetApp.getUi().alert("Insert requested values: Spreadsheet url, Worksheet name, Form url");
    return;
  }

  // Check if spreadsheet url is the same spreadsheet of settings
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

  // Check form url
  try {
    formPage = UrlFetchApp.fetch(formURL).getContentText();
  }
  catch (e) {
    SpreadsheetApp.getUi().alert("Form url not valid");
    return;
  }


  // Clear the previously found fields
  formSheet.getRange("A2:E" + formSheet.getLastRow() + 1).clear({ contentsOnly: true, validationsOnly: true });



  try {
    scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteAllProperties();
  } catch (err) {
    // TODO (developer) - Handle exception
    console.log('Failed with error %s', err.message);
  }



  // Get the header names of the columns in the source sheet to be used in dropdown
  const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];


  // All data of the form is contained in the FB_PUBLIC_LOAD_DATA_ variable in the form page source
  var tmp = formPage.split("FB_PUBLIC_LOAD_DATA_ = ")[1];
  var publicLoadData = JSON.parse(tmp.split(";</script>")[0]);

  // Take only the part of array we need
  var formData = publicLoadData[1][1];


  /*  
      Contains form questions necessary data
      [questionType, title, entryid, required, validations]
      validations can be choices in checkboxes / multiple choices or response validation for short answer / paragraph
  */
  var formFields = [];

  // Contains data about responses of multiple choices / checkboxes question
  var responses = [];

  // Contains the possible responses values of multiple choices / checkboxes question
  var responsesValues = [];

  // Type of question
  /*
    0  Short Answer
    1  Paragraph
    2  Multiple Choice
    3  Dropdown
    4  Checkboxes
    5  Linear Scale
    6  Panel
    7  Multiple choice grid / Checkbox grid
    8  Section
    9  Date
    10 Time
    11 Image
    12 Video
    13 File Upload
  */

  var questionType;

  for (let n = 0; n < formData.length; n++) {

    responsesValues = [];
    questionType = formData[n][3];

    if (questionType == "8") {
      sectionsCounter++;
    }
    // Skip if question is Panel or Section or Image or Video
    else if (questionType != "6" && questionType != "8" && questionType != "11" && questionType != "12") {

      // Check if Multiple choice grid or Checkbox grid
      if (questionType == "7") {
        if (formData[n][4][0][11][0] == "0") {
          questionType = "2";
        }
        else {
          questionType = "4";
        }

        // Get the columns data of the grid
        responses = formData[n][4][0][1];

        // Extract only the possible answers
        for (let i = 0; i < responses.length; i++) {
          responsesValues.push(responses[i][0]);
        }

        // Loop for each row of the grid (each row is managed like a single multple choice/checkboxes question)
        for (let x = 0; x < formData[n][4].length; x++) {
          formFields.push([questionType, formData[n][4][x][3][0], formData[n][4][x][0], formData[n][4][x][2], responsesValues]);
        }

      }
      // Multiple Choice, Dropdown, Checkboxes, Linear Scale
      else if (questionType == "2" || questionType == "3" || questionType == "4" || questionType == "5") {

        responses = formData[n][4][0][1];

        for (let i = 0; i < responses.length; i++) {
          responsesValues.push(responses[i][0]);
        }

        // Checkbox can have response validation
        if (questionType == "4") {
          formFields.push([questionType, formData[n][1], formData[n][4][0][0], formData[n][4][0][2], [responsesValues, formData[n][4][0][4]]]);
        }
        else {
          formFields.push([questionType, formData[n][1], formData[n][4][0][0], formData[n][4][0][2], responsesValues]);
        }
      }
      else if (questionType == "0" || questionType == "1") {

        // Shortanswer, paragraph

        if (formData[n][4][0][4] != null) {

          // var validationType = formData[n][4][0][4][0][0];
          var validationCondition = formData[n][4][0][4][0][1];
          var validationValue = "";

          if (validationCondition != 102 && validationCondition != 103) {
            validationValue = formData[n][4][0][4][0][2];
          }

          formFields.push([questionType, formData[n][1], formData[n][4][0][0], formData[n][4][0][2], [validationCondition, validationValue]]);
        }
        else {
          formFields.push([questionType, formData[n][1], formData[n][4][0][0], formData[n][4][0][2]]);
        }

      }
      else {
        // Date, time
        formFields.push([questionType, formData[n][1], formData[n][4][0][0], formData[n][4][0][2]]);
      }
    }
  }

  // Dropdown Data validation for the columns to use as input
  var inputColumns = SpreadsheetApp.newDataValidation().requireValueInList(headers).setAllowInvalid(false).build();



  // Display form fields in the sheet

  var startingRow = 2;
  var customValueCell;

  for (let j = 0; j < formFields.length; j++, startingRow++) {
    // Question Title
    formSheet.getRange("A" + startingRow).setValue(formFields[j][1]);
    // Add dropdown for inputColumns
    formSheet.getRange("B" + startingRow).setDataValidation(inputColumns);

    // Check if question is required
    if (formFields[j][3] == "1") {
      formSheet.getRange("C" + startingRow).setValue("X");
    }

    // Add checkbox for custom value
    formSheet.getRange("D" + startingRow).insertCheckboxes();

    // Check if there are some validations (multiple choice/checkbox/short answer/paragraph)
    if (formFields[j][4] != null) {

      customValueCell = formSheet.getRange("E" + startingRow);
      questionType = formFields[j][0];

      if (questionType == "4") {
        // if checkbox I generate all possible subsets
        customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(getAllSubsets(formFields[j][4])).setAllowInvalid(false).build());
      }
      else if (questionType == "2" || questionType == "3" || questionType == "5") {
        customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(formFields[j][4]).setAllowInvalid(false).build());
      }
      else if (questionType == "0" || questionType == "1") {

        validationCondition = formFields[j][4][0];
        validationValue = formFields[j][4][1];

        if (validationValue != "") {

          /*
            1 Number
              1 Greater than
              2 Grether than or equal to
              3 Less than
              4 Less than or equal to
              5 Equal to
              6 Not Equal to
              7 Between
              8 Not Between
              9 is number
              10 Whole number
  
            2 Text
              100 Contains
              101 Doesn't Contains
              102 Email
              103 URL
  
            4 Regular Expression
              299 Contains
              300 Doesn't Contains
              301 Matches
              302 Doesn't Matches
  
            6 Length
              202 Maximum character count
              203 Minimum character count
          */

          switch (validationCondition) {
            case 1:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThan(validationValue[0]).setAllowInvalid(false).setHelpText("Number must be greater than " + validationValue[0]).build());
              break;
            case 2:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(validationValue[0]).setAllowInvalid(false).setHelpText("Number must be greater than or equal to " + validationValue[0]).build());
              break;
            case 3:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireNumberLessThan(validationValue[0]).setAllowInvalid(false).setHelpText("Number must be less than " + validationValue[0]).build());
              break;
            case 4:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireNumberLessThanOrEqualTo(validationValue[0]).setAllowInvalid(false).setHelpText("Number must be less than or equal to " + validationValue[0]).build());
              break;
            case 5:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireNumberEqualTo(validationValue[0]).setAllowInvalid(false).setHelpText("Number must be equal to " + validationValue[0]).build());
              break;
            case 6:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireNumberNotEqualTo(validationValue[0]).setAllowInvalid(false).setHelpText("Number must be different from " + validationValue[0]).build());
              break;
            case 7:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(validationValue[0], validationValue[1]).setAllowInvalid(false).setHelpText("Number must be between " + validationValue[0] + " and " + validationValue[1]).build());
              break;
            case 8:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireNumberNotBetween(validationValue[0], validationValue[1]).setAllowInvalid(false).setAllowInvalid(false).setHelpText("Number must not be between " + validationValue[0] + " and " + validationValue[1]).build());
              break;
            case 9:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireFormulaSatisfied("=ISNUMBER(E" + (j + 2) + ")").setAllowInvalid(false).setHelpText("Value must be a number ").build());
              break;
            case 10:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireFormulaSatisfied("=MOD(E" + (j + 2) + ",1)=0").setAllowInvalid(false).setHelpText("Value must be a whole number ").build());
              break;
            case 100:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireTextContains(validationValue).setAllowInvalid(false).setHelpText("Text must contains " + validationValue[0]).build());
              break;
            case 101:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireTextDoesNotContain(validationValue).setAllowInvalid(false).setHelpText("Text must not contains " + validationValue[0]).build());
              break;
            case 102:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireTextIsEmail().setAllowInvalid(false).setHelpText("Text must be a valid email address").build());
              break;
            case 103:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireTextIsUrl().setAllowInvalid(false).setHelpText("Text must be a valid URL ").build());
              break;
            case 299:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireFormulaSatisfied("=REGEXMATCH(E" + (j + 2) + "," + validationValue + ")").setAllowInvalid(false).setHelpText("Text must contains the regular expression:" + validationValue[0]).build());
              break;
            case 300:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireFormulaSatisfied("=NOT(REGEXMATCH(E" + (j + 2) + "," + validationValue + "))").setAllowInvalid(false).setHelpText("Text must not contains the regular expression " + validationValue[0]).build());
              break;
            case 301:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireFormulaSatisfied('=REGEXMATCH(E' + (j + 2) + ',"^' + validationValue + '$")').setAllowInvalid(false).setHelpText("Text must match the regular expression: " + validationValue[0]).build());
              break;
            case 302:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireFormulaSatisfied("=NOT(REGEXMATCH(E" + (j + 2) + ",^" + validationValue + "$))").setAllowInvalid(false).setHelpText("Text must not match the regular expression: " + validationValue[0]).build());
              break;
            case 202:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireFormulaSatisfied("=LEN(E" + (j + 2) + ")<=" + validationValue).setAllowInvalid(false).setHelpText("Text must have a maximum of " + validationValue[0] + " characters").build());
              break;
            case 203:
              customValueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireFormulaSatisfied("=LEN(E" + (j + 2) + ")>=" + validationValue).setAllowInvalid(false).setHelpText("Text must have a minimum of " + validationValue[0] + " characters").build());
              break;
          }
        }
      }
    }
  }

  // Save some data to be reused when submitting the form
  try {
    scriptProperties.setProperty('formURL', formURL);
    scriptProperties.setProperty('sectionsCounter', sectionsCounter);
    scriptProperties.setProperty('formFields', JSON.stringify(formFields));
    scriptProperties.setProperty('headers', JSON.stringify(headers));
  }
  catch (err) {
    // TODO (developer) - Handle exception
    console.log('Failed with error %s', err.message);
  }
}


// Generate possible combinations for checkboxes to display in the dropdown
function getAllSubsets(array) {
  const subsets = [[]];

  for (const el of array[0]) {
    const last = subsets.length - 1;
    for (let i = 0; i <= last; i++) {
      subsets.push([...subsets[i], el]);
    }
  }


  const operators = {
    '>=': function (a, b) { return a >= b },
    '<=': function (a, b) { return a <= b },
    '=': function (a, b) { return a == b }
  };
  var operator;

  /*
    Checkbox response validation

    201 at least
    204 at most
    206 exactly

  */

  if (array[1] != null) {
    switch (array[1][0][1]) {
      case 200:
        operator = ">=";
        break;
      case 201:
        operator = "<=";
        break;
      case 204:
        operator = "="
        break;
    }

    subsets = subsets.reduce((acc, item) => {
      if (operators[operator](item.length, array[1][0][2][0])) {
        acc.push(item);
      }
      return acc;
    }, []);

  }

  // Sort array by items num
  subsets.sort(function (a, b) {
    return a.length - b.length;
  });

  // Convert in flat array
  const flatArray = subsets.reduce((acc, item) => {
    if (Array.isArray(item)) {
      // acc.push(item.toString());
      // I add this white space to recognize every single item in the checkboxes
      // Example: user select Option 1 and Option 2 and Option 3
      // so when submitting the form I need the three item values (I don't use the comma as separator because the checkbox value could contain it)
      acc.push(item.join("\u2009"));
    } else {
      acc.push(item);
    }
    return acc;
  }, []);

  return flatArray;
}
