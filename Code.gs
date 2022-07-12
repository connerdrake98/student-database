/******************************************************************/
/*****************GLOBAL VARIABLES*********************************/
/******************************************************************/
const ss = SpreadsheetApp.getActiveSpreadsheet();

const formWS = ss.getSheetByName("Form");
const formWSSelector = formWS.getDataRange().getValues();
const settingsWS = ss.getSheetByName("Settings");
const dataWS = ss.getSheetByName("Data");
const dataWSSelector = dataWS.getDataRange().getValues();
const idCell = formWS.getRange("C2");
const idValue = formWS.getRange("C2").getValue();

const databasePropertyID = "databaseNumber";

// // list of input field names and their input cells
// // Update these values if you ever change the format of the Form document!
// // Then to use the new input field cells, run the computeInputFields function so
// // the new input fields can be stored for use in searching/modifying records.
// For dependency reasons, don't change the names of the properties of this object.
const inputFieldsObject = {
  [databasePropertyID]: { idCell: "B2", inputCell: "C2" },
  // databaseNumber does not need a "neededForValidEntry" property because if the
  // user forgets to set the database number when saving the record, it will automatically
  // be calculated.
  familyStudentID: { idCell: "B4", inputCell: "C4", neededForValidEntry: true },
  lastName: { idCell: "B6", inputCell: "C6", neededForValidEntry: true },
  firstName: { idCell: "B8", inputCell: "C8", neededForValidEntry: true },
  middleName: { idCell: "B10", inputCell: "C10", neededForValidEntry: false },
  dateOfBirth: { idCell: "B12", inputCell: "C12", neededForValidEntry: false },
  grade: { idCell: "B14", inputCell: "C14", neededForValidEntry: false },
  teacher: { idCell: "E6", inputCell: "F6", neededForValidEntry: false },
  team: { idCell: "E8", inputCell: "F8", neededForValidEntry: false },
  transportation: {
    idCell: "E10",
    inputCell: "F10",
    neededForValidEntry: false,
  },
  parentEmail: { idCell: "E12", inputCell: "F12", neededForValidEntry: false },
  parentPhone: { idCell: "E14", inputCell: "F14", neededForValidEntry: false },
  contact1: { idCell: "B18", inputCell: "C18", neededForValidEntry: false },
  contact1Phone: {
    idCell: "B19",
    inputCell: "C19",
    neededForValidEntry: false,
  },
  contact2: { idCell: "E18", inputCell: "F18", neededForValidEntry: false },
  contact2Phone: {
    idCell: "E19",
    inputCell: "F19",
    neededForValidEntry: false,
  },
  contact3: { idCell: "B21", inputCell: "C21", neededForValidEntry: false },
  contact3Phone: {
    idCell: "B22",
    inputCell: "C22",
    neededForValidEntry: false,
  },
  healthNotes: {
    idCell: "B24",
    inputCells: ["B25", "B26"],
    neededForValidEntry: false,
  },
  iepNotes: {
    idCell: "B28",
    inputCells: ["B29", "B30"],
    neededForValidEntry: false,
  },
  generalNotes: {
    idCell: "B32",
    inputCells: ["B33", "B34"],
    neededForValidEntry: false,
  },
};
const databaseOriginCell = "A1";

const errorMessageSettings = {
  clearWarning: { inputCell: "E2", defVal: true },
  displayedContinueClearOption: { inputCell: "E3", defVal: false },
};

// These arrays should ALWAYS be left blank. They are calculated as needed by the functions below.
const inputFieldsList = [];
const requiredCellsForValidEntry = [];

/*************************************************************************/
/*****************END OF GLOBAL VARIABLES*********************************/
/*************************************************************************/

// Alerts user with a message
function alertUser(message) {
  SpreadsheetApp.getUi().alert(message);
}

/*************************************************************************/

// computes input fields from inputFieldsObject
// This function does not need to be called, it is called when needed in major functions.

const computeInputFields = function () {
  for (const { idCell, inputCell, inputCells } of Object.values(
    inputFieldsObject
  )) {
    if (idCell) {
      if (inputCell) {
        inputFieldsList.push(inputCell);
      } else if (inputCells) {
        inputCells.forEach((cellId) => inputFieldsList.push(cellId));
      } else {
        alertUser(
          "Error: property of inputFieldsList does not contain input cell(s)"
        );
      }
    }
  }
};

/*************************************************************************/

// Computes the information needed for a valid record entry in Form.WS based on
// the "neededForValidEntry" property of inputFieldsObject
// // this function does not need to be called, it is called when needed in major functions

function computeInputFieldsForValidEntry() {
  for (const { inputCell, neededForValidEntry } of Object.values(
    inputFieldsObject
  )) {
    if (neededForValidEntry === true && inputCell && inputCell != "") {
      requiredCellsForValidEntry.push(inputCell);
    }
  }
}

/*************************************************************************/

// Takes the names of the fields from the "Form" worksheet and pastes them as column names
// in the "Data" worksheet

function setDatabaseColumnNames() {
  computeInputFields();

  let i = 0;

  for (const { idCell, inputCell, inputCells } of Object.values(
    inputFieldsObject
  )) {
    if (idCell) {
      if (inputCell && !inputCells) {
        dataWS.getRange(1, i + 1).setValue(formWS.getRange(idCell).getValue());
        ++i;
      } else if (!inputCell && inputCells) {
        inputCells.forEach(function (inputCell) {
          dataWS
            .getRange(1, i + 1)
            .setValue(formWS.getRange(idCell).getValue());
          ++i;
        });
      }
    }
  }
}

/*************************************************************************/

// prompt user and display yes/no prompt

function promptUserYesNo(message) {
  if (!message) message = "undefined";

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(message, ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    return "yes";
  } else if (response == ui.Button.NO) {
    return "no";
  }
}

// Get Database Number in Settings Worksheet
// (Helper Function - calling this function on its own won't do anything)

function getDatabaseNumberSettings() {
  for (const [id, el] of Object.entries(inputFieldsObject)) {
    if (id === "databaseNumber") {
      const { idCell, inputCell } = el;
      return formWS.getRange(inputCell).getValue();
    }
  }
}

/*************************************************************************/

// Set Database Number in the Settings Worksheet

function setDatabaseNumberSettings(num) {
  let databaseRange;
  for (const [id, el] of Object.entries(inputFieldsObject)) {
    if (id === "databaseNumber") {
      const { idCell, inputCell } = el;
      formWS.getRange(inputCell).setValue(num);
    }
  }
}

/*************************************************************************/

// gets error message settings by index
function getErrorMessageSetting(settingID) {
  for (const [key, value] of Object.entries(errorMessageSettings)) {
    if (key === settingID) {
      return value.inputCell
        ? settingsWS.getRange(value.inputCell).getValue()
        : undefined;
    }
  }
  return undefined;
}

// sets error message settings by settingID
function setErrorMessageSetting(settingID, valueToSet) {
  for (const [key, value] of Object.entries(errorMessageSettings)) {
    if (key === settingID) {
      if (value.inputCell)
        settingsWS.getRange(value.inputCell).setValue(valueToSet);
    }
  }
}

/*************************************************************************/

// Wipe/Reset Database

function wipeResetDatabase() {
  // clear Data worksheet contents, but not formatting information
  dataWS.clearContents();

  // reset settings to default values
  for (const { inputCell, defVal } of Object.values(errorMessageSettings)) {
    if (defVal != undefined && inputCell != undefined) {
      settingsWS.getRange(inputCell).setValue(defVal);
    }
  }

  // reset column names
  setDatabaseColumnNames();

  // reset Database # counter in settings to 1
  settingsWS.getRange("A2").setValue(1);

  // reset Database # in Form worksheet to 1
  setDatabaseNumberSettings(1);
}

/*************************************************************************/

// Finds the row number of the next empty data row
// (Helper Function - calling this function on its own won't do anything)

function findNextEmptyDataRow() {
  let currCell = dataWS.getRange(1, 1);
  let currCellEmpty = false;
  let i = 1;
  while (!currCellEmpty) {
    if (currCell.getValue() === "") {
      currCellEmpty = true;
    } else {
      i++;
      currCell = dataWS.getRange(i, 1);
    }
  }
  return i;
}

/*************************************************************************/

// Clears all data fields in the Form worksheet
function clearDataFieldsInFormWorksheet() {
  // get user confirmation
  let clearFields;
  let continueClearWarning;

  if (getErrorMessageSetting("clearWarning") === true) {
    clearFields =
      promptUserYesNo(
        "This will clear all data fields. It won't affect values in the database. Would you like to continue?"
      ) === "yes";
    if (
      clearFields &&
      getErrorMessageSetting("displayedContinueClearOption") === false
    ) {
      continueClearWarning =
        promptUserYesNo(
          "Would you like to continue to receive this warning message when clearing for this session?"
        ) === "yes";
      setErrorMessageSetting("displayedContinueClearOption", true);
      if (continueClearWarning)
        setErrorMessageSetting("clearWarning", continueClearWarning);
    }
  } else {
    clearFields = true;
  }

  if (clearFields) {
    if (inputFieldsList.length === 0) computeInputFields();

    inputFieldsList.forEach(function (inputFieldCell) {
      formWS.getRange(inputFieldCell).setValue("");
    });
  }
}

/*************************************************************************/

// Takes all values from the input fields and saves them as a new record

function saveRecord() {
  // set selected cell in DataWS to A1 to avoid interference
  dataWS.setActiveSelection("A1");

  // set correct database number in the settings worksheet
  setDatabaseNumberSettings(findNextEmptyDataRow() - 1);

  // sync Form Worksheet database number with database number from settings
  settingsWS.getRange("A2").setValue(getDatabaseNumberSettings());

  // add data from Form worksheet as a new record in the Data worksheet
  if (inputFieldsList.length === 0) computeInputFields();
  const fieldValues = inputFieldsList.map((f) => formWS.getRange(f).getValue());
  dataWS.appendRow(fieldValues);
}

/*************************************************************************/

// keeps searching for each search input until it removes all students from the list
// of matches that don't fit all of the search parameters
function searchRecursive(
  studentMatches,
  searchInputInformation,
  searchPropertiesIndex
) {
  // only search if the search properties index is below the number of search inputs
  if (searchPropertiesIndex < searchInputInformation.length) {
    // these will be the students out of the current matches that match the next search input
    let currSearchParameterID;
    let currSearchParameterValue;

    studentMatches.forEach(function (currMatch) {
      let matchIndex = 0;

      // see if current student matches the current search property
      currSearchParameterID = searchInputInformation[searchPropertiesIndex][1];
      currSearchParameterValue =
        searchInputInformation[searchPropertiesIndex][2];
      if (currMatch[currSearchParameterID] != currSearchParameterValue) {
        // current Student is not a match, remove student from matches list
        studentMatches.splice(i, 1);
        matchIndex--;
      }
      matchIndex++;
    });
    searchRecursive(
      studentMatches,
      searchInputInformation,
      searchPropertiesIndex + 1
    );
  }
}

/*************************************************************************/

// Searches the database for any record that matches the given input inormation
function search(deletion) {
  let searchInputInformation = [];
  // format will be: [[propertyID, propertyValue], [propertyID, propertyValue], ...]

  // compute input information from inputFieldsObject
  let currInputCellValue;

  let i = 0;
  for (const [propertyID, { idCell, inputCell, inputCells }] of Object.entries(
    inputFieldsObject
  )) {
    if (inputCell) currInputCellValue = formWS.getRange(inputCell).getValue();
    if (idCell && currInputCellValue && propertyID != "databaseNumber") {
      searchInputInformation.push([i, propertyID, currInputCellValue]);
    }
    i++;
  }

  let numSearchProperties = searchInputInformation.length;

  if (numSearchProperties != 0) {
    // set correct database number in the settings worksheet
    setDatabaseNumberSettings(findNextEmptyDataRow() - 1);

    // sync Form Worksheet database number with database number from settings
    settingsWS.getRange("A2").setValue(getDatabaseNumberSettings());

    let matchesFound = false;
    let studentMatches = [];
    let currStudentIndex = 0;
    let lastStudentIndex = getDatabaseNumberSettings() - 1;
    let searchPropertiesIndex = 0;

    let currStudentPropertyValue;
    let currSearchValue;

    // create an array of student objects that are matches for the first search parameter
    while (currStudentIndex <= lastStudentIndex) {
      currStudentPropertyValue = dataWS
        .getRange(
          currStudentIndex + 2,
          searchInputInformation[searchPropertiesIndex][0] + 1
        )
        .getValue();
      currSearchValue = searchInputInformation[0][2];

      if (currStudentPropertyValue === currSearchValue) {
        matchesFound = true;

        // for each data property, add student data into object and push to studentMatches array
        let i = 0;
        let currStudentMatch = {};
        for (const [propName, { inputCell, inputCells }] of Object.entries(
          inputFieldsObject
        )) {
          if (inputCell) {
            currStudentMatch[propName] = dataWS
              .getRange(currStudentIndex + 2, i + 1)
              .getValue();
            i++;
          } else if (inputCells) {
            let k = 0;
            inputCells.forEach(function (cell) {
              currStudentMatch[propName + String(Number(k + 1))] = dataWS
                .getRange(currStudentIndex + 2, i + k + 1)
                .getValue();
              // test code:
              //alertUser(propName + String(Number(k + 1)) + ": " + currStudentMatch[propName + String(Number(k + 1))]);
              k++;
              i++;
            });
          }
        }
        studentMatches.push(currStudentMatch);
      }
      currStudentIndex++;
    }

    if (matchesFound) {
      searchRecursive(studentMatches, searchInputInformation, 0);
      if (studentMatches.length === 0) {
        alertUser("No students were found matching the given search input.");
      } else if (studentMatches.length === 1) {
        if (!deletion)
          alertUser("Match found. Press 'ok' to load into input fields.");
        if (inputFieldsList.length === 0) computeInputFields();
        let studentMatchDataRow = studentMatches[0][databasePropertyID] + 1;
        let i = 0;
        let currPropertyValue;
        for (const { idCell, inputCell, inputCells } of Object.values(
          inputFieldsObject
        )) {
          if (idCell) {
            currPropertyValue = dataWS
              .getRange(studentMatchDataRow, i + 1)
              .getValue();
            if (inputCell) {
              formWS.getRange(inputCell).setValue(currPropertyValue);
            } else if (inputCells) {
              inputCells.forEach(function (cellID) {
                formWS.getRange(cellID).setValue(currPropertyValue);
              });
            } else {
              alertUser(
                "Error: property of inputFieldsList does not contain input cell(s)"
              );
            }
          }
          i++;
        }
        return studentMatchDataRow;
      } else {
        const showMatches =
          promptUserYesNo(
            "Multiple Matches found. Would you like to see them? Results may take a few seconds to load."
          ) === "yes";
        if (showMatches) {
          let studentMatchIndex = 0;
          let sidebarHtml = "";
          for (const student of studentMatches) {
            sidebarHtml +=
              '<p style="text-align:center">Student Match ' +
              String(studentMatchIndex + 1) +
              "</p>";
            for (const [key, value] of Object.entries(student)) {
              sidebarHtml += "<p>" + key + ": " + value + "</p>";
            }
            studentMatchIndex++;
          }
          let htmlOutput =
            HtmlService.createHtmlOutput(sidebarHtml).setTitle(
              "Student Matches"
            );
          SpreadsheetApp.getUi().showSidebar(htmlOutput);
        }
      }
    } else {
      alertUser("No matches found.");
    }
  } else {
    alertUser("There are no search parameters.");
  }
  return -1;
}

/*************************************************************************/

// Deletes the currently loaded student
function deleteStudent() {
  // check if student is a match in the database
  warningMsg = "Are you sure you would like to delete this student?";
  noDeletionMatchMsg = "There is no student match for the current information.";
  let deleteStudent;
  let rowToDelete = search(true);
  if (rowToDelete >= 0) {
    deleteStudent = promptUserYesNo(warningMsg) === "yes";
  } else {
    alertUser(noDeletionMatchMsg);
  }
  dataWS.deleteRow(rowToDelete);
  setDatabaseNumberSettings(getDatabaseNumberSettings() - 1);
}
