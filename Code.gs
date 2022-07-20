/******************************************************************/
/*****************GLOBAL VARIABLES*********************************/
/******************************************************************/
  const ss = SpreadsheetApp.getActiveSpreadsheet()

  const formWS = ss.getSheetByName("Form")
  const formWSSelector = formWS.getDataRange().getValues()
  const settingsWS = ss.getSheetByName("Settings")
  const dataWS = ss.getSheetByName("Data")
  const dataWSSelector = dataWS.getDataRange().getValues()
  const idCell = formWS.getRange("C2")
  const idValue = formWS.getRange("C2").getValue()
  
  const databasePropertyID = 'databaseNumber';

  // // list of input field names and their input cells
  // // Update these values if you ever change the format of the Form document!
  // // Then to use the new input field cells, run the computeInputFields function so
  // // the new input fields can be stored for use in searching/modifying records.
  // For dependency reasons, don't change the names of the properties of this object.
const inputFieldsObject = {
  [databasePropertyID]: {idCell: "B2", inputCell: "C2"}, 
  // databaseNumber does not need a "neededForValidEntry" property because if the
  // user forgets to set the database number when saving the record, it will automatically
  // be calculated.
  familyStudentID: {idCell: "B4", inputCell: "C4", neededForValidEntry: true},
  lastName: {idCell: "B6", inputCell: "C6", neededForValidEntry: true},
  firstName: {idCell: "B8", inputCell: "C8", neededForValidEntry: true},
  middleName: {idCell: "B10", inputCell: "C10", neededForValidEntry: false},
  dateOfBirth: {idCell: "B12", inputCell: "C12", neededForValidEntry: false},
  grade: {idCell: "B14", inputCell: "C14", neededForValidEntry: false},
  teacher: {idCell: "E6", inputCell: "F6", neededForValidEntry: false},
  team: {idCell: "E8", inputCell: "F8", neededForValidEntry: false},
  transportation: {idCell: "E10", inputCell: "F10", neededForValidEntry: false},
  parentEmail: {idCell: "E12", inputCell: "F12", neededForValidEntry: false},
  parentPhone: {idCell: "E14", inputCell: "F14", neededForValidEntry: false},
  contact1: {idCell: "B18", inputCell: "C18", neededForValidEntry: false},
  contact1Phone: {idCell: "B19", inputCell: "C19", neededForValidEntry: false},
  contact2: {idCell: "E18", inputCell: "F18", neededForValidEntry: false},
  contact2Phone: {idCell: "E19", inputCell: "F19", neededForValidEntry: false},
  contact3: {idCell: "B21", inputCell: "C21", neededForValidEntry: false},
  contact3Phone: {idCell: "B22", inputCell: "C22", neededForValidEntry: false},
  healthNotes: {idCell: "B24", inputCells: ["B25", "B26"], neededForValidEntry: false},
  iepNotes: {idCell: "B28", inputCells: ["B29", "B30"], neededForValidEntry: false},
  generalNotes: {idCell: "B32", inputCells: ["B33", "B34"], neededForValidEntry: false}
}
const databaseOriginCell = "A1";

const errorMessageSettings = {
  clearWarning: { inputCell: "E2", defVal: true },
  displayedContinueClearOption: { inputCell: "E3", defVal: false }
}


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

function test(range, value) {
  formWS.getRange(range).setValue(value);
}

/*************************************************************************/

function testH1(value) {
  formWS.getRange('h1').setValue(value);
}

/*************************************************************************/

// computes input fields from inputFieldsObject
// This function does not need to be called, it is called when needed in major functions.
  
const computeInputFields = function() {
    for (const { idCell, inputCell, inputCells } of Object.values(inputFieldsObject)) {
      if (idCell) {
        if (inputCell) {
          inputFieldsList.push(inputCell);
        } else if (inputCells) {
          inputCells.forEach(cellId => inputFieldsList.push(cellId));
        } else {
          alertUser('Error: property of inputFieldsList does not contain input cell(s)');
        }
      }
    }
}

/*************************************************************************/

// Computes the information needed for a valid record entry in Form.WS based on
// the "neededForValidEntry" property of inputFieldsObject
// // this function does not need to be called, it is called when needed in major functions

function computeInputFieldsForValidEntry() {  
  for (const { inputCell, neededForValidEntry } of Object.values(inputFieldsObject)) {
    if (neededForValidEntry === true && inputCell && inputCell != '' ) { 
      requiredCellsForValidEntry.push(inputCell);
    }
  }
}

/*************************************************************************/

// Takes the names of the fields from the "Form" worksheet and pastes them as column names
// in the "Data" worksheet

function setDatabaseColumnNames () {
  
  computeInputFields();
  
  let i = 0;
  
  for (const { idCell, inputCell, inputCells } of Object.values(inputFieldsObject)) {
    if (idCell) {
      if (inputCell && !inputCells) {
       dataWS.getRange(1, i + 1).setValue(formWS.getRange(idCell).getValue())
       ++i;
      } else if (!inputCell && inputCells) {
        inputCells.forEach(function(inputCell) {
          dataWS.getRange(1, i + 1).setValue(formWS.getRange(idCell).getValue())
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
    return 'yes';
  } else if (response == ui.Button.NO) {
    return 'no';
  }
}

// Get Database Number in Settings Worksheet
// (Helper Function - calling this function on its own won't do anything)

function getDatabaseNumberSettings() {
  for (const [id, el] of Object.entries(inputFieldsObject)) {
    if (id === "databaseNumber") {
      const {idCell, inputCell} = el;
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
      const {idCell, inputCell} = el;
      formWS.getRange(inputCell).setValue(num);
    }
  }
}

/*************************************************************************/

// gets error message settings by index
function getErrorMessageSetting(settingID) {
  for (const [key, value] of Object.entries(errorMessageSettings)) {
    if (key === settingID) {
      return value.inputCell ? settingsWS.getRange(value.inputCell).getValue() : undefined;
    }
  }
  return undefined;
}

// sets error message settings by settingID
function setErrorMessageSetting(settingID, valueToSet) {
  for (const [key, value] of Object.entries(errorMessageSettings)) {
    if (key === settingID) {
      if (value.inputCell) settingsWS.getRange(value.inputCell).setValue(valueToSet);
    }
  }
}

/*************************************************************************/

function generateStudentRowNumbers() {
  // TODO A2 is hard-coded as the database number in settings. Add a global variable or object and replace with the id here.
  return Array.from({ length: settingsWS.getRange("A2").getValue() }, (_, i) => i + 2);
}

/*************************************************************************/

// Wipe/Reset Database

function wipeResetDatabase () {
  // clear Data worksheet contents, but not formatting information
  dataWS.clearContents();
  
  // reset settings to default values
  for (const {inputCell, defVal} of Object.values(errorMessageSettings)) {
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
  let currCell = dataWS.getRange(1,1);
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
  return i - 1;
}


/*************************************************************************/

// Clears all data fields in the Form worksheet
function clearDataFieldsInFormWorksheet() {
  // get user confirmation
  let clearFields;
  let continueClearWarning;
  
  if (getErrorMessageSetting('clearWarning') === true) {
   clearFields = promptUserYesNo("This will clear all data fields. It won't affect values in the database. Would you like to continue?") === 'yes';
    if (clearFields && getErrorMessageSetting('displayedContinueClearOption') === false) {
      continueClearWarning = promptUserYesNo("Would you like to continue to receive this warning message when clearing for this session?") === 'yes';
      setErrorMessageSetting('displayedContinueClearOption', true)
      if (continueClearWarning) setErrorMessageSetting('clearWarning', continueClearWarning);
    }
  } else { clearFields = true; }
  
  if (clearFields) {
    
   if (inputFieldsList.length === 0) computeInputFields();
  
    inputFieldsList.forEach(function(inputFieldCell) {
      formWS.getRange(inputFieldCell).setValue("");
    }); 
  }
  
}

/*************************************************************************/

// Takes all values from the input fields and saves them as a new record

function saveNewRecord () {
  // set selected cell in DataWS to A1 to avoid interference
  dataWS.setActiveSelection("A1");
  
  // set correct database number in the settings worksheet
  setDatabaseNumberSettings(findNextEmptyDataRow());
  
  // sync Form Worksheet database number with database number from settings
  settingsWS.getRange("A2").setValue(getDatabaseNumberSettings());
  
  // search for matches for the current student so we don't add duplicate information.
  // TODO: hardcoded 2 in as the column number for the student id in the Settings WS
  // TODO: calculate the above instead.
  let studentIsDuplicate = false;
  generateStudentRowNumbers().forEach(function(rowNum) {
    if (dataWS.getRange(rowNum, 2).getValue() === formWS.getRange(inputFieldsObject.familyStudentID.inputCell).getValue()) {
      studentIsDuplicate = true;
      alertUser('You are attempting to add a student with a duplicate ID. This is not allowed.');
    }
  });
  if (studentIsDuplicate) return false;
  
  // add data from Form worksheet as a new record in the Data worksheet
  if (inputFieldsList.length === 0) computeInputFields();
  const fieldValues = inputFieldsList.map(f => formWS.getRange(f).getValue());
  dataWS.appendRow(fieldValues);
}

/******************************************************************************/

function calculateSearchInputInformation(onlyIncludePropertiesNeededForValidEntry) {
  let searchInputInformation = [];
  let currInputCellValue;
  let i = 1; 
  
  for (const [propertyID, {idCell, inputCell, inputCells, neededForValidEntry }] of Object.entries(inputFieldsObject)) {
    if (inputCell) currInputCellValue = formWS.getRange(inputCell).getValue();
    if (idCell && currInputCellValue && propertyID != 'databaseNumber') {
      if (!onlyIncludePropertiesNeededForValidEntry) {
        searchInputInformation.push([i, propertyID, currInputCellValue]);
      } else {
       if (neededForValidEntry) searchInputInformation.push([1,1]);
      }
    } 
    i++;
  }
  
  return searchInputInformation;
}

/*************************************************************************/

/*function calculateSearchObject(searchInputInformation) {
  let currStudentIndex = 0; // put in function
  let lastStudentIndex = getDatabaseNumberSettings() - 1; // put in function
  let searchPropertiesIndex = 0; // put in function and see if still works
  let matchesFound = false;
  let studentMatches = [];
  
  let currStudentPropertyValue;
  let currSearchValue;
  
  // create an array of student objects that are matches for the first search parameter
  while (currStudentIndex <= lastStudentIndex) {
    currStudentPropertyValue = dataWS.getRange(currStudentIndex + 2, searchInputInformation[searchPropertiesIndex][0] + 1).getValue();
    currSearchValue = searchInputInformation[0][2];
    
    if (currStudentPropertyValue === currSearchValue) {
      matchesFound = true;
      
      // for each data property, add student data into object and push to studentMatches array
      let i = 0;
      let currStudentMatch = {};
      for (const [propName, { inputCell, inputCells }] of Object.entries(inputFieldsObject)) {
        if (inputCell) {
          currStudentMatch[propName] = dataWS.getRange(currStudentIndex + 2, i + 1).getValue();
          i++; 
        } else if (inputCells) {
          let k = 0;
          inputCells.forEach(function(cell) {
            currStudentMatch[propName + String(Number(k + 1))] = dataWS.getRange(currStudentIndex + 2, i + k + 1).getValue();
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
  
  return { matchesFound: matchesFound, studentMatches: studentMatches };
}*/

/*************************************************************************/

// loads a student into the Form WS based on a row index
function loadStudent(studentDataRow) {
  let i = 0;
        let currPropertyValue;
        for (const { idCell, inputCell, inputCells } of Object.values(inputFieldsObject)) {
          if (idCell) {
            currPropertyValue = dataWS.getRange(studentDataRow, i + 1).getValue()
            if (inputCell) {
              formWS.getRange(inputCell).setValue(currPropertyValue);
            } else if (inputCells) {
              inputCells.forEach(function(cellID) {
                formWS.getRange(cellID).setValue(currPropertyValue);
              });
            } else {
              alertUser('Error: property of inputFieldsList does not contain input cell(s)');
              return -1;
            }
          }
          i++;
        }
}

/*************************************************************************/
/*
// gets an object from the data WS that represents a student in this format: { propertyId: propertyValue, propertyId: propertyValue, ... }
function getStudentObjectFromDataWS(rowNum) {
  if (!rowNum) rowNum = formWS.getRange("C2").getValue();
  if (dataWS.getRange(rowNum, 1).getValue === "") {
    alertUser('The current input information doesn\'t match a student. Load a student first to add Student Info');
    return null;
  }
  
  let i = 0;
  let student = {};
  
  for (const [propName, { idCell, inputCell, inputCells }] of Object.entries(inputFieldsObject)) {
     if (idCell) {
       if (inputCell) {
         student[propName] = dataWS.getRange(rowNum, i + 1).getValue();
       }
       if (inputCells) {
         let j = 0;
         inputCells.forEach(function(cell) {
           student[propName + String(j + 1)] = dataWS.getRange(rowNum, i).getValue();
           i++
         }); 
       }
       i++;
     }
   }
  return student;
}
*/

/*************************************************************************/

// Searches the database for any record that matches the given input information
function search(searchType){
  
  let searchInputInformation = calculateSearchInputInformation(false);
  // format: [propIndex1, propID1, propValue1, propIndex2, propID2, propValue2, ...]
  // e.g. 1,familyStudentID,123,2,lastName,Drake,3,firstName,Conner,4,middleName,Read,7,teacher,Flannagan,8,team,Green,9,transportation,Honda
   
  let numSearchProperties = searchInputInformation.length;
  
  if (numSearchProperties != 0) {
    
    // set correct database number in the settings worksheet
    setDatabaseNumberSettings(findNextEmptyDataRow() - 1);
    
    // sync Form Worksheet database number with database number from settings
    settingsWS.getRange("A2").setValue(getDatabaseNumberSettings());
    
    // get an array that will represent the row numbers of every student in the database
    let studentRowNums = generateStudentRowNumbers();
    // studentRowNums now includes all row numbers for all students in the database.
    let studentMatches = [];
    
    
    studentRowNums.forEach(function(currStudentRowNum, i) {
      console.log('student ' + String(i + 1));
      studentMatches.push(currStudentRowNum);
      
      searchInputInformation.every(function(currentSearchInputEntry, j) {
        let currColumn = searchInputInformation[j][0];
        let currSearchParameterID = searchInputInformation[j][1];
        let currParameterValueInputted = searchInputInformation[j][2];
        let currStudentCurrPropertyValue = dataWS.getRange(currStudentRowNum, currColumn).getValue();
        
        if (currStudentCurrPropertyValue != currParameterValueInputted) {
          // current Student is not a match, remove student from matches list
          studentMatches.pop();
          return false;
        }
      });
    });
    
    if (studentMatches.length === 0) {
      if (!searchType) { 
        alertUser('No students were found matching the given search input.');
      }
    } else if (studentMatches.length === 1) {
      if (!searchType) {
        alertUser('Match found. Press \'ok\' to load into input fields.');
        if (inputFieldsList.length === 0) computeInputFields();
        let studentMatchDataRow = studentMatches[0];
        
        if (loadStudent(studentMatchDataRow) === -1) return -1;
        
        return studentMatchDataRow; 
      }
      return 1;
    } else {
      let showMatches = false;
      if (!searchType) {
        showMatches = promptUserYesNo('Multiple Matches found. Would you like to see them? Results may take a few seconds to load.') === 'yes';
        alertUser('Note that to load in student data, you must refine your search.');
      }
      if (showMatches) {
        let studentMatchIndex = 0;
        let sidebarHtml = "";
        for (const el of studentMatches) {
          sidebarHtml += '<p style="text-align:center">Student Match ' + String(studentMatchIndex + 1) + '</p>';
          
          let i = 0;
          for (const [propName, { idCell, inputCell, inputCells }] of Object.entries(inputFieldsObject)) {
            if (idCell) {
              if (inputCell) {
                sidebarHtml += '<p>' + propName + ': ' + dataWS.getRange(studentMatches[studentMatchIndex], i + 1).getValue() + '<p>';
              }
              if (inputCells) {
                let j = 0;
                inputCells.forEach(function(cell) {
                  sidebarHtml += '<p>' + propName + Number(j + 1) + ': ' + dataWS.getRange(studentMatches[studentMatchIndex], i + 1).getValue() + '<p>';
                  j++;
                  i++;
                }); 
              }
              i++;
            }
          }
       
          studentMatchIndex++;
        }
        let htmlOutput = HtmlService.createHtmlOutput(sidebarHtml).setTitle('Student Matches');
        SpreadsheetApp.getUi().showSidebar(htmlOutput);
      } else {
        return -2;
      }
    }
  } else {
    alertUser('There are no search parameters. Make sure to press \'enter\' after entering input information before searching.');
  }
  return -1;
} 

/*************************************************************************/

// TODO: Consider case where someone wants to change the student ID. I need to report errors saying you must delee the student and re-make another student with a different ID.
// Edit current student's info.
function syncStudentInfo() {
  // set correct database number in the settings worksheet
  setDatabaseNumberSettings(findNextEmptyDataRow() - 1);
  
  // sync Form Worksheet database number with database number from settings
  settingsWS.getRange("A2").setValue(getDatabaseNumberSettings());
  
  // check to make sure there is a student present at the database number
  let databaseNumInputCellFormWS = inputFieldsObject[databasePropertyID].inputCell
  let formWSDatabaseNum = formWS.getRange(databaseNumInputCellFormWS).getValue();
  if (formWSDatabaseNum === '') {
   alertUser('No valid student loaded. Make sure to search for and load a student before attempting to change their information. Note that to change a Family/Student ID, the student must be deleted and a new student must be created.');
   return -1; 
  }
  let dataWSFirstCellDataAtDatabaseNum = dataWS.getRange(formWSDatabaseNum + 1, 1).getValue();
  
  //check to make sure there is a match for the inputted Family/Student ID
  if (inputFieldsList.length === 0) computeInputFields();
  let currFamilyStudentID = formWS.getRange(inputFieldsObject.familyStudentID.inputCell).getValue();
  let matchingIDRowNum = -1;
  
  // TODO compute the column number of the student ID, right now just using 2 as a hard-coded value  
  generateStudentRowNumbers().every(function(rowNum) {
    if (dataWS.getRange(rowNum, 2).getValue() === currFamilyStudentID) {
      if (dataWS.getRange(rowNum, 1).getValue() === formWS.getRange(databaseNumInputCellFormWS).getValue()) {
        matchingIDRowNum = rowNum;
      }
      return false;
    }
    return true;
  });
                                       
  // DEBUG: matchingIDRowNum is -1...
  if (!dataWSFirstCellDataAtDatabaseNum || matchingIDRowNum === -1) {
    alertUser('No valid student loaded. Make sure to search for and load a student before attempting to change their information. Note that to change a Family/Student ID, the student must be deleted and a new student must be created.');
    return -2;
  }
  
  // calculate which items in the database are different in the input Fields in the Form WS and add an extra warning if any of them are going to be deleted since nothing was inputted.
  // do for each item but excluce calculations for database num and student id
  
  let inputtedChanges = [];
  // will be of the structure: [[propID, formWSInputCell, FormWSvalue, dataWSValue, dataWSColumn], ...]
  
  let i = 0;
  let currDataWSValue;
  for (const [propID, { idCell, inputCell, inputCells }] of Object.entries(inputFieldsObject)) {
     if (idCell) {
       if (inputCell) {
         currDataWSValue = dataWS.getRange(matchingIDRowNum, i + 1).getValue();
         if (formWS.getRange(inputCell).getValue() != currDataWSValue) {
           inputtedChanges.push([propID, inputCell, formWS.getRange(inputCell).getValue(), currDataWSValue, i + 1]);
         }
       }
       if (inputCells) {
         let j = 0;
         inputCells.forEach(function(cell) {
           currDataWSValue = dataWS.getRange(matchingIDRowNum, i + 1).getValue();
           if (formWS.getRange(cell).getValue() != currDataWSValue) {
            inputtedChanges.push([propID, cell, formWS.getRange(cell).getValue(), currDataWSValue, i + 1]);
           }
           j++;
           i++;
         }); 
       }
       i++;
     }
   }
  console.log(inputtedChanges);
  
  if (inputtedChanges.length === 0) {
    alertUser('The inputted information is the same as the student record information and no changes will be made.')
    return false;
  }
  
  // confirm changes
  // TODO somewhat hard-coded here, fix later with computer property IDs.
  let userChanges = 'You are about to change the information for student ';
  userChanges += formWS.getRange(inputFieldsObject.familyStudentID.inputCell).getValue() + ". ";
  userChanges += 'Here are your proposed changes: ';
  
  inputtedChanges.forEach(function(change, i) {
    userChanges += '(' + Number(i + 1) + ') - ';
    userChanges += change[0] + ': ' + change[3] + ' -> ' + change[2];
    if (i < inputtedChanges.length - 1) userChanges += ', ';
  });
  
  
  if (promptUserYesNo(userChanges) === 'yes') {
    inputtedChanges.forEach(function(change, i) {
      dataWS.getRange(matchingIDRowNum, change[4]).setValue(change[2]);
    });
  }
}


/*************************************************************************/



// Deletes the currently loaded student
function deleteStudent() {
  // check if student is a match in the database
  warningMsg = 'Are you sure you would like to delete this student?';
  noDeletionMatchMsg = 'There is no student match for the current information.';
  let deleteStudent;
  let rowToDelete = search('deletion');
  if (rowToDelete >= 0) {
    deleteStudent = promptUserYesNo(warningMsg) === 'yes'
  } else {
    alertUser(noDeletionMatchMsg);
  }
  dataWS.deleteRow(rowToDelete);
  setDatabaseNumberSettings(getDatabaseNumberSettings() - 1);
}
  
