# student-database
A simple google sheets database access script. This README includes instructions on how to copy a simple Google Sheets UI for student database access as well as where to paste the code that allows searching/modifying the database from the UI.


**To Use:**

1 - Make a copy of the UI created in google sheets here -> https://docs.google.com/spreadsheets/d/1ujfHrzamwVwahlufm-h-W_h9NE-Y3upjemnVn_8WWvE/edit?usp=sharing. There will be three sheets, the first of which is the UI. The second sheet is the settings and the third is the database.

2 - Copy the code in the Code.gs file. Go to the google  and paste it under Extensions->Apps Script in the Code.gs file. The functions in this file must keep their names because they are synced up with the buttons in the UI.


**Documentation:**

Form Worksheet: This worksheet is how the database is searched and edited. 

Database #: This field will be ignored for searches, since the database number of the student matches of the data being searched for is not known at the time of a search.

Search Button: Searches the database for the information inputted into the input fields. If multiple matches are found, there will be a prompt to view them in the sidebar. When inputting information to search, follow **warning 1** (see bottom of this README).

Save Button: Saves a new student in the database matching the input information in the database. The Database # on the Form Worksheet and in the Settings Worksheet will automatically update to the correct number so the student gets saved as the most recent entry. To avoid unexpected vehavior, see **warning 2**

Delete Button: Searches the database for the information inputted into the input fields and deletes that entry in the database. This function also updates the database number in the Settings Worksheet and the Form Worksheet.

**Resetting the Database:** Resetting the database will clear all student entries, reset all settings to their default values, and populate the database columns with the student properties listed as labels to the input fields in the Form Worksheet. This function must be run from the Apps Script Extension. To get there, go to Extensions -> Apps Script, select the resetDatabase function, and click the play icon to run the function.

**Warning 1:** When searching, saving, and deleting, make sure you have updated all cells with input information. Relevant functions will identify cell information based on what is stored. If you modify a cell and then forget to press enter or deselect the cell, the information stored there will not be updated, meaning you could get unexpected behavior, i.e. if you change the student name you are searching for, but forget to press enter or deselect the name input cell, it will search for whatever was in that cell before you inputted the new name since the information was not updated in the spreadsheet.

**Warning 2:** The current version DOES NOT check for duplicates or support updating student information for a given student entry, so expect a new entry every time the Save Button is pressed and expect only one student match to be deleted when the Delete Button is pressed.
