/**
* Create Google Form from Google Sheet entries.
*
* Directions
* 1. Tools -> Create a Form. Open the Form.
* 2. This form operates under the following principals (if not this script must be modified):
*    a. You have the header row in Row 1.
*    b. Your header row has 4 columns: Name, URL, Done, and Keywords.
*    c. Name is just a string name for each row entry.
*    d. URL is a direct link to a valid image.
*    e. Keywords contains several keywords for the pic separated by commas.
*    f. All of your names are unique values.
* 3. If the above is true, change updateValue to however many rows you want at once and run updateForm().
*/

function updateForm() {
  
  //  Set how many rows you want to do at once
  var updateValue = 10;
  PropertiesService.getScriptProperties().setProperty('rowUpdateValue', updateValue);  
  
  //  Declare variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Sheet1");
  var formURL = spreadsheet.getFormUrl();
  var form = FormApp.openByUrl(formURL)
  .setDescription('Update your keywords from your Google Sheet: ' + spreadsheet.getUrl())
  .setConfirmationMessage('Your response has been recorded. Please give the form another minute to regenerate for updated keywords.');
  var range = sheet.getDataRange();
  var rangeValues = range.getDisplayValues();
  var imageNameHeader = rangeValues[0].indexOf("Name");
  var imageURLHeader = rangeValues[0].indexOf("URL");
  var imageKeyWordHeader = rangeValues[0].indexOf("Keywords");
  var checkboxHeader = rangeValues[0].indexOf("Done");
  var img = "";
  var imageItem = "";
  var textItem = "";
  var keyWordList = "";
  var pattern = "";
  var textValidation = "";
  var keyWords = "";
  
  //  Set how many you want to update at a time
  var valueCheck = (updateValue < rangeValues.length) ? (updateValue + 1) : rangeValues.length; 
  
  //  Delete all current form questions
  deleteFormItems(form);
  
  //  Go through each row on sheet to add as a form question, skipping header row
  for (var row = 1; row < valueCheck; row++){
    
    console.log(rangeValues[row][checkboxHeader]);
    
    //    Make sure we haven't already done this one
    if (rangeValues[row][checkboxHeader] == "TRUE"){
      
      console.log("Already did " + rangeValues[row][checkboxHeader]);
      
      if (valueCheck < rangeValues.length){
        valueCheck++;
      }
      
    } else {
      
      //    Split the keywords string into an array
      keyWords = rangeValues[row][imageKeyWordHeader].split(",").filter(function (el) { 
        return el != null && el != '';
      });
      
      // Add image
      img = UrlFetchApp.fetch(rangeValues[row][imageURLHeader]);
      imageItem = form.addImageItem()
      .setTitle(rangeValues[row][imageNameHeader])
      .setImage(img);
      
      //    Add text prompt
      textItem = form.addTextItem()
      .setHelpText("Add new keyword. (Multiple keywords? Separate by commas.) Current keywords: " + keyWords.join())
      .setTitle(rangeValues[row][imageNameHeader]);
      
      /*    NOTE: This next section was supposed to include text validation so the user couldn't 
      enter a keyword that's already been entered, but this functionality is broken. When used, the
      form cannot be used (it requests you to reload the page, but still doesn't work). Leaving it in 
      to work on in the future. */
      
      //  Build regex: https://support.google.com/a/answer/1371417?hl=en
      //    keyWordList = keyWords.join("|");
      //    pattern = "(?i)(\W|^)(" + keyWordList + ")(\W|$)";
      
      // Add a text item to a form and require it to not be from the list of keywords.
      //    textValidation = FormApp.createTextValidation()
      //    .setHelpText('Keyword already in list')    
      //    .requireTextDoesNotContainPattern(pattern)
      //    .build();
      //    textItem.setValidation(textValidation);
      
    }
  }
  
  // Deletes all triggers in the current project.
  setTriggers(spreadsheet);
}

// ********************************************************************************************************************************

/**
* Delete current form questions.
*
* @param {Object} form This is the current form attached to the spreadsheet.

*/

function deleteFormItems(form){
  
  //  Make sure we have the form
  var form = form || FormApp.openByUrl(SpreadsheetApp.getActiveSpreadsheet().getFormUrl()); 
  
  //  Collect form items
  var formItems = form.getItems();
  
  //  Loop through and delete each form item
  for (var count = 0; count < formItems.length; count++){
    form.deleteItem(formItems[count]);
  } 
}

// ********************************************************************************************************************************

/**
* Convert "Done" column to checkboxes
*
* @param {Object} sheet This is our primary sheet.
*/

function convertToCheckboxes(sheet){
  
  //  Make sure we have the sheet
  var sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
  
  //  Collect sheet items
  var range = sheet.getDataRange();
  var lastRow = sheet.getLastRow();
  var rangeValues = range.getDisplayValues();  
  var checkboxHeader = rangeValues[0].indexOf("Done");
  
  //  Get items that should be checkboxes
  var checkboxRange = sheet.getRange(2, checkboxHeader + 1, lastRow - 1, 1);
  
  //  Set checkboxes to range
  checkboxRange.insertCheckboxes();
}
// ********************************************************************************************************************************

/**
* Deletes all triggers in the current project.
*
* @param {Object} spreadsheet This is our primary spreadsheet.
*/

function setTriggers(spreadsheet){
  
  // Deletes all triggers in the current project.
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  //    Create trigger to capture new form submissions
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(spreadsheet).onFormSubmit()
  .create();
  
}

// ********************************************************************************************************************************
/**
* A trigger-driven function that updates the sheet and form after a user responds to the form.
*
* @param {Object} e The event parameter for form submission to a spreadsheet;
*     see https://developers.google.com/apps-script/understanding_events
*/

function onFormSubmit(e) {
  
  //  Declare variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Sheet1");
  var range = sheet.getDataRange();
  var rangeValues = range.getDisplayValues();
  var imageNameHeader = rangeValues[0].indexOf("Name");
  var imageURLHeader = rangeValues[0].indexOf("URL");
  var imageKeyWordHeader = rangeValues[0].indexOf("Keywords");
  var checkboxHeader = rangeValues[0].indexOf("Done");
  var name = "";
  var keys = "";
  var keyWordList = [];
  var updatedKeyWords = [];
  
  //  Get how many rows we're doing at once, already set in updateForm()
  var updateValue = PropertiesService.getScriptProperties().getProperty('rowUpdateValue');
  
  //  Set how many you want to update at a time
  var valueCheck = (updateValue < rangeValues.length) ? (updateValue + 1) : rangeValues.length;  
  
  //  Go through rows to get keywords from each record
  for (var row = 1; row < valueCheck; row++){
    name = "";
    keys = "";
    keyWordList.length = 0;
    name = rangeValues[row][imageNameHeader];
    keys = rangeValues[row][imageKeyWordHeader];
    
    // Make sure we haven't already done this one
    if (rangeValues[row][checkboxHeader] == "TRUE"){
      console.log("Already did " + name);
      
      if (valueCheck < rangeValues.length){
        valueCheck++;
      }
      
    } else {
      
      // Make sure this row has been updated in the form
      if (Object.keys(e.namedValues).indexOf(name) == -1){
        console.log(name + " not returned");
        
      } else {
        // Strip empty keywords, remove empty values
        keyWordList = keys.split(",").filter(function (el) { 
          return el != null && el != '';
        });
        updatedKeyWords = e.namedValues[name].filter(function (el) { 
          return el != null && el != '';
        });
        
        // If there is a new keyword property, update the keyword list with the new keyword       
        if (updatedKeyWords.toString() != ""){
          
          // Get new keyword(s), filter null (empty) values if any, update current list of keywords
          rangeValues[row][imageKeyWordHeader] = (keyWordList.concat(updatedKeyWords)).toString();
          
          // Add check to row on sheet
          rangeValues[row][checkboxHeader] = "TRUE";            
        } else {
          console.log("No new properties returned for " + name);
        }
      }
    }
  }
  
  // Add to sheet  
  range.setValues(rangeValues);
  
  // Convert range to checks
  convertToCheckboxes(sheet);
  
  //  Update form
  try {
    updateForm();
  } catch (e) {
    console.log("Form not up to date");
    console.log(e);
  }
}