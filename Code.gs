/**
* Create Google Form from Google Sheet entries.
*
* Directions
* 1. Tools -> Create a Form. Open the Form.
* 2. This form operates under the following principals (if not this script must be modified):
*    a. You have the header row in Row 1.
*    b. Your header row has 3 columns: Name, URL, and Keywords.
*    c. Name is just a string name for each row entry.
*    d. URL is a direct link to a valid image.
*    e. Keywords contains several keywords for the pic separated by commas.
* 3. If the above is true, run updateForm().
*/

function updateForm() {
  
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
  var img = "";
  var imageItem = "";
  var textItem = "";
  var keyWordList = "";
  var pattern = "";
  var textValidation = "";
  var keyWords = "";
  
  
  //  Delete all current form questions
  deleteFormItems(form);
  
  //  Go through each row on sheet, skipping header row
  for (var row = 1; row < rangeValues.length; row++){
    
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
    
    // Deletes all triggers in the current project.
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    
    //    Create trigger to capture new form submissions
    ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(spreadsheet).onFormSubmit()
    .create();
    
  }
}

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
  var name = "";
  var keys = "";
  var keyWordList = [];
  var updatedKeyWords = [];
  
  //  Go through rows to get keywords from each record
  for (var row = 1; row < rangeValues.length; row++){
    name = "";
    keys = "";
    keyWordList.length = 0;
    name = rangeValues[row][imageNameHeader];
    console.log(name);
    keys = rangeValues[row][imageKeyWordHeader];
    console.log(keys);
    keyWordList = keys.split(",").filter(function (el) { 
      return el != null && el != '';
    });
    updatedKeyWords = e.namedValues[name].filter(function (el) { 
      return el != null && el != '';
    });
    console.log(keyWordList);
    console.log(e.namedValues[name]);
    
    //    Get new keyword(s), filter null (empty) values if any, update current list of keywords
    rangeValues[row][imageKeyWordHeader] = (keyWordList.concat(updatedKeyWords)).toString();
    console.log(rangeValues[row][imageKeyWordHeader]);
  }
  
  // Add to sheet  
  range.setValues(rangeValues);
  
  //  Update form
  try {
    updateForm();
  } catch (e) {
    console.log("Form not up to date");
    console.log(e);
  }
}