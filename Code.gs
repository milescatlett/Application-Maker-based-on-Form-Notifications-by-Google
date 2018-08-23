/*
 * Copyright 2015 Google Inc. All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 * This work has been modified and redistributed by Miles Catlett, http://milescatlett.com
 */

// A global constant String holding the title of the add-on. This is
// used to identify the add-on in the notification emails.

var ADDON_TITLE = 'Application Maker';

// A global constant 'notice' text to include with each email notification.

var NOTICE = "This email was sent using Application Maker, created by Miles Catlett.";

// Create a global variable that will generate a link to be sent via email
var LINK;

// Add the menu items (order below is order they appear in actual menu)
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Template Headers', 'showSidebarTH')
      .addItem('Configure Email', 'showSidebar')
      .addItem('Template Picker', 'showSidebarTP')
      .addItem('Folder Picker', 'showSidebarFP')
      .addItem('Copy to Additional Folders', 'showSidebarAF')
      .addItem('Sum Spreadsheet Values', 'showSidebarSum')
      .addItem('Reverse Scale', 'reverseScale')
      .addSeparator()
      .addItem('About', 'showAbout')
      .addToUi();
}

// Runs on open and on install
function onInstall(e) {
  onOpen(e);
}

// Add sidebar items
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Configure')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Application Maker');
  SpreadsheetApp.getUi().showSidebar(ui);
}
function showSidebarTH() {
  var ui = HtmlService.createHtmlOutputFromFile('Template Headers')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Application Maker');
  SpreadsheetApp.getUi().showSidebar(ui);
}
function showSidebarSum() {
  var ui = HtmlService.createHtmlOutputFromFile('SumValues')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Application Maker');
  SpreadsheetApp.getUi().showSidebar(ui);
}
// These modal dialog boxes fit the Google picker much better
function showSidebarTP() {
  var ui = HtmlService.createHtmlOutputFromFile('Picker')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(640)
      .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Choose your document template using the picker below.');
}
function showSidebarFP() {
  var ui = HtmlService.createHtmlOutputFromFile('FolderPicker')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(640)
      .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Use the picker to choose the folder your new document will be copied to.');
}
function showSidebarAF() {
  var ui = HtmlService.createHtmlOutputFromFile('AdditionalFolders')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(640)
      .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Choose additional folders using the picker.');
}
// This is the main modal dialog
function reverseScale() {
  var ui = HtmlService.createHtmlOutputFromFile('Reverse')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(640)
      .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Set Reverse Scaling Questions');
}
// This is about "about this add on" modal dialog.
function showAbout() {
  var ui = HtmlService.createHtmlOutputFromFile('About')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(640)
      .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(ui, 'About Application Maker');
}

// Get the form associated with this spreadsheet
function getTheForm() {
  try {
    var formID = SpreadsheetApp.getActiveSpreadsheet().getFormUrl().match(/\/d\/(.{25,})\//)[1];
  } catch(e) { throw "You must have form attached to spreadsheet. Go to 'Tools' and 'Create a form' or open form and attach spreadsheet."; }
  return FormApp.openById(formID);
  Logger.log('Could not get attached form.');
}

// Save settings to container-bound properties apps script service
function saveSettings(settings) {
  PropertiesService.getDocumentProperties().setProperties(settings);
  adjustFormSubmitTrigger();
}

function getSettings() {
  var settings = PropertiesService.getDocumentProperties().getProperties();

  // Use a default email if the creator email hasn't been provided yet.
  if (!settings.creatorEmail) {
    settings.creatorEmail = Session.getEffectiveUser().getEmail();
  }

  // Get text field items in the form and compile a list
  //   of their titles and IDs.
  var form = getTheForm();
  // var textItems = form.getItems(FormApp.ItemType.TEXT);
  var textItems = form.getItems();
  settings.textItems = [];
  for (var i = 0; i < textItems.length; i++) {
    settings.textItems.push({
      title: textItems[i].getTitle(),
      id: textItems[i].getTitle()        // we used to have .getId(), which gets the id of the form item
    });
  }
  var formItems = form.getItems();
  settings.formItems = [];
  for (var i = 0; i < formItems.length; i++) {
    settings.formItems.push({
      title: formItems[i].getTitle(),
      id: formItems[i].getTitle()        // we used to have .getId(), which gets the id of the form item
    });
  }
  var dailyQuota = MailApp.getRemainingDailyQuota();
  settings.dailyQuota = [];
  settings.dailyQuota.push(dailyQuota);
  
  var reverseItems = form.getItems(FormApp.ItemType.SCALE);
  settings.reverseItems = [];
  for (var i = 0; i < reverseItems.length; i++) {
    settings.reverseItems.push({
      title: reverseItems[i].getTitle(),
      id: reverseItems[i].getId() 
    });
  }
  
  return settings;
}
// Adjust the onFormSubmit trigger based on user's requests.

function adjustFormSubmitTrigger() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet(); // Get this spreadsheet
  var triggers = ScriptApp.getUserTriggers(sheet);
  var settings = PropertiesService.getDocumentProperties();
  var triggerNeeded =
      settings.getProperty('creatorNotify') == 'true' ||
      settings.getProperty('respondentNotify1') == 'true' ||
      settings.getProperty('respondentNotify2') == 'true' ||
      settings.getProperty('respondentNotify3') == 'true' ||
      settings.getProperty('respondentNotify4') == 'true' ||
      settings.getProperty('respondentNotify5') == 'true' ||
      settings.getProperty('respondentNotify6') == 'true' ||
      // Need to add a settings here for document template
      settings.getProperty('documentTemplate') != '';
  // Create a new trigger if required; delete existing trigger
  //   if it is not needed.
  var existingTrigger = null;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getEventType() == ScriptApp.EventType.ON_FORM_SUBMIT) {
      existingTrigger = triggers[i];
      break;
    }
  }
  if (triggerNeeded && !existingTrigger) {
    var trigger = ScriptApp.newTrigger('respondToFormSubmit')
        .forSpreadsheet(sheet)
        .onFormSubmit()
        .create();
  } else if (!triggerNeeded && existingTrigger) {
    ScriptApp.deleteTrigger(existingTrigger);
  }
}

// Responds to a form submission event if an onFormSubmit trigger has been
// enabled.

function respondToFormSubmit(e) {
  var lock = LockService.getScriptLock();
    try {
        lock.waitLock(29*1000); // wait 29 seconds for others' use of the code section and lock to stop and then proceed
    } catch(e) {
        Logger.log('Could not obtain lock after 29 seconds.');
    }
  var settings = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

  // Check if the actions of the trigger require authorizations that have not
  // been supplied yet -- if so, warn the active user via email (if possible).
  if (authInfo.getAuthorizationStatus() ==
      ScriptApp.AuthorizationStatus.REQUIRED) {
        // Re-authorization
        sendReauthorizationRequest();
  } else {
    if (settings.getProperty('reverseScale') == 'true') {
      try {
        changeScaleItems();
      } catch(e) { Logger.log('Error: '+e); }
      Utilities.sleep(3*1000);
    }
    // Trigger these functions on form submit
    if (settings.getProperty('documentTemplate') != '') {
      try {
        createNewDocument();
      } catch(e) { Logger.log('Error: '+e); }
    }
    if (settings.getProperty('creatorNotify') == 'true') {
      sendCreatorNotification();
    }
    // Send the emails to respondents
    if (settings.getProperty('respondentNotify1') == 'true' ||
        settings.getProperty('respondentNotify2') == 'true' ||
        settings.getProperty('respondentNotify3') == 'true' ||
        settings.getProperty('respondentNotify4') == 'true' ||
        settings.getProperty('respondentNotify5') == 'true' ||
        settings.getProperty('respondentNotify6') == 'true' 
        && MailApp.getRemainingDailyQuota() > 6
        ){
          try { 
            sendRespondentNotification(e.response);
          } catch(e) { Logger.log('Error: '+e); }
     } else { 
            var dailyQuota = MailApp.getRemainingDailyQuota();
            Logger.log('Unable to send mail. Your current dailyQuota is '+dailyQuota+'.');
     }
    if (settings.getProperty('folderCheck1') == 'true' ||
        settings.getProperty('folderCheck2') == 'true' ||
        settings.getProperty('folderCheck3') == 'true' ||
        settings.getProperty('folderCheck4') == 'true' ||
        settings.getProperty('folderCheck5') == 'true' ||
        settings.getProperty('folderCheck6') == 'true') {
          Utilities.sleep(3*1000);
          try {
            matchFolderByHeader();
          } catch(e) { Logger.log('Error: '+e); }
    } else { 
           Logger.log('Unable to run extra folder function.'); 
    }
  }
  lock.releaseLock();
}
/**
 * I adapted this function from the below author:
 * @desc This is an Google Apps Script for getting column number by column name
 * @author Misha M.-Kupriyanov https://plus.google.com/104512463398531242371/
 * @link https://gist.github.com/5520691
 *
 *Below are helper functions: 
 */
function getColumnNrByName(name) {
  var sheet = SpreadsheetApp.getActiveSheet(); // Get this spreadsheet
  var range = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var values = range.getValues();
  
  for (var row in values) {
    for (var col in values[row]) {
      if (values[row][col] == name) {
        return parseInt(col);
      }
    }
  }
  Logger.log('Failed to get column by name. Variable name = '+name+'.');
}

// Cleans question marks and other characters out of form/spreadsheet 
// headers that mess with script execution.
function replaceData(str) {
  var replacedData = str.replace(/\W/ig,'');
  return replacedData;
  Logger.log('Failed to create usable header tag.');
}

// This gets the top row where the headers from the form questions
function getHeaders() {
  var sheet = SpreadsheetApp.getActiveSheet(); // Get this spreadsheet
  var startRow = 1;  // First row of data to process
  var startCol = 1;  // First column of data to process
  var numCols = sheet.getLastColumn();  // Number of columns to process
  var numRows = 1;   // Number of rows to process
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols); // Fetch the range of cells 
  return dataRange.getValues(); // Outputs the headers for this spreadsheet as an object
  Logger.log('Could not get usable header.'); // Alerts if data not found
}

// This gets the values from the latest form submission
function getLastValues() {
  var sheet = SpreadsheetApp.getActiveSheet(); // Get this spreadsheet
  var startRow = sheet.getLastRow();  // First row of data to process
  var startCol = 1;  // First column of data to process
  var numCols = sheet.getLastColumn();  // Number of columns to process
  var numRows = 1;   // Number of rows to process
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols); // Fetch the range of cells 
  return dataRange.getValues(); // Outputs the last row of values added to the spreadsheet as an objec
  Logger.log('Could not get usable values.'); // Alerts if data not found
}

// Function that gets the headers from this spreadsheet and puts them in an array with the last values 
// submitted to this spreadsheet.
function getLastFormData() {
  var dataHeaders = getHeaders();
  var data = getLastValues();
  var formData = [];
  for (i in data[0]) {
    formData.push({
      header: dataHeaders[0][i],
      value: data[0][i]
    });
  }
  return formData;
  Logger.log('Could not get execute getLastFormData function.'); // Alerts if data not found
}

// getNewDocumentName function takes one variable string, and
// gets the document name from settings and transforms it based on document values on form submit
function getNewDocName(str) {
  var array = getLastFormData(); // header and value, example: formData[i].header;
  
  // get the start and stop positions for each header tag and push them to an array
  var leftArrows = '<<';
  var rightArrows = '>>';
  // count the arrows and put them into the array
  var count = 0;
  var posLeft = str.indexOf(leftArrows);
  var posRight = str.indexOf(rightArrows);
  if (posLeft == -1 && posRight == -1) {
    return str;
  }
  var myArr = [];
    // push the first position as it will not be caught by the loop
    myArr.push({
        start: posLeft + 2,
        stop: posRight
    });
  // push the rest of the positions into the array
  while (posLeft !== -1 && posRight !== -1) {
    count++;
    posLeft = str.indexOf(leftArrows, posLeft + 1 );
    posRight = str.indexOf(rightArrows, posRight + 1 );
    if (posLeft !== -1 && posRight !== -1) {
      myArr.push({
        start: posLeft + 2,
        stop: posRight
      });
    }
  }
  var msgArray = [];
  for (var i = 0; i < myArr.length; i++) {
     var start = parseInt(myArr[i].start);
     var stop = parseInt(myArr[i].stop);
     var res = str.substring(start, stop);
     res = replaceData(res);
     msgArray.push({
       header: res,
       begin: start,
       end: stop
     });
  }

  for (var i = 0; i < msgArray.length; i++) {
    for (var j = 0; j < array.length; j++) {
      var headerVal = replaceData(msgArray[i].header);
      var arrayHeader = replaceData(array[j].header);
      if (headerVal == arrayHeader) {
        var newString = str.replace(leftArrows + headerVal + rightArrows, array[j].value);
      }
    }
    str = newString;
  } 
  return newString;
  Logger.log('Could not get new document name.'); // Alerts if data not found
}

// Sum values function gets all the values from selected questions and adds them together
// for each form submission.
function sumValues() {
  var settings = PropertiesService.getDocumentProperties().getProperties();
  var headers = settings.multipleSelect;
  var array = headers.split(", ");
  var sheet = SpreadsheetApp.getActiveSheet();
  var formData = getLastFormData(); // header and value, example: formData[i].header;
  var sumArray = [];
  for (var i = 0; i < array.length; i++) {
    for (var j = 0; j < formData.length; j++) {
      var header = replaceData(formData[j].header);
      var comp = replaceData(array[i]);
      if (comp == header) {
        var value = formData[j].value;
        sumArray.push(value);
      }
    }
  }
  var sum = sumArray.reduce(add, 0);
  function add(a, b) {
    return a + b;
  }
  return sum;
}
// End helper functions

/* These are the functions that 
 * that RespondToFormSubmit()
 */
// This is the function that replaces values on form submit
function changeScaleItems() {
  // Get the sheet and form
  var sheet = SpreadsheetApp.getActiveSheet();
  var form = getTheForm();
  // Get only scale items
  var items = form.getItems(FormApp.ItemType.SCALE);
  // Get settings saved on the index.html page. These are the ids of the questions
  var settings = PropertiesService.getDocumentProperties().getProperties();
  // Break the string of ids into an array and create a new array
  var reverseSelect = settings.reverseSelect.split(', ');
  var scaleItems = [];
  // Double loop through the saved ids from settings and the ids from the form
  // so they can be compared with each other
  for (var i = 0; i < items.length; i++) {
    for (var h = 0; h < settings.reverseSelect.length; h++) {
      var itemId = items[i].getId();
      if (itemId == reverseSelect[h]) {
        // Get the upper and lower bound of the scale
        var upperBound = items[i].asScaleItem().getUpperBound();
        var lowerBound = items[i].asScaleItem().getLowerBound();
        // Count from the lower bound to the upper bound and push all values to array
        var possibleScores = [];
        for (var j = lowerBound; j <= upperBound; j++) {
          possibleScores.push(j);
        }
        // Count backwards from upper bound to lower bound and push to array
        // Tried to use array.prototype.reverse, but it alters the original array
        var reverseScores = [];
        for (var j = upperBound; j >= lowerBound; j--) {
          reverseScores.push(j);
        }
        // index the id of each item, plus the array of possibleScores with reverseScores
        // to a new array/object
        scaleItems.push({
          id: itemId,
          possibleScores: possibleScores,
          reverseScores: reverseScores
        });
      }
    }
  }
  var reverseSelect = settings.reverseSelect; // Get the form id numbers for items to reverse scale
  reverseSelect = reverseSelect.split(', '); // Split them into an array
  for (var i = 0; i < reverseSelect.length; i++) {   // Loop through each saved id number for items to reverse scale
    var item = form.getItemById(reverseSelect[i]); // Get the id of scaling questions in the form 
    var name = item.getTitle(); // Get the title by the id number from the form
    var questions = form.getItems(FormApp.ItemType.SCALE);
    for (var j = 0; j < questions.length; j++) {
      var question = questions[j].getTitle();
      if (question == name) {
        var colVal = getColumnNrByName(name) + 1; // Get the column number for the column that holds the values you want to reverse
        var dataRangeVal = sheet.getRange(sheet.getLastRow(), colVal, 1, 1); // Get the range for the data in that column
        var cellVal = dataRangeVal.getCell(1, 1); // Get the cell of the column you are manipulating
        var value = cellVal.getValue(); // Get the value for the question from the form itself
        var index = scaleItems[i].possibleScores.indexOf(value); // Get the index in the possible scores array of that value
        var newVal = scaleItems[i].reverseScores[index]; // Find the reversed or new value from index in the reversed array
        var col = getColumnNrByName(name) + 1; // Get the column number for the column that holds the values you want to reverse
        var dataRange = sheet.getRange(sheet.getLastRow(), col, 1, 1); // Get the range for the data in that column
        var cell = dataRange.getCell(1, 1); // Get the cell of the column you are manipulating
        cell.setValue(newVal); // Set that cell to the new, reverse scaled value
      }
    }
  }
}
// Creates the new document from the selected template
function createNewDocument() {
  var settings = PropertiesService.getDocumentProperties().getProperties();
    // Get text field items in the form and compile a list
  //   of their titles and IDs.
   var newDocName = settings.newDocumentName;
   var docTemplate = settings.documentTemplate;
   var docName = getNewDocName(newDocName);
   var folderID = settings.folderTemplate;
   var docFolder = DriveApp.getFolderById(folderID);
   var eValues = getLastFormData(); // header and value, example: formData[i].header;

  // Get information from form and set as variables
  // Remember to specify the exact title of each column below. 
  // If column names change, remember to change them here as well.
  
  var template = DriveApp.getFileById(docTemplate);
  var file = template.makeCopy(docName, docFolder);  
  var copyId = file.getId();
  
   for (var i = 0; i < eValues.length; i++) {
      var headerTitle = replaceData(eValues[i]['header']); // Replaces tags with values from spreadsheet
      var copyDoc = DocumentApp.openById(copyId); // Open the temporary document
      var copyBody = copyDoc.getBody(); // Get the document’s body section
  
      var eValue = eValues[i]['value'];
      copyBody.replaceText('<<'+headerTitle+'>>', eValue); // Replace place holder keys, in google doc template
      
     if (settings.multipleSave == 'true') {  // Executes if sum values box is checked
        var sum = sumValues(); // Adds the values using sumValues function
        copyBody.replaceText('<<SumOfQuestions>>', sum); // Replace sum place holder with key for <<SumOfQuestions>>
     } else { copyBody.replaceText('<<SumOfQuestions>>', ''); } // If no sum or box not checked, replaces these tags with ''
      copyDoc.saveAndClose(); // Save and close the document
   }
  LINK = DriveApp.getFileById(copyId).getUrl(); // Get link to Google Doc
  settings = PropertiesService.getDocumentProperties(); // Redefine settings
  var lastFormData = getLastFormData(); // Get the form data to match the name
  for (var i = 1; i <= 6; i++ ) { // Loop through each of the six fields that keeps respondent and other emails
    var respondentEmailItemId = 'respondentEmailItemId'+i; // Prepare the respondentEmail variable
    respondentEmailItemId = settings.getProperty(respondentEmailItemId); // Define respondentEmail variable
    var respondentNotify = 'respondentNotify'+i; // Prepare the respondentNotify variable
    respondentNotify = settings.getProperty(respondentNotify); // Define respondentNotify variable
    var respondentLink = 'respondentLink'+i; // Prepare the respondentLink variable
    respondentLink = settings.getProperty(respondentLink); // Define respondentLink variable
    for (var j = 0; j < lastFormData.length; j++) { // Compare each of the 6 emails with all the headers in the spreadsheet, 6x
      if (lastFormData[j].header == respondentEmailItemId && // See that 3 conditions are met: that the header matches the email header
          respondentNotify == 'true' && // That the box saying the respondent needs to be notified is checked
          respondentLink == 'true') { // That the box saying that the link should be sent is also checked
        var emailAddress = lastFormData[j].value; // Get the email address saved in the spreadsheet if all 3 conditions above are met
        try { // This try statement keeps the addViewer command below from failing if the admin of the domain has disabled sharing
          file.addViewer(emailAddress); // This adds the user with the email address, whose field was chosen in the form, as viewer
        } catch(e) { // This catches the exception
          Logger.log('Cannot add viewer'); // And logs the error message, which appears to work in the above case after testing.
        }
      }
    }
  }
}

// Called when the user needs to reauthorize. 
function sendReauthorizationRequest() {
  var settings = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  var lastAuthEmailDate = settings.getProperty('lastAuthEmailDate');
  var today = new Date().toDateString();
  if (lastAuthEmailDate != today) {
    if (MailApp.getRemainingDailyQuota() > 0) {
      var template = HtmlService.createTemplateFromFile('AuthorizationEmail');
      template.url = authInfo.getAuthorizationUrl();
      template.notice = NOTICE;
      var message = template.evaluate();
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
          'Authorization Required',
          message.getContent(), {
            name: ADDON_TITLE,
            htmlBody: message.getContent()
          });
    }
    settings.setProperty('lastAuthEmailDate', today);
  }
}

 // Sends out creator notification email(s) if the current number
 // of form responses is an even multiple of the response step setting.

function sendCreatorNotification() {
  var form = getTheForm();
  var formTitle = form.getTitle()
  var settings = PropertiesService.getDocumentProperties();
  var responseStep = settings.getProperty('responseStep');
  responseStep = responseStep ? parseInt(responseStep) : 10;

  // If the total number of form responses is an even multiple of the
  // response step setting, send a notification email(s) to the form
  // creator(s). For example, if the response step is 10, notifications
  // will be sent when there are 10, 20, 30, etc. total form responses
  // received.
  if (form.getResponses().length % responseStep == 0) {
    var addresses = settings.getProperty('creatorEmail').split(',');
    if (MailApp.getRemainingDailyQuota() > addresses.length) {
      var template = HtmlService.createTemplateFromFile('CreatorNotification');
      template.sheet = DriveApp.getFileById(form.getDestinationId()).getUrl();
      template.link = LINK;
      template.summary = form.getSummaryUrl();
      template.responses = form.getResponses().length;
      template.title = formTitle;
      template.responseStep = responseStep;
      template.formUrl = form.getEditUrl();
      template.notice = NOTICE;
      var message = template.evaluate();
      Utilities.sleep(1000);
      MailApp.sendEmail(settings.getProperty('creatorEmail'),
          formTitle + ': Form submissions detected',
          message.getContent(), {
            name: ADDON_TITLE,
            htmlBody: message.getContent()
          });
    }
  }
}

/**
 * Sends out respondent notification emails.
 *
 * @param {FormResponse} response FormResponse object of the event
 *      that triggered this notification
 */
function sendRespondentNotification(response) {
  var form = getTheForm(); // Get the form attached to the spreadsheet
  var settings = PropertiesService.getDocumentProperties(); // Get properties in friendly format
  var sheet = SpreadsheetApp.getActiveSheet(); // Get the spreadsheet container-bound
  var data = getLastValues(); // Get the last values from the spreadsheet
  var emailData = []; // Create an array for use later, outside the loop
  for (var i = 1; i <= 6; i++) { // Loop through all 6 of the checkboxes
    var respondentNotify = 'respondentNotify'+i; // Prepare the notify checkbox value
    respondentNotify = settings.getProperty(respondentNotify); // Get that value from settings
    if (respondentNotify == 'true') { // If it's checked, then...
      var respondentEmailItemId = 'respondentEmailItemId'+i; // Perpare the email header saved by sidebar
      var responseSubject = 'responseSubject'+i; // Prepare the subject saved in sidebar
      var responseText = 'responseText'+i; // Prepare the body saved in sidebar
      var emailID = settings.getProperty(respondentEmailItemId); // Get email id from settings
      var j = getColumnNrByName(emailID); // Get the column of that email header, this is a number
      var mySubject = settings.getProperty(responseSubject); // Get the subject from settings
      var myBody = settings.getProperty(responseText); // Get the body from settings
      mySubject = getNewDocName(mySubject); // Run subject through filter to replace tags
      myBody = getNewDocName(myBody); // Run body through filter to replace tags
      var respondentEmail = data[0][j];  // Get the email from spreadsheet values
      if (respondentEmail) { // If the email exists
        var template = HtmlService.createTemplateFromFile('RespondentNotification'); // Get the email template in script window
        template.paragraphs = myBody.split('\n'); // Include paragraph splits from settings
        var respondentLink = 'respondentLink'+i; // Prepare the link
        if (settings.getProperty(respondentLink) == 'true') { // If the link needs to be inclued in email...
          template.link = LINK; // Then get it from the global variable LINK
        } else { template.link = ''; } // Other wise it will not show in email
      template.notice = NOTICE; // Include global NOTICE text
      var message = template.evaluate(); // Prepare the message
        try {
          MailApp.sendEmail(respondentEmail, // Send the email
            mySubject,
            message.getContent(), {
              name: form.getTitle(),
                htmlBody: message.getContent()
          });
        } catch(e) {
          Logger.log(e+' - Email failed to send.');
        }
      } else { Logger.log('Could not send email to '+respondentEmailItemId+'.'); } // Log the errors
    } else { Logger.log(respondentNotify+' not checked.'); }
  }
}

// Stuff for extra folders....
function matchFolderByHeader() {
  var settings = PropertiesService.getDocumentProperties(); // Get settings in friendly format
  var sheet = SpreadsheetApp.getActiveSheet(); // Get the spreadsheet container-bound
  for (var i = 1; i <= 6; i++) { // Loop through each of the 6 folder ids, if present
    var folderNum = 'folderCheck'+i; // Prepare the variable to tell if extra folder check box checked
    var folderCheck = settings.getProperty(folderNum); // Get the setting for extra folder check box
    if (folderCheck == 'true') {      // If it's checked then...
      var folderHeader = 'folderHeader'+i; // Prep the column header for the folder
      var folderValue = 'folderValue'+i; // Prep the value for the condition to see if doc copied
      var folderLoc = 'folderLoc'+i; // Prep folder id
      var name = settings.getProperty(folderHeader); // Get folder header
      var startRow = sheet.getLastRow(); // Get the last row submitted
      var startCol = getColumnNrByName(name) + 1; // Get the column for condition to check
      var numCols = 1; // One cell
      var numRows = 1; // One cell
      var dataRange = sheet.getRange(startRow, startCol, numCols, numRows); // Get the cell
      var value = dataRange.getValue(); // Get the data from that cell
      var comp = settings.getProperty(folderValue); // Get the value to compare to that data
      var folder = settings.getProperty(folderLoc); // Get the folder id, to save copy
      if (comp == value) { // If the comparison and the actual value match then...
        var docID = DocumentApp.openByUrl(LINK).getId(); // Get the folder using the global LINK variable
        var folderDes = DriveApp.getFolderById(folder); // Get the folder it will be copied to
        var newDoc = DriveApp.getFileById(docID).makeCopy(folderDes); // Make a copy of file in that folder
      } else { Logger.log('Could not match extra folder.'); }
    } else { Logger.log('Extra folder not selected.'); }
  }
}

// This is for the picker
function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}

function showModalDialog() { // For testing the add on...
// Do some stuff
  var form = getTheForm();
  var html = form.getTitle(); // Put your variable you want to show on the modal so you can see if it works
// Display a modal dialog box with custom HtmlService content.
 var htmlOutput = HtmlService
     .createHtmlOutput(html)
     .setWidth(250)
     .setHeight(300);
 SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Troubleshooter');
}
