/**
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

var NOTICE = "This email was sent using Application Maker, created by Miles Catlett."

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
// This is about "about this add on" modal dialog.
function showAbout() {
  var ui = HtmlService.createHtmlOutputFromFile('About')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(640)
      .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(ui, 'About Application Maker');
}

// I borrowed this function from the below author on github
/**
 * @desc This is an Google Apps Script for getting column number by column name
 * @author Misha M.-Kupriyanov https://plus.google.com/104512463398531242371/
 * @link https://gist.github.com/5520691
 */
function getColumnNrByName(sheet, name) {
  var range = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var values = range.getValues();
  
  for (var row in values) {
    for (var col in values[row]) {
      if (values[row][col] == name) {
        return parseInt(col);
      }
    }
  }
  
  throw 'failed to get column by name';
}

// save settings to container-bound properties apps script service
function saveSettings(settings) {
  PropertiesService.getDocumentProperties().setProperties(settings);
  adjustFormSubmitTrigger();
}

// cleans question marks and other characters out of form/spreadsheet 
// headers that mess with script execution
function replaceData(str) {
  var replacedData = str.replace(/\W/ig,'');
  return replacedData;
}

function getSettings() {
  var settings = PropertiesService.getDocumentProperties().getProperties();

  // Use a default email if the creator email hasn't been provided yet.
  if (!settings.creatorEmail) {
    settings.creatorEmail = Session.getEffectiveUser().getEmail();
  }

  // Get text field items in the form and compile a list
  //   of their titles and IDs.
  var formID = SpreadsheetApp.getActiveSpreadsheet().getFormUrl().match(/\/d\/(.{25,})\//)[1];
  var form = FormApp.openById(formID);
  var textItems = form.getItems(FormApp.ItemType.TEXT);
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
  return settings;
}

/* This is the old get settings function that gets values from the spreadsheet.
   I changed this to the form because if you delete a form item, it remains as
   a header in the spreadsheet. 
   
function getSettings() {
  var settings = PropertiesService.getDocumentProperties().getProperties();

  // Use a default email if the creator email hasn't been provided yet.
  if (!settings.creatorEmail) {
    settings.creatorEmail = Session.getEffectiveUser().getEmail();
  }
  // Old teachers script to try and merge with the one below
  // Get text field items in the form and compile a list
  //   of their titles and IDs.
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 1;  // First row of data to process
  var startCol = 1;  // First column of data to process
  var numCols = sheet.getLastColumn();  // Number of columns to process
  var numRows = 1;   // Number of rows to process
  // Fetch the range of cells 
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  // get the settings save ready
  settings.textItems = [];
  for (i in data[0]) {
    var col = data[0][i];
    settings.textItems.push({
      title: col,
    });
  } 
  return settings;
}
*/


 // Adjust the onFormSubmit trigger based on user's requests.

function adjustFormSubmitTrigger() {
  var form = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getUserTriggers(form);
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
        .forSpreadsheet(form)
        .onFormSubmit()
        .create();
  } else if (!triggerNeeded && existingTrigger) {
    ScriptApp.deleteTrigger(existingTrigger);
  }
}


// Responds to a form submission event if an onFormSubmit trigger has been
// enabled.

function respondToFormSubmit(e) {
  var settings = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

  // Check if the actions of the trigger require authorizations that have not
  // been supplied yet -- if so, warn the active user via email (if possible).
  if (authInfo.getAuthorizationStatus() ==
      ScriptApp.AuthorizationStatus.REQUIRED) {
    // Re-authorization
    sendReauthorizationRequest();
  } else {
    // These will all trigger these functions on form submit
    if (settings.getProperty('creatorNotify') == 'true') {
      sendCreatorNotification();
    }
    if (settings.getProperty('documentTemplate') != '') {
      createNewDocument();
    }
    if (settings.getProperty('folderCheck1') == 'true') {
         matchFolderByHeader1();
    }
    if (settings.getProperty('folderCheck2') == 'true') {
         matchFolderByHeader2();
    }
    if (settings.getProperty('folderCheck3') == 'true') {
         matchFolderByHeader3();
    }
    if (settings.getProperty('folderCheck4') == 'true') {
         matchFolderByHeader4();
    }
    if (settings.getProperty('folderCheck5') == 'true') {
         matchFolderByHeader5();
    }
    if (settings.getProperty('folderCheck6') == 'true') {
         matchFolderByHeader6();
    }
    if (settings.getProperty('multipleSave') == 'true') {
         sumValues();
    }

    // Be sure to respect the remaining email quota.
    if (settings.getProperty('respondentNotify1') == 'true' &&
        MailApp.getRemainingDailyQuota() > 0) {
      sendRespondentNotification1(e.response);
    }
    if (settings.getProperty('respondentNotify2') == 'true' &&
        MailApp.getRemainingDailyQuota() > 0) {
      sendRespondentNotification2(e.response);
    }
    if (settings.getProperty('respondentNotify3') == 'true' &&
        MailApp.getRemainingDailyQuota() > 0) {
      sendRespondentNotification3(e.response);
    }
    if (settings.getProperty('respondentNotify4') == 'true' &&
        MailApp.getRemainingDailyQuota() > 0) {
      sendRespondentNotification4(e.response);
    }
    if (settings.getProperty('respondentNotify5') == 'true' &&
        MailApp.getRemainingDailyQuota() > 0) {
      sendRespondentNotification5(e.response);
    }
    if (settings.getProperty('respondentNotify6') == 'true' &&
        MailApp.getRemainingDailyQuota() > 0) {
      sendRespondentNotification6(e.response);
    }
  }
}

// get the form values with headers  
function getLastFormData() {
  //get headers
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 1;  // First row of data to process
  var startCol = 1;  // First column of data to process
  var numCols = sheet.getLastColumn();  // Number of columns to process
  var numRows = 1;   // Number of rows to process
  // Fetch the range of cells 
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols)
  // Fetch values for each row in the Range.
  var dataHeaders = dataRange.getValues();
  // get values
  var startRow = sheet.getLastRow();  // First row of data to process
  var startCol = 1;  // First column of data to process
  var numCols = sheet.getLastColumn();  // Number of columns to process
  var numRows = 1;   // Number of rows to process
  // Fetch the range of cells 
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var formData = [];
  for (i in data[0]) {
    formData.push({
      header: dataHeaders[0][i],
      value: data[0][i]
    });
  }
  return formData;
}

// getNewDocumentName function takes one variable string, and
// gets the document name from settings and transforms it based on document values on form submit
function getNewDocName(str) {
  var array = getLastFormData(); // header and value, example: formData[i].header;
  
  // get the start and stop positions for each header tag and push them to an array
  var settings = PropertiesService.getDocumentProperties().getProperties();
  var leftArrows = '<<';
  var rightArrows = '>>';
  // count the arrows and put them into the array
  var count = 0;
  var posLeft = str.indexOf(leftArrows);
  var posRight = str.indexOf(rightArrows);
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
}

// SUM values function gets all the values from selected questions and adds them together
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

// When Form Gets submitted
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
  
  var copyId = DriveApp.getFileById(docTemplate)
  .makeCopy(docName, docFolder)
  .getId();
  
   for (var i = 0; i < eValues.length; i++) {
      var headerTitle = replaceData(eValues[i]['header']);
      // Open the temporary document
      var copyDoc = DocumentApp.openById(copyId);
      // Get the document’s body section
      var copyBody = copyDoc.getBody();
  
      var eValue = eValues[i]['value'];
      // Replace place holder keys,in google doc template
      copyBody.replaceText('<<'+headerTitle+'>>', eValue);
      // Replace sum place holder with key for <<SumOfQuestions>>
     if (settings.multipleSave == 'true') {  
        var sum = sumValues();
        copyBody.replaceText('<<SumOfQuestions>>', sum);
     } else { copyBody.replaceText('<<SumOfQuestions>>', ''); }
      // Save and close the document
      copyDoc.saveAndClose();
   }
  // Convert temporary document to PDF
  LINK = DriveApp.getFileById(copyId).getUrl();
}


 // Called when the user needs to reauthorize. 
function sendReauthorizationRequest() {
  var settings = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  var lastAuthEmailDate = settings.getProperty('lastAuthEmailDate');
  var today = new Date().toDateString();
  if (lastAuthEmailDate != today) {
    if (MailApp.getRemainingDailyQuota() > 0) {
      var template =
          HtmlService.createTemplateFromFile('AuthorizationEmail');
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
  var formID = SpreadsheetApp.getActiveSpreadsheet().getFormUrl().match(/\/d\/(.{25,})\//)[1];
  var form = FormApp.openById(formID);
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
      var template =
          HtmlService.createTemplateFromFile('CreatorNotification');
      template.sheet =
          DriveApp.getFileById(form.getDestinationId()).getUrl();
      template.summary = form.getSummaryUrl();
      template.responses = form.getResponses().length;
      template.title = form.getTitle();
      template.responseStep = responseStep;
      template.formUrl = form.getEditUrl();
      template.notice = NOTICE;
      var message = template.evaluate();
      MailApp.sendEmail(settings.getProperty('creatorEmail'),
          form.getTitle() + ': Form submissions detected',
          message.getContent(), {
            name: ADDON_TITLE,
            htmlBody: message.getContent()
          });
    }
  }
}
//     \\\\\\\\\\\\\\\///////////////
//// Begin Respondent Notification Form responses \\\\\
       ////////////////\\\\\\\\\\\\\\\
/**
 * Sends out respondent notification emails.
 *
 * @param {FormResponse} response FormResponse object of the event
 *      that triggered this notification
 */
function sendRespondentNotification1(response) {
  var formID = SpreadsheetApp.getActiveSpreadsheet().getFormUrl().match(/\/d\/(.{25,})\//)[1];
  var form = FormApp.openById(formID);
  var settings = PropertiesService.getDocumentProperties();
  var emailId = settings.getProperty('respondentEmailItemId1');
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = sheet.getLastRow();  // First row of data to process
  var startCol = 1;  // First column of data to process
  var numCols = sheet.getLastColumn();  // Number of columns to process
  var numRows = 1;   // Number of rows to process
  // Fetch the range of cells 
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var i = getColumnNrByName(sheet, emailId);
  var mySubject = settings.getProperty('responseSubject1');
  mySubject = getNewDocName(mySubject);
  var myBody = settings.getProperty('responseText1');
  myBody = getNewDocName(myBody);
  var respondentEmail = data[0][i];
  if (respondentEmail) {
    var template =
        HtmlService.createTemplateFromFile('RespondentNotification');
    template.paragraphs = myBody.split('\n');
    if (settings.getProperty('respondentLink1') == 'true') {
       template.link = LINK;
    }
    template.notice = NOTICE;
    var message = template.evaluate();
    MailApp.sendEmail(respondentEmail,
        mySubject,
        message.getContent(), {
          name: form.getTitle(),
            htmlBody: message.getContent()
        });
  }
}

 // Sends out respondent notification emails.

function sendRespondentNotification2(response) {
  var formID = SpreadsheetApp.getActiveSpreadsheet().getFormUrl().match(/\/d\/(.{25,})\//)[1];
  var form = FormApp.openById(formID);
  var settings = PropertiesService.getDocumentProperties();
  var emailId = settings.getProperty('respondentEmailItemId2');
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = sheet.getLastRow();  // First row of data to process
  var startCol = 1;  // First column of data to process
  var numCols = sheet.getLastColumn();  // Number of columns to process
  var numRows = 1;   // Number of rows to process
  // Fetch the range of cells 
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var i = getColumnNrByName(sheet, emailId);
  var respondentEmail = data[0][i];
  if (respondentEmail) {
    var template =
        HtmlService.createTemplateFromFile('RespondentNotification');
    template.paragraphs = settings.getProperty('responseText2').split('\n');
    if (settings.getProperty('respondentLink2') == 'true') {
       template.link = LINK;
    }
    template.notice = NOTICE;
    var message = template.evaluate();
    MailApp.sendEmail(respondentEmail,
        settings.getProperty('responseSubject2'),
        message.getContent(), {
          name: form.getTitle(),
            htmlBody: message.getContent()
        });
  }
}
// Sends out respondent notification emails.
function sendRespondentNotification3(response) {
  var formID = SpreadsheetApp.getActiveSpreadsheet().getFormUrl().match(/\/d\/(.{25,})\//)[1];
  var form = FormApp.openById(formID);
  var settings = PropertiesService.getDocumentProperties();
  var emailId = settings.getProperty('respondentEmailItemId3');
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = sheet.getLastRow();  // First row of data to process
  var startCol = 1;  // First column of data to process
  var numCols = sheet.getLastColumn();  // Number of columns to process
  var numRows = 1;   // Number of rows to process
  // Fetch the range of cells 
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var i = getColumnNrByName(sheet, emailId);
  var respondentEmail = data[0][i];
  if (respondentEmail) {
    var template =
        HtmlService.createTemplateFromFile('RespondentNotification');
    template.paragraphs = settings.getProperty('responseText3').split('\n');
    if (settings.getProperty('respondentLink3') == 'true') {
       template.link = LINK;
    }
    template.notice = NOTICE;
    var message = template.evaluate();
    MailApp.sendEmail(respondentEmail,
        settings.getProperty('responseSubject3'),
        message.getContent(), {
          name: form.getTitle(),
            htmlBody: message.getContent()
        });
  }
}
// Sends out respondent notification emails.
function sendRespondentNotification4(response) {
  var formID = SpreadsheetApp.getActiveSpreadsheet().getFormUrl().match(/\/d\/(.{25,})\//)[1];
  var form = FormApp.openById(formID);
  var settings = PropertiesService.getDocumentProperties();
  var emailId = settings.getProperty('respondentEmailItemId4');
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = sheet.getLastRow();  // First row of data to process
  var startCol = 1;  // First column of data to process
  var numCols = sheet.getLastColumn();  // Number of columns to process
  var numRows = 1;   // Number of rows to process
  // Fetch the range of cells 
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var i = getColumnNrByName(sheet, emailId);
  var respondentEmail = data[0][i];
  if (respondentEmail) {
    var template =
        HtmlService.createTemplateFromFile('RespondentNotification');
    template.paragraphs = settings.getProperty('responseText4').split('\n');
    if (settings.getProperty('respondentLink4') == 'true') {
       template.link = LINK;
    }
    template.notice = NOTICE;
    var message = template.evaluate();
    MailApp.sendEmail(respondentEmail,
        settings.getProperty('responseSubject4'),
        message.getContent(), {
          name: form.getTitle(),
            htmlBody: message.getContent()
        });
  }
}
// Sends out respondent notification emails.
function sendRespondentNotification5(response) {
  var formID = SpreadsheetApp.getActiveSpreadsheet().getFormUrl().match(/\/d\/(.{25,})\//)[1];
  var form = FormApp.openById(formID);
  var settings = PropertiesService.getDocumentProperties();
  var emailId = settings.getProperty('respondentEmailItemId5');
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = sheet.getLastRow();  // First row of data to process
  var startCol = 1;  // First column of data to process
  var numCols = sheet.getLastColumn();  // Number of columns to process
  var numRows = 1;   // Number of rows to process
  // Fetch the range of cells 
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var i = getColumnNrByName(sheet, emailId);
  var respondentEmail = data[0][i];
  if (respondentEmail) {
    var template =
        HtmlService.createTemplateFromFile('RespondentNotification');
    template.paragraphs = settings.getProperty('responseText5').split('\n');
    if (settings.getProperty('respondentLink5') == 'true') {
       template.link = LINK;
    }
    template.notice = NOTICE;
    var message = template.evaluate();
    MailApp.sendEmail(respondentEmail,
        settings.getProperty('responseSubject5'),
        message.getContent(), {
          name: form.getTitle(),
            htmlBody: message.getContent()
        });
  }
}
// Sends out respondent notification emails.
function sendRespondentNotification6(response) {
  var formID = SpreadsheetApp.getActiveSpreadsheet().getFormUrl().match(/\/d\/(.{25,})\//)[1];
  var form = FormApp.openById(formID);
  var settings = PropertiesService.getDocumentProperties();
  var emailId = settings.getProperty('respondentEmailItemId6');
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = sheet.getLastRow();  // First row of data to process
  var startCol = 1;  // First column of data to process
  var numCols = sheet.getLastColumn();  // Number of columns to process
  var numRows = 1;   // Number of rows to process
  // Fetch the range of cells 
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var i = getColumnNrByName(sheet, emailId);
  var respondentEmail = data[0][i];
  if (respondentEmail) {
    var template =
        HtmlService.createTemplateFromFile('RespondentNotification');
    template.paragraphs = settings.getProperty('responseText6').split('\n');
    if (settings.getProperty('respondentLink6') == 'true') {
       template.link = LINK;
    }
    template.notice = NOTICE;
    var message = template.evaluate();
    MailApp.sendEmail(respondentEmail,
        settings.getProperty('responseSubject6'),
        message.getContent(), {
          name: form.getTitle(),
            htmlBody: message.getContent()
        });
  }
}
// This is for the picker
function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}

// Stuff for extra folders....
function matchFolderByHeader1() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var settings = PropertiesService.getDocumentProperties().getProperties();
  var name = settings.folderHeader1;
  var comp = settings.folderValue1;
  var col = getColumnNrByName(sheet, name) + 1;
  var dataRange = sheet.getRange(sheet.getLastRow(), col, 1, 1);
  var value = dataRange.getValue();
  var folder = settings.folderLoc1;
  
  if (comp == value) {
    var docID = DocumentApp.openByUrl(LINK).getId();
    var folderDes = DriveApp.getFolderById(folder);
    var newDoc = DriveApp.getFileById(docID).makeCopy(folderDes);
  }
}
function matchFolderByHeader2() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var settings = PropertiesService.getDocumentProperties().getProperties();
  var name = settings.folderHeader2;
  var comp = settings.folderValue2;
  var col = getColumnNrByName(sheet, name) + 1;
  var dataRange = sheet.getRange(sheet.getLastRow(), col, 1, 1);
  var value = dataRange.getValue();
  var folder = settings.folderLoc2;
  
  if (comp == value) {
    var docID = DocumentApp.openByUrl(LINK).getId();
    var folderDes = DriveApp.getFolderById(folder);
    var newDoc = DriveApp.getFileById(docID).makeCopy(folderDes);
  }
}
function matchFolderByHeader3() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var settings = PropertiesService.getDocumentProperties().getProperties();
  var name = settings.folderHeader3;
  var comp = settings.folderValue3;
  var col = getColumnNrByName(sheet, name) + 1;
  var dataRange = sheet.getRange(sheet.getLastRow(), col, 1, 1);
  var value = dataRange.getValue();
  var folder = settings.folderLoc3;
  
  if (comp == value) {
    var docID = DocumentApp.openByUrl(LINK).getId();
    var folderDes = DriveApp.getFolderById(folder);
    var newDoc = DriveApp.getFileById(docID).makeCopy(folderDes);
  }
}
function matchFolderByHeader4() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var settings = PropertiesService.getDocumentProperties().getProperties();
  var name = settings.folderHeader4;
  var comp = settings.folderValue4;
  var col = getColumnNrByName(sheet, name) + 1;
  var dataRange = sheet.getRange(sheet.getLastRow(), col, 1, 1);
  var value = dataRange.getValue();
  var folder = settings.folderLoc4;
  
  if (comp == value) {
    var docID = DocumentApp.openByUrl(LINK).getId();
    var folderDes = DriveApp.getFolderById(folder);
    var newDoc = DriveApp.getFileById(docID).makeCopy(folderDes);
  }
}
function matchFolderByHeader5() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var settings = PropertiesService.getDocumentProperties().getProperties();
  var name = settings.folderHeader5;
  var comp = settings.folderValue5;
  var col = getColumnNrByName(sheet, name) + 1;
  var dataRange = sheet.getRange(sheet.getLastRow(), col, 1, 1);
  var value = dataRange.getValue();
  var folder = settings.folderLoc5;
  
  if (comp == value) {
    var docID = DocumentApp.openByUrl(LINK).getId();
    var folderDes = DriveApp.getFolderById(folder);
    var newDoc = DriveApp.getFileById(docID).makeCopy(folderDes);
  }
}
function matchFolderByHeader6() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var settings = PropertiesService.getDocumentProperties().getProperties();
  var name = settings.folderHeader6;
  var comp = settings.folderValue6;
  var col = getColumnNrByName(sheet, name) + 1;
  var dataRange = sheet.getRange(sheet.getLastRow(), col, 1, 1);
  var value = dataRange.getValue();
  var folder = settings.folderLoc6;
  
  if (comp == value) {
    var docID = DocumentApp.openByUrl(LINK).getId();
    var folderDes = DriveApp.getFolderById(folder);
    var newDoc = DriveApp.getFileById(docID).makeCopy(folderDes);
  }
}
