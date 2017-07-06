# Application-Maker-based-on-Form-Notifications-by-Google

This Sheets add-on was originally a forms add-on by Google called Form Notifications. 
I wanted an add-on that would also create a document from a template and email the link, 
along with other features needed for my job as a school counselor.

This add-on was originally a script run for our school to manage the submission of applications by 
8th graders and also teachers submitting student recommendation forms. We wanted a form that 
would not just store responses in a spreadsheet, but would also create a document from a 
template, send out custom notification emails, and do other advanced functions.

This is my first add-on. I know that some of my code is messy, so I welcome critiques and help 
cleaning up. Also, I did a lot of freeCodeCamp style Read-Search-Ask (Although I'm far enough 
along that I didn't have to do any asking...).

Template Headers is a side bar that allows you to save name for your new document, that will
be created from the template you select later. Now that you have created your form, all
the fields that you put into the form will appear on this page, with whitespace and symbols
eliminated, and arrows that create tags. Once you have created your Google Doc template,
you can place these tags anywhere in the template and the data from each form submission
will be used to create a unique document that contains the information from that form submission.

Configure Email is pretty much the same page as the Form Notications Add-on, but with some added
items. You can check a box and add a link. You can also add your tagged items, that look like this: <<Name>>,
to the subject or body of the email in order to customize it. 

Now that you have created your Google Doc template, Template Picker uses the Google Picker API to
allow you to select the document you want as the template. Click select and the box disappears, but saves 
the folder ID into the Apps Script settings.

The Folder Picker does much the same thing, but chooses a folder for your new documents to go into. 
In addition, Copy to Additional Folders allows you to choose up to six additional folders, and 
specify a condition under which a copy of your document (with the form submission data) will be 
placed in another folder. This is a feature we use for teacher recommendations.

Finally, Sum Spreadsheet Values is another feature we need. We have forms with Likert Scale questions,
and we need to see a sum of those values be created on our document. The multiple select box
(which was a pain to integrate with jQuery) allows you to select all the form questions you
want to be added together to create your sum.
