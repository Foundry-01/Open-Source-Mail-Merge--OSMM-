/**
 * Creates a menu entry in the Google Sheets UI when the document is opened.
 * This function runs automatically when the spreadsheet is opened.
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('OSMM')
    .addItem('Start', 'showSidebar')
    .addToUi();
}

function doGet() {
  return HtmlService.createHtmlOutput('This is a web app.');
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('OSMM')
    .setWidth(300);
  
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSheetData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  return {
    headers: data[0],
    rows: data.slice(1).filter(row => row[0] && row[1]) // Only return rows with name and email
  };
}

function getDraftTemplates() {
  try {
    var drafts = GmailApp.getDrafts();
    // Get only the 5 most recent drafts (drafts are already sorted by date, most recent first)
    return drafts.slice(0, 5).map(function(draft) {
      var message = draft.getMessage();
      return {
        id: draft.getId(),
        subject: message.getSubject() || '(no subject)'
      };
    });
  } catch (error) {
    throw new Error('Failed to load drafts. Please make sure you have authorized the script to access Gmail.');
  }
}

function sendMailMerge(templateId, senderName) {
  var draft = GmailApp.getDraft(templateId);
  if (!draft) {
    throw new Error('Selected draft template not found');
  }
  
  var message = draft.getMessage();
  var template = {
    subject: message.getSubject() || '(no subject)',
    body: message.getBody()
  };
  
  var data = getSheetData();
  var rows = data.rows;
  var sent = 0;
  
  rows.forEach(function(row) {
    var name = row[0];
    var email = row[1];
    
    try {
      // Replace {{First Name}} with the actual name if it exists in the template
      var personalizedBody = template.body.replace(/{{First Name}}/g, name);
      
      GmailApp.sendEmail(
        email,
        template.subject,
        '',  // Plain text body (empty since we're sending HTML)
        {
          htmlBody: personalizedBody,
          name: senderName  // Use the provided sender name
        }
      );
      sent++;
    } catch (error) {
      Logger.log('Failed to send email to ' + email + ': ' + error.toString());
    }
  });
  
  return sent + ' emails sent successfully';
}

function onHomepage(e) {
  var builder = CardService.newCardBuilder();
  builder.setHeader(CardService.newCardHeader().setTitle('Mail Merge'));
  
  var section = CardService.newCardSection()
    .addWidget(CardService.newTextButton()
      .setText('Start')
      .setOnClickAction(CardService.newAction().setFunctionName('showSidebar')));
  
  builder.addSection(section);
  return builder.build();
} 