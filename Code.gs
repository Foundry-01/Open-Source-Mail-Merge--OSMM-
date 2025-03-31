/**
 * Creates the add-on interface when opened in Google Sheets.
 */
function onHomepage(e) {
  var builder = CardService.newCardBuilder();
  builder.setHeader(CardService.newCardHeader().setTitle('Open Source Mail Merge'));
  
  var section = CardService.newCardSection()
    .addWidget(CardService.newTextButton()
      .setText('Start Mail Merge')
      .setOnClickAction(CardService.newAction().setFunctionName('showSidebar')));
  
  builder.addSection(section);
  return builder.build();
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Open Source Mail Merge')
    .setWidth(300);
  
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSheetData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  Logger.log('Raw data from sheet:');
  Logger.log(data);
  
  var rows = data.slice(1).filter(row => {
    Logger.log('Checking row: ' + JSON.stringify(row));
    return row[0] && row[1];
  });
  
  Logger.log('Filtered rows:');
  Logger.log(rows);
  
  return {
    headers: data[0],
    rows: rows
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
  var headers = data.headers.map(function(header) { return header.toLowerCase(); });
  var rows = data.rows;
  var sent = 0;
  var warnings = [];
  
  rows.forEach(function(row, rowIndex) {
    try {
      var personalizedBody = template.body;
      var personalizedSubject = template.subject;
      
      // Replace all variables in both subject and body
      headers.forEach(function(header, index) {
        // Convert value to string, but handle 0 properly
        var value = row[index];
        // Check if value is numeric (including 0) or any other type
        var stringValue = (value === 0 || value) ? String(value) : '';
        
        // Check for oversized values
        if (stringValue.length > 50000) {
          warnings.push(`Warning: Value for ${header} in row ${rowIndex + 2} exceeds 50,000 characters and may be truncated`);
          stringValue = stringValue.substring(0, 50000) + '...';
        }
        
        var regex = new RegExp('{{\\s*' + header + '\\s*}}', 'gi');
        personalizedBody = personalizedBody.replace(regex, stringValue);
        personalizedSubject = personalizedSubject.replace(regex, stringValue);
      });
      
      // Check final email size (rough estimate)
      if (personalizedBody.length > 10000000) { // 10MB limit
        warnings.push(`Warning: Email for row ${rowIndex + 2} exceeds recommended size and may not send properly`);
      }
      
      // Limit subject line to 998 characters (email standard)
      if (personalizedSubject.length > 998) {
        warnings.push(`Warning: Subject line for row ${rowIndex + 2} exceeds 998 characters and will be truncated`);
        personalizedSubject = personalizedSubject.substring(0, 998);
      }
      
      // Get email from the second column (maintaining existing behavior)
      var email = row[1];
      
      GmailApp.sendEmail(
        email,
        personalizedSubject,
        '',  // Plain text body (empty since we're sending HTML)
        {
          htmlBody: personalizedBody,
          name: senderName  // Use the provided sender name
        }
      );
      sent++;
    } catch (error) {
      Logger.log('Failed to send email to ' + row[1] + ': ' + error.toString());
      warnings.push(`Error sending to ${row[1]}: ${error.toString()}`);
    }
  });
  
  // Return success message with any warnings
  var message = sent + ' emails sent successfully';
  if (warnings.length > 0) {
    message += '\n\n' + warnings.join('\n');
  }
  return message;
}

function removeRecipient(index) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();
    
    // Add 1 to skip header row
    var rowToDelete = index + 2;
    
    // Delete the row
    sheet.deleteRow(rowToDelete);
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function sendTestEmail(templateId, senderName, senderEmail) {
  var draft = GmailApp.getDraft(templateId);
  if (!draft) {
    throw new Error('Selected draft template not found');
  }
  
  var message = draft.getMessage();
  var template = {
    subject: message.getSubject() || '(no subject)',
    body: message.getBody()
  };
  
  // Get sheet data to access headers and first row for test email
  var data = getSheetData();
  var headers = data.headers.map(function(header) { return header.toLowerCase(); });
  var firstRow = data.rows[0] || [];
  
  var personalizedBody = template.body;
  var personalizedSubject = template.subject;
  
  // Replace all variables in both subject and body using first row data
  headers.forEach(function(header, index) {
    var value = firstRow[index] || '';
    var regex = new RegExp('{{\\s*' + header + '\\s*}}', 'gi');
    personalizedBody = personalizedBody.replace(regex, value);
    personalizedSubject = personalizedSubject.replace(regex, value);
  });
  
  try {
    GmailApp.sendEmail(
      senderEmail,
      personalizedSubject,
      '',  // Plain text body (empty since we're sending HTML)
      {
        htmlBody: personalizedBody,
        name: senderName  // Use the provided sender name
      }
    );
    return 'Test email sent successfully to ' + senderEmail;
  } catch (error) {
    throw new Error('Failed to send test email: ' + error.toString());
  }
}

function getUserEmail() {
  var email = Session.getEffectiveUser().getEmail();
  if (!email) {
    throw new Error('Could not determine user email. Please make sure you are logged in.');
  }
  return email;
} 