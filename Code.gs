/**
 * Creates the add-on interface when opened in Google Sheets.
 */
function onHomepage(e) {
  var builder = CardService.newCardBuilder();
  builder.setHeader(CardService.newCardHeader()
    .setTitle('')
    .setImageStyle(CardService.ImageStyle.SQUARE));
  
  var section = CardService.newCardSection()
    .addWidget(CardService.newTextButton()
      .setText('Start OSMM')
      .setBackgroundColor("#2f974b")
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('showSidebar')))
    .addWidget(CardService.newTextParagraph()
      .setText("Send personalized emails using Gmail drafts as templates. Create dynamic emails with variables from your spreadsheet."))
    .addWidget(CardService.newTextButton()
      .setText('Privacy Policy')
      .setOnClickAction(CardService.newAction().setFunctionName('showPrivacyPolicy')));
  
  builder.addSection(section);
  return builder.build();
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  // Try possible case variations
  const possibleNames = ['Sidebar.html', 'sidebar.html'];
  let html;
  let success = false;

  for (let fileName of possibleNames) {
    try {
      html = HtmlService.createHtmlOutputFromFile(fileName)
        .setTitle(' ')
        .setWidth(300);
      success = true;
      break;
    } catch (error) {
      if (!error.message.includes('No HTML file named')) {
        throw error; // If it's a different error, throw it immediately
      }
      // Otherwise continue trying other cases
    }
  }

  if (!success) {
    throw new Error('Could not find sidebar file. Please ensure either "Sidebar.html" or "sidebar.html" exists in your project.');
  }

  SpreadsheetApp.getUi().showSidebar(html);
}

function getSheetData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  Logger.log('Raw data from sheet:');
  Logger.log(data);

  // Find email column index and all cc column indices using strict matches
  var emailColumnIndex = -1;
  var ccColumnIndices = [];
  var filteredHeaders = [];

  var headers = (data && data.length > 0) ? data[0] : [];

  headers.forEach(function(header, index) {
    var headerLower = header.toString().trim().toLowerCase();

    // Capture non-empty, non-cc headers for variable display
    if (headerLower && headerLower !== 'cc' && headerLower !== 'bcc') {
      filteredHeaders.push(header);
    }

    if (headerLower === 'cc') {
      ccColumnIndices.push(index);
      return;
    }

    // Strict matching only: exactly "email" or "email address"
    if (emailColumnIndex === -1 && (headerLower === 'email' || headerLower === 'email address')) {
      emailColumnIndex = index;
    }
  });

  // Require a valid email header to proceed
  if (emailColumnIndex === -1) {
    var headerPreview = headers.map(function(h) { return '"' + String(h) + '"'; }).join(', ');
    throw new Error('Email column not found. Add a header named exactly "Email" or "Email Address" (case-insensitive). Headers found: ' + headerPreview + '.');
  }

  // Build recipient rows only if we have a detected email column and at least one data row
  var rows = [];
  if (data && data.length > 1) {
    rows = data.slice(1).filter(function(row) {
      try {
        var emailCell = row[emailColumnIndex];
        return emailCell && String(emailCell).trim() !== '';
      } catch (e) {
        return false;
      }
    });
  }

  Logger.log('Filtered rows:');
  Logger.log(rows);

  return {
    headers: headers, // original headers
    displayHeaders: filteredHeaders,
    rows: rows,
    emailColumnIndex: emailColumnIndex,
    ccColumnIndices: ccColumnIndices
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
  var emailColumnIndex = data.emailColumnIndex;
  var ccColumnIndices = data.ccColumnIndices;
  var sent = 0;
  var errors = [];
  
  rows.forEach(function(row, rowIndex) {
    try {
      var personalizedBody = template.body;
      var personalizedSubject = template.subject;
      
      // Replace all variables in both subject and body
      headers.forEach(function(header, index) {
        var value = row[index];
        var stringValue = (value === 0 || value) ? String(value) : '';
        var regex = new RegExp('{{\\s*' + header + '\\s*}}', 'gi');
        personalizedBody = personalizedBody.replace(regex, stringValue);
        personalizedSubject = personalizedSubject.replace(regex, stringValue);
      });
      
      // Get email from the email column and clean it
      var email = String(row[emailColumnIndex]).toLowerCase().trim();
      
      // Get CC emails from all CC columns
      var ccEmails = [];
      ccColumnIndices.forEach(function(ccIndex) {
        var ccEmail = String(row[ccIndex] || '').toLowerCase().trim();
        // Basic email validation for CC
        if (ccEmail && ccEmail.includes('@') && ccEmail.slice(ccEmail.indexOf('@')).includes('.')) {
          ccEmails.push(ccEmail);
        }
      });
      
      var options = {
        htmlBody: personalizedBody,
        name: senderName || 'Mail Merge'
      };
      
      // Add CC recipients if any valid emails were found
      if (ccEmails.length > 0) {
        options.cc = ccEmails.join(',');
      }
      
      // Keep strict validation for primary email
      if (!email.match(/^[^@\s]+@[^@\s]+\.[^@\s]+$/)) {
        throw new Error('Invalid email format: ' + email);
      }
      
      // Send the email
      GmailApp.sendEmail(
        email,
        personalizedSubject,
        '',  // Plain text body (empty since we're sending HTML)
        options
      );
      sent++;
      
      // Add small delay between sends
      Utilities.sleep(500);
      
    } catch (error) {
      Logger.log('Failed to send email to ' + row[emailColumnIndex] + ': ' + error.toString());
      errors.push('Error sending to ' + row[0] + ': ' + error.toString());
    }
  });
  
  var message = sent + ' emails sent successfully';
  if (errors.length > 0) {
    message += '\n' + errors.join('\n');
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
  
  var data = getSheetData();
  var headers = data.headers.map(function(header) { return header.toLowerCase(); });
  var firstRow = data.rows[0] || [];
  
  var personalizedBody = template.body;
  var personalizedSubject = template.subject;
  
  headers.forEach(function(header, index) {
    var value = firstRow[index];
    var stringValue = (value === 0 || value) ? String(value) : '';
    var regex = new RegExp('{{\\s*' + header + '\\s*}}', 'gi');
    personalizedBody = personalizedBody.replace(regex, stringValue);
    personalizedSubject = personalizedSubject.replace(regex, stringValue);
  });
  
  senderEmail = String(senderEmail).toLowerCase().trim();
  
  try {
    // For test emails, we only send to the user without any CC
    GmailApp.sendEmail(
      senderEmail,
      '[TEST] ' + personalizedSubject,  // Add [TEST] prefix to subject
      '',  // Plain text body (empty since we're sending HTML)
      {
        htmlBody: personalizedBody,
        name: senderName || 'Mail Merge'
      }
    );
    return 'Test email sent successfully to ' + senderEmail;
  } catch (error) {
    Logger.log('Failed to send test email: ' + error.toString());
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

/**
 * Shows the privacy policy in a modal dialog
 */
function showPrivacyPolicy() {
  var html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Privacy Policy - Open Source Mail Merge</title>
        <style>
            body {
                font-family: 'Google Sans', Arial, sans-serif;
                line-height: 1.6;
                margin: 0;
                padding: 20px;
                color: #202124;
            }
            h1 {
                color: #1a73e8;
                font-size: 2em;
                margin-bottom: 1em;
            }
            h2 {
                color: #202124;
                font-size: 1.5em;
                margin-top: 1.5em;
            }
            p {
                margin-bottom: 1em;
            }
            ul {
                margin-bottom: 1em;
                padding-left: 20px;
            }
            li {
                margin-bottom: 0.5em;
            }
            .last-updated {
                color: #5f6368;
                font-style: italic;
                margin-top: 2em;
            }
            a {
                color: #1a73e8;
                text-decoration: none;
            }
            a:hover {
                text-decoration: underline;
            }
        </style>
    </head>
    <body>
        <h1>Privacy Policy for Open Source Mail Merge</h1>

        <p>This Privacy Policy describes how Open Source Mail Merge ("we", "our", or "the add-on") handles information when you use our Google Workspace Add-on.</p>

        <h2>Information We Access</h2>
        <p>Our add-on requires access to:</p>
        <ul>
            <li>Google Sheets: To read recipient information from your spreadsheet</li>
            <li>Gmail: To access your draft emails and send personalized emails</li>
            <li>Your email address: To send test emails and identify you as the sender</li>
        </ul>

        <h2>How We Use Your Information</h2>
        <p>The add-on uses your information solely to:</p>
        <ul>
            <li>Read recipient data from your active spreadsheet</li>
            <li>Access your Gmail drafts to use as email templates</li>
            <li>Send personalized emails to your recipients</li>
            <li>Send test emails to your account</li>
        </ul>

        <h2>Data Storage and Retention</h2>
        <p>Open Source Mail Merge:</p>
        <ul>
            <li>Does not store any of your data outside of your Google Workspace account</li>
            <li>Does not collect or retain any personal information</li>
            <li>Does not share or transmit your data to any third parties</li>
            <li>Only processes data during active use of the add-on</li>
        </ul>

        <h2>Data Security</h2>
        <p>We prioritize the security of your data by:</p>
        <ul>
            <li>Operating entirely within Google's secure infrastructure</li>
            <li>Using only official Google APIs for all operations</li>
            <li>Not storing or transmitting data to external servers</li>
            <li>Requiring explicit user authorization for all permissions</li>
        </ul>

        <h2>Your Rights</h2>
        <p>You have the right to:</p>
        <ul>
            <li>Know what data the add-on accesses</li>
            <li>Revoke the add-on's access to your Google account at any time</li>
            <li>Request information about how your data is used</li>
        </ul>

        <h2>Changes to This Policy</h2>
        <p>We may update this Privacy Policy from time to time. We will notify users of any material changes through the Google Workspace Marketplace listing.</p>

        <h2>Contact Us</h2>
        <p>If you have any questions about this Privacy Policy or our data practices, please contact us at <a href="mailto:contact@binaryheart.org">contact@binaryheart.org</a>.</p>

        <p class="last-updated">Last updated: March 2024</p>
    </body>
    </html>
  `)
    .setWidth(600)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Privacy Policy');
} 