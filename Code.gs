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
 * Utilities: header normalization, email parsing, safe templating, batching
 */
function normalizeHeaderKey(key) {
  return String(key || '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function isEmailHeader(normalizedKey) {
  // Accept listed aliases for "email"
  var aliases = {
    email: true,
    emailaddress: true,
    email_address: true, // kept for clarity; normalize removes non-alnum
    email_address_alias: true // placeholder, not used
  };
  // Because we strip non-alnum, all dashes/underscores collapse
  return !!({
    email: true,
    emailaddress: true,
    emailaddressalias: true // no effect, safety
  }[normalizedKey] || aliases[normalizedKey]);
}

function isCcHeader(normalizedKey) {
  return normalizedKey === 'cc';
}

function isBccHeader(normalizedKey) {
  return normalizedKey === 'bcc';
}

function escapeRegex(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function extractEmailAddress(token) {
  var s = String(token || '').trim();
  // Prefer address inside angle brackets if present
  var angleMatch = s.match(/<([^>]+)>/);
  var addr = angleMatch ? angleMatch[1] : s;
  addr = addr.trim().toLowerCase();
  return addr;
}

function isLikelyValidEmail(addr) {
  // Moderate validation suitable for Gmail; avoids over-rejection
  return /^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/.test(String(addr || ''));
}

function parseAddressList(cellValue) {
  var raw = String(cellValue || '');
  if (!raw) return [];
  var parts = raw.split(/[;,]/).map(function(p) { return p.trim(); }).filter(function(p) { return p; });
  var seen = {};
  var result = [];
  parts.forEach(function(part) {
    var addr = extractEmailAddress(part);
    if (isLikelyValidEmail(addr) && !seen[addr]) {
      seen[addr] = true;
      result.push(addr);
    }
  });
  return result;
}

function parsePrimaryAddress(cellValue) {
  var list = parseAddressList(cellValue);
  if (list.length === 0) {
    throw new Error('No valid recipient email found');
  }
  if (list.length > 1) {
    throw new Error('Multiple emails provided for primary recipient; please provide exactly one');
  }
  return list[0];
}

function buildPlaceholderOrder(headers) {
  // Sort headers by descending length to avoid substring collisions
  return headers.slice().sort(function(a, b) { return String(b).length - String(a).length; });
}

function replacePlaceholders(text, headers, rowValues) {
  var ordered = buildPlaceholderOrder(headers);
  var result = String(text || '');
  ordered.forEach(function(header, index) {
    var key = String(header || '');
    var value = rowValues[index];
    var stringValue = (value === 0 || value) ? String(value) : '';
    var pattern = new RegExp('{{\\s*' + escapeRegex(key.toLowerCase()) + '\\s*}}', 'gi');
    result = result.replace(pattern, stringValue);
  });
  return result;
}

var OSMM_BATCH_SIZE = 50;

function getProgressKey(templateId) {
  return 'osmm_progress_' + String(templateId || 'default');
}

function readProgress(templateId) {
  var props = PropertiesService.getUserProperties();
  var val = props.getProperty(getProgressKey(templateId));
  return val ? parseInt(val, 10) : 0;
}

function writeProgress(templateId, index) {
  var props = PropertiesService.getUserProperties();
  props.setProperty(getProgressKey(templateId), String(index));
}

function clearProgress(templateId) {
  var props = PropertiesService.getUserProperties();
  props.deleteProperty(getProgressKey(templateId));
}

function resolveHtmlFile(possibleNames) {
  for (var i = 0; i < possibleNames.length; i++) {
    try {
      return HtmlService.createHtmlOutputFromFile(possibleNames[i]);
    } catch (error) {
      if (!String(error && error.message || '').includes('No HTML file named')) {
        throw error;
      }
    }
  }
  throw new Error('Could not find sidebar file. Please ensure either "Sidebar.html" or "sidebar.html" exists in your project.');
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  var html = resolveHtmlFile(['Sidebar.html', 'sidebar.html'])
    .setTitle(' ')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSheetData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetName = sheet.getName();
  var dataRange = sheet.getDataRange();
  var displayData = dataRange.getDisplayValues();

  var emailColumnIndex = -1;
  var ccColumnIndices = [];
  var bccColumnIndices = [];
  var filteredHeaders = [];
  var duplicateNonRecipientHeaders = [];

  var headers = (displayData && displayData.length > 0) ? displayData[0] : [];
  var seenNormalized = {};
  
  headers.forEach(function(header, index) {
    var headerStr = String(header || '').trim();
    var normalized = normalizeHeaderKey(headerStr);

    // Track duplicates (non-cc/bcc) for awareness; use first occurrence
    if (normalized && seenNormalized[normalized] !== undefined && !isCcHeader(normalized) && !isBccHeader(normalized)) {
      duplicateNonRecipientHeaders.push(headerStr + ' (duplicate at column ' + (index + 1) + ')');
    } else if (normalized) {
      seenNormalized[normalized] = index;
    }

    // Variables displayed exclude cc/bcc
    if (normalized && !isCcHeader(normalized) && !isBccHeader(normalized)) {
      filteredHeaders.push(headerStr);
    }

    if (isCcHeader(normalized)) {
      ccColumnIndices.push(index);
      return;
    }
    if (isBccHeader(normalized)) {
      bccColumnIndices.push(index);
      return;
    }

    if (emailColumnIndex === -1 && (normalized === 'email' || normalized === 'emailaddress' || normalized === 'emailaddressalias')) {
      emailColumnIndex = index;
    }
  });

  if (emailColumnIndex === -1) {
    var headerPreview = headers.map(function(h) { return '"' + String(h) + '"'; }).join(', ');
    throw new Error('Email column not found on sheet "' + sheetName + '". Add a header like "Email" or a supported variant. Headers found: ' + headerPreview + '.');
  }

  var rows = [];
  if (displayData && displayData.length > 1) {
    rows = displayData.slice(1).filter(function(row) {
      try {
        var emailCell = row[emailColumnIndex];
        return emailCell && String(emailCell).trim() !== '';
      } catch (e) {
        return false;
      }
    });
  }

  var preflightSummary = 'Sheet "' + sheetName + '" — recipients: ' + rows.length + ', cc cols: ' + ccColumnIndices.length + ', bcc cols: ' + bccColumnIndices.length + (duplicateNonRecipientHeaders.length ? (', duplicate headers: ' + duplicateNonRecipientHeaders.join('; ')) : '');

  return {
    sheetName: sheetName,
    headers: headers,
    displayHeaders: filteredHeaders,
    rows: rows,
    emailColumnIndex: emailColumnIndex,
    ccColumnIndices: ccColumnIndices,
    bccColumnIndices: bccColumnIndices,
    preflightSummary: preflightSummary
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
  var headers = data.headers.map(function(header) { return String(header).toLowerCase(); });
  var rows = data.rows;
  var emailColumnIndex = data.emailColumnIndex;
  var ccColumnIndices = data.ccColumnIndices;
  var bccColumnIndices = data.bccColumnIndices;
  var sheetName = data.sheetName;

  var startIndex = readProgress(templateId);
  var endIndex = Math.min(rows.length, startIndex + OSMM_BATCH_SIZE);

  var sent = 0;
  var errors = [];

  for (var i = startIndex; i < endIndex; i++) {
    var row = rows[i];
    var rowNumber = i + 2; // 1-based with header row
    try {
      var personalizedBody = replacePlaceholders(template.body, headers, row);
      var personalizedSubject = replacePlaceholders(template.subject, headers, row);

      // Validate primary email with row-aware error
      var toAddress;
      try {
        toAddress = parsePrimaryAddress(row[emailColumnIndex]);
      } catch (e) {
        var msgTo = 'Sheet \'' + sheetName + '\', row ' + rowNumber + ' (email=\'' + String(row[emailColumnIndex] || '') + '\'): ' + e.message;
        Logger.log(msgTo);
        errors.push(msgTo);
        continue;
      }

      // CC/BCC parsing with per-token validation and warnings
      var ccSet = {};
      var ccEmails = [];
      var ccInvalids = [];
      ccColumnIndices.forEach(function(ccIndex) {
        var raw = String(row[ccIndex] || '');
        var tokens = raw.split(/[;,]/).map(function(p) { return p.trim(); }).filter(function(p) { return p; });
        tokens.forEach(function(tok) {
          var addr = extractEmailAddress(tok);
          if (isLikelyValidEmail(addr)) {
            if (addr !== toAddress && !ccSet[addr]) {
              ccSet[addr] = true;
              ccEmails.push(addr);
            }
          } else {
            ccInvalids.push(tok);
          }
        });
      });

      var bccSet = {};
      var bccEmails = [];
      var bccInvalids = [];
      bccColumnIndices.forEach(function(bccIndex) {
        var rawB = String(row[bccIndex] || '');
        var tokensB = rawB.split(/[;,]/).map(function(p) { return p.trim(); }).filter(function(p) { return p; });
        tokensB.forEach(function(tok) {
          var addr = extractEmailAddress(tok);
          if (isLikelyValidEmail(addr)) {
            if (addr !== toAddress && !bccSet[addr]) {
              bccSet[addr] = true;
              bccEmails.push(addr);
            }
          } else {
            bccInvalids.push(tok);
          }
        });
      });

      // Warn about unresolved placeholders in subject/body
      var unresolved = [];
      var uBody = personalizedBody.match(/{{\s*[^}]+\s*}}/g) || [];
      var uSubject = personalizedSubject.match(/{{\s*[^}]+\s*}}/g) || [];
      if (uBody.length) unresolved = unresolved.concat(uBody);
      if (uSubject.length) unresolved = unresolved.concat(uSubject);

      var options = {
        htmlBody: personalizedBody,
        name: senderName || 'Mail Merge'
      };
      if (ccEmails.length > 0) options.cc = ccEmails.join(',');
      if (bccEmails.length > 0) options.bcc = bccEmails.join(',');

      GmailApp.sendEmail(
        toAddress,
        personalizedSubject,
        '',
        options
      );
      sent++;
      Utilities.sleep(250);

      // Row-level warnings appended to errors list for surfacing
      if (ccInvalids.length) {
        var warnCc = 'Sheet \'' + sheetName + '\', row ' + rowNumber + ': dropped invalid CC entries: ' + ccInvalids.join(', ');
        Logger.log(warnCc);
        errors.push(warnCc);
      }
      if (bccInvalids.length) {
        var warnBcc = 'Sheet \'' + sheetName + '\', row ' + rowNumber + ': dropped invalid BCC entries: ' + bccInvalids.join(', ');
        Logger.log(warnBcc);
        errors.push(warnBcc);
      }
      if (unresolved.length) {
        var seenPH = {};
        unresolved.forEach(function(p) { seenPH[p] = true; });
        var phList = Object.keys(seenPH);
        var warnPh = 'Sheet \'' + sheetName + '\', row ' + rowNumber + ': unresolved placeholders: ' + phList.join(', ');
        Logger.log(warnPh);
        errors.push(warnPh);
      }
    } catch (error) {
      var problematic = String(row[emailColumnIndex] || '');
      var errMsg = 'Sheet \'' + sheetName + '\', row ' + rowNumber + ' (email=\'' + problematic + '\'): ' + error.toString();
      Logger.log(errMsg);
      errors.push(errMsg);
    }
  }

  var finished = endIndex >= rows.length;
  if (!finished) {
    writeProgress(templateId, endIndex);
  } else {
    clearProgress(templateId);
  }

  var summary = (sent + ' emails sent in this run. ' + (finished ? 'All done.' : 'Resume needed to continue.')) + '\n' + data.preflightSummary;
  if (errors.length > 0) {
    summary += '\n' + errors.join('\n');
  }
  return summary;
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
  var headers = data.headers.map(function(header) { return String(header).toLowerCase(); });
  var firstRow = data.rows[0] || [];
  
  var personalizedBody = replacePlaceholders(template.body, headers, firstRow);
  var personalizedSubject = replacePlaceholders(template.subject, headers, firstRow);
  
  try {
    senderEmail = parsePrimaryAddress(senderEmail);
  } catch (e) {
    throw new Error('Invalid test recipient email: ' + e.message);
  }
  
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