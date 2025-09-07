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
      .setOpenLink(CardService.newOpenLink().setUrl('https://osmailmerge.com/privacy')))
    .addWidget(CardService.newTextButton()
      .setText('Terms of Service')
      .setOpenLink(CardService.newOpenLink().setUrl('https://osmailmerge.com/terms')));
  
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

  // Recompute display headers to exclude email, cc, and bcc columns
  filteredHeaders = headers.filter(function(h, idx) {
    if (idx === emailColumnIndex) return false;
    if (ccColumnIndices.indexOf(idx) !== -1) return false;
    if (bccColumnIndices.indexOf(idx) !== -1) return false;
    return String(h || '').trim() !== '';
  });

  // Enforce: duplicate headers are not allowed (except cc/bcc which are additive)
  if (duplicateNonRecipientHeaders.length > 0) {
    throw new Error('Duplicate headers not allowed (except cc/bcc): ' + duplicateNonRecipientHeaders.join('; '));
  }

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
    var rowHasError = false;
    try {
      var personalizedBody = replacePlaceholders(template.body, headers, row);
      var personalizedSubject = replacePlaceholders(template.subject, headers, row);

      // Validate primary email with row-aware error
      var toAddress;
      try {
        toAddress = parsePrimaryAddress(row[emailColumnIndex]);
      } catch (e) {
        var emailHeaderName = String(data.headers[emailColumnIndex] || 'email');
        var msgTo = emailHeaderName + ': Row ' + rowNumber + ' — ' + e.message + " (value: '" + String(row[emailColumnIndex] || '') + "')";
        Logger.log(msgTo);
        errors.push(msgTo);
        rowHasError = true; // don't continue; still validate CC/BCC
        toAddress = '';
      }

      // CC/BCC parsing with per-token validation and warnings
      var ccSet = {};
      var ccEmails = [];
      var ccInvalids = [];
      var ccRunningIndex = 0;
      ccColumnIndices.forEach(function(ccIndex) {
        var raw = String(row[ccIndex] || '');
        var tokens = raw.split(/[;,]/).map(function(p) { return p.trim(); }).filter(function(p) { return p; });
        tokens.forEach(function(tok) {
          ccRunningIndex++;
          var addr = extractEmailAddress(tok);
          if (isLikelyValidEmail(addr)) {
            if (addr !== toAddress && !ccSet[addr]) {
              ccSet[addr] = true;
              ccEmails.push(addr);
            }
          } else {
            ccInvalids.push({ index: ccRunningIndex, token: tok });
            rowHasError = true;
          }
        });
      });

      var bccSet = {};
      var bccEmails = [];
      var bccInvalids = [];
      var bccRunningIndex = 0;
      bccColumnIndices.forEach(function(bccIndex) {
        var rawB = String(row[bccIndex] || '');
        var tokensB = rawB.split(/[;,]/).map(function(p) { return p.trim(); }).filter(function(p) { return p; });
        tokensB.forEach(function(tok) {
          bccRunningIndex++;
          var addr = extractEmailAddress(tok);
          if (isLikelyValidEmail(addr)) {
            if (addr !== toAddress && !bccSet[addr]) {
              bccSet[addr] = true;
              bccEmails.push(addr);
            }
          } else {
            bccInvalids.push({ index: bccRunningIndex, token: tok });
            rowHasError = true;
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

      // Row-level warnings appended to errors list for surfacing
      if (ccInvalids.length) {
        ccInvalids.forEach(function(item) {
          var line = 'cc: Row ' + rowNumber + ' — Email ' + item.index + " invalid '" + item.token + "'";
          Logger.log(line);
          errors.push(line);
        });
      }
      if (bccInvalids.length) {
        bccInvalids.forEach(function(item) {
          var lineB = 'bcc: Row ' + rowNumber + ' — Email ' + item.index + " invalid '" + item.token + "'";
          Logger.log(lineB);
          errors.push(lineB);
        });
      }
      if (unresolved.length) {
        var seenPH = {};
        unresolved.forEach(function(p) { seenPH[p] = true; });
        var phList = Object.keys(seenPH);
        var warnPh = 'placeholders: Row ' + rowNumber + ' — unresolved ' + phList.join(', ');
        Logger.log(warnPh);
        errors.push(warnPh);
      }

      // Only send if there are no errors for this row
      if (!rowHasError) {
        GmailApp.sendEmail(
          toAddress,
          personalizedSubject,
          '',
          options
        );
        sent++;
        Utilities.sleep(250);
      }
    } catch (error) {
      var problematic = String(row[emailColumnIndex] || '');
      var errMsg = 'email: Row ' + rowNumber + " — " + error.toString() + " (value: '" + problematic + "')";
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

  var summary = (sent + ' emails sent in this run. ' + (finished ? '' : 'Resume needed to continue.'));
  if (errors.length > 0) {
    summary += '\nErrors (' + errors.length + '):\n' + errors.join('\n');
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
  // Build headers/row excluding recipient columns for templating in test mode
  var headers = [];
  var rowForTemplate = [];
  for (var h = 0; h < data.headers.length; h++) {
    if (h === data.emailColumnIndex) continue;
    if (data.ccColumnIndices.indexOf(h) !== -1) continue;
    if (data.bccColumnIndices.indexOf(h) !== -1) continue;
    headers.push(String(data.headers[h]).toLowerCase());
    rowForTemplate.push((data.rows[0] || [])[h]);
  }
  
  var personalizedBody = replacePlaceholders(template.body, headers, rowForTemplate);
  var personalizedSubject = replacePlaceholders(template.subject, headers, rowForTemplate);
  
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
  var link = CardService.newOpenLink().setUrl('https://osmailmerge.com/privacy');
  var action = CardService.newActionResponseBuilder().setOpenLink(link).build();
  return action;
}