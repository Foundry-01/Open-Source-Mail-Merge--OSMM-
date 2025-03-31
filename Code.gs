/**
 * Creates a menu entry in the Google Sheets UI when the document is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📧 Mail Merge')
    .addItem('Start Mail Merge', 'showSidebar')
    .addToUi();
}

function doGet() {
  return HtmlService.createHtmlOutput('This is a web app.');
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  var html = HtmlService.createHtmlOutput(`
    <div style="padding: 10px;">
      <h2>Mail Merge</h2>
      
      <div style="margin-bottom: 20px;">
        <label for="senderName" style="display: block; margin-bottom: 5px;">Your Name:</label>
        <input type="text" id="senderName" style="width: 100%; padding: 5px; margin-bottom: 10px;" placeholder="Enter your name">
      </div>
      
      <button onclick="loadData()" style="margin-bottom: 10px; background-color: #4CAF50; color: white; padding: 8px 15px; border: none; border-radius: 4px; cursor: pointer;">Load Data</button>
      <div id="status"></div>
      
      <div id="recipientList" style="display: none; margin-top: 20px;">
        <h3>Recipients:</h3>
        <div id="recipients"></div>
        
        <h3 style="margin-top: 20px;">Select Template:</h3>
        <select id="templateSelect" style="margin-bottom: 10px; width: 100%; padding: 5px;">
          <option value="">Loading drafts...</option>
        </select>
        
        <button onclick="sendEmails()" style="margin-top: 10px; background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">Send Emails</button>
      </div>
    </div>

    <script>
      function loadData() {
        document.getElementById('status').innerHTML = 'Loading...';
        google.script.run
          .withSuccessHandler(function(data) {
            document.getElementById('status').innerHTML = 'Success!';
            document.getElementById('recipientList').style.display = 'block';
            
            // Display recipients
            var recipientHtml = '';
            data.rows.forEach(function(row) {
              recipientHtml += '<div style="margin: 5px 0;">' + row[0] + ' (' + row[1] + ')</div>';
            });
            document.getElementById('recipients').innerHTML = recipientHtml;
            
            // Load draft templates
            loadDraftTemplates();
          })
          .getSheetData();
      }

      function loadDraftTemplates() {
        google.script.run
          .withSuccessHandler(function(drafts) {
            var select = document.getElementById('templateSelect');
            select.innerHTML = '<option value="">Select a draft template...</option>';
            drafts.forEach(function(draft) {
              var option = document.createElement('option');
              option.value = draft.id;
              option.text = draft.subject || '(no subject)';
              select.appendChild(option);
            });
          })
          .withFailureHandler(function(error) {
            document.getElementById('status').innerHTML = 'Error loading drafts: ' + error;
          })
          .getDraftTemplates();
      }

      function sendEmails() {
        var templateId = document.getElementById('templateSelect').value;
        var senderName = document.getElementById('senderName').value.trim();
        
        if (!templateId) {
          alert('Please select a template first');
          return;
        }
        
        if (!senderName) {
          alert('Please enter your name');
          return;
        }
        
        if (!confirm('Are you sure you want to send this email to all recipients?')) {
          return;
        }
        
        document.getElementById('status').innerHTML = 'Sending emails...';
        google.script.run
          .withSuccessHandler(function() {
            document.getElementById('status').innerHTML = 'Emails sent successfully!';
          })
          .withFailureHandler(function(error) {
            document.getElementById('status').innerHTML = 'Error: ' + error;
          })
          .sendMailMerge(templateId, senderName);
      }
    </script>
  `)
    .setTitle('Mail Merge')
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
    return drafts.map(function(draft) {
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
      .setText('Start Mail Merge')
      .setOnClickAction(CardService.newAction().setFunctionName('showSidebar')));
  
  builder.addSection(section);
  return builder.build();
} 