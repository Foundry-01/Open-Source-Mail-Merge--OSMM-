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
    <style>
      body {
        font-family: 'Google Sans', Arial, sans-serif;
        color: #202124;
        margin: 0;
        padding: 0;
      }
      .container {
        padding: 16px;
      }
      h2 {
        color: #1a73e8;
        font-size: 20px;
        margin: 0 0 20px 0;
        padding-bottom: 8px;
        border-bottom: 2px solid #e8eaed;
      }
      h3 {
        color: #202124;
        font-size: 14px;
        margin: 24px 0 12px 0;
      }
      .input-group {
        margin-bottom: 24px;
        background: #f8f9fa;
        padding: 16px;
        border-radius: 8px;
      }
      label {
        display: block;
        margin-bottom: 8px;
        font-weight: 500;
        color: #202124;
      }
      input, select {
        width: 100%;
        padding: 8px 12px;
        border: 1px solid #dadce0;
        border-radius: 4px;
        font-size: 14px;
        margin-bottom: 8px;
      }
      input:focus, select:focus {
        outline: none;
        border-color: #1a73e8;
      }
      .helper-text {
        font-size: 12px;
        color: #5f6368;
        margin-top: 4px;
        line-height: 1.4;
      }
      .button {
        background-color: #1a73e8;
        color: white;
        padding: 8px 16px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        font-size: 14px;
        font-weight: 500;
        transition: background-color 0.2s;
      }
      .button:hover {
        background-color: #1557b0;
      }
      .button.send {
        background-color: #188038;
        width: 100%;
        padding: 12px;
        margin-top: 16px;
      }
      .button.send:hover {
        background-color: #137333;
      }
      #status {
        margin: 12px 0;
        padding: 8px;
        border-radius: 4px;
        font-size: 14px;
      }
      .success {
        background-color: #e6f4ea;
        color: #137333;
      }
      .error {
        background-color: #fce8e6;
        color: #c5221f;
      }
      .recipient {
        padding: 8px;
        margin: 4px 0;
        background: #f8f9fa;
        border-radius: 4px;
        font-size: 14px;
      }
    </style>
    <div class="container">
      <h2>Mail Merge</h2>
      
      <div class="input-group">
        <label for="senderName">Your Name</label>
        <input type="text" id="senderName" placeholder="Enter your name">
        <div class="helper-text">
          This only changes how your name appears in recipients' inboxes.<br>
          Make sure to include your signature in the email draft itself.
        </div>
      </div>
      
      <button onclick="loadData()" class="button">Load Recipients & Templates</button>
      <div id="status"></div>
      
      <div id="recipientList" style="display: none;">
        <h3>Recipients</h3>
        <div id="recipients"></div>
        
        <h3>Email Template</h3>
        <select id="templateSelect">
          <option value="">Loading drafts...</option>
        </select>
        
        <button onclick="sendEmails()" class="button send">Send Emails</button>
      </div>
    </div>

    <script>
      function loadData() {
        document.getElementById('status').innerHTML = '<div class="loading">Loading...</div>';
        google.script.run
          .withSuccessHandler(function(data) {
            document.getElementById('status').innerHTML = '<div class="success">✓ Data loaded successfully!</div>';
            document.getElementById('recipientList').style.display = 'block';
            
            // Display recipients
            var recipientHtml = '';
            data.rows.forEach(function(row) {
              recipientHtml += '<div class="recipient">' + row[0] + ' (' + row[1] + ')</div>';
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
            select.innerHTML = '<option value="">Select an email template...</option>';
            drafts.forEach(function(draft) {
              var option = document.createElement('option');
              option.value = draft.id;
              option.text = draft.subject || '(no subject)';
              select.appendChild(option);
            });
          })
          .withFailureHandler(function(error) {
            document.getElementById('status').innerHTML = '<div class="error">Error loading drafts: ' + error + '</div>';
          })
          .getDraftTemplates();
      }

      function sendEmails() {
        var templateId = document.getElementById('templateSelect').value;
        var senderName = document.getElementById('senderName').value.trim();
        
        if (!templateId) {
          alert('Please select an email template first');
          return;
        }
        
        if (!senderName) {
          alert('Please enter your name');
          return;
        }
        
        if (!confirm('Are you sure you want to send this email to all recipients?')) {
          return;
        }
        
        document.getElementById('status').innerHTML = '<div class="loading">Sending emails...</div>';
        google.script.run
          .withSuccessHandler(function(result) {
            document.getElementById('status').innerHTML = '<div class="success">✓ ' + result + '</div>';
          })
          .withFailureHandler(function(error) {
            document.getElementById('status').innerHTML = '<div class="error">Error: ' + error + '</div>';
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
      .setText('Start Mail Merge')
      .setOnClickAction(CardService.newAction().setFunctionName('showSidebar')));
  
  builder.addSection(section);
  return builder.build();
} 