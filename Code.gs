/**
 * Creates a menu entry in the Google Sheets UI when the document is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('OSMM')
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
        background: #fff;
      }
      .container {
        padding: 20px;
      }
      .title {
        display: flex;
        align-items: center;
        margin-bottom: 24px;
        padding-bottom: 12px;
        border-bottom: 1px solid #dadce0;
      }
      .title-icon {
        width: 20px;
        height: 20px;
        margin-right: 12px;
        background: #e8f0fe;
        border-radius: 4px;
        display: flex;
        align-items: center;
        justify-content: center;
      }
      h2 {
        color: #202124;
        font-size: 16px;
        margin: 0;
        font-weight: 500;
      }
      .input-section {
        margin-bottom: 24px;
      }
      label {
        display: block;
        font-size: 14px;
        color: #202124;
        margin-bottom: 8px;
        font-weight: 500;
      }
      input {
        width: 100%;
        padding: 8px 0;
        border: none;
        border-bottom: 1px solid #dadce0;
        font-size: 14px;
        background: transparent;
        margin-bottom: 4px;
      }
      input:focus {
        outline: none;
        border-bottom-color: #1a73e8;
      }
      .helper-text {
        font-size: 12px;
        color: #5f6368;
        margin-top: 4px;
      }
      .load-button {
        display: flex;
        align-items: center;
        color: #1a73e8;
        font-size: 14px;
        padding: 0;
        background: none;
        border: none;
        cursor: pointer;
        font-weight: 500;
      }
      .load-button.done {
        color: #188038;
      }
      .load-button .check {
        margin-right: 8px;
      }
      .help-icon {
        color: #5f6368;
        cursor: pointer;
        margin-left: auto;
      }
      .section-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin: 24px 0 12px;
      }
      .section-title {
        font-size: 14px;
        color: #202124;
        font-weight: 500;
      }
      .action-button {
        color: #1a73e8;
        background: none;
        border: none;
        font-size: 14px;
        font-weight: 500;
        cursor: pointer;
        padding: 0;
      }
      .recipient-list {
        margin-bottom: 24px;
      }
      .recipient {
        display: flex;
        align-items: center;
        padding: 8px 0;
        border-bottom: 1px solid #dadce0;
      }
      .recipient:last-child {
        border-bottom: none;
      }
      .recipient-initial {
        width: 32px;
        height: 32px;
        border-radius: 50%;
        background: #e8f0fe;
        color: #1967d2;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 14px;
        margin-right: 12px;
        font-weight: 500;
      }
      .recipient-info {
        flex-grow: 1;
      }
      .recipient-name {
        font-size: 14px;
        color: #202124;
        margin-bottom: 2px;
      }
      .recipient-email {
        font-size: 12px;
        color: #5f6368;
      }
      .more-actions {
        color: #5f6368;
        padding: 4px;
        cursor: pointer;
      }
      select {
        width: 100%;
        padding: 8px 0;
        border: none;
        border-bottom: 1px solid #dadce0;
        font-size: 14px;
        background: transparent;
        margin-bottom: 24px;
        color: #202124;
        -webkit-appearance: none;
        appearance: none;
        background-image: url('data:image/svg+xml;charset=US-ASCII,<svg width="24" height="24" xmlns="http://www.w3.org/2000/svg"><path d="M7 10l5 5 5-5z" fill="%235F6368"/></svg>');
        background-repeat: no-repeat;
        background-position: right center;
      }
      select:focus {
        outline: none;
        border-bottom-color: #1a73e8;
      }
      .send-button {
        width: 100%;
        background: #202124;
        color: white;
        border: none;
        border-radius: 4px;
        padding: 12px;
        font-size: 14px;
        font-weight: 500;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
      }
      .send-button svg {
        margin-right: 8px;
      }
      .send-button:hover {
        background: #000;
      }
      #status {
        margin: 8px 0;
        font-size: 14px;
      }
    </style>
    <div class="container">
      <div class="title">
        <div class="title-icon">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="#1967d2">
            <path d="M20 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 4l-8 5-8-5V6l8 5 8-5v2z"/>
          </svg>
        </div>
        <h2>Open Source Mail Merge</h2>
      </div>
      
      <div class="input-section">
        <label for="senderName">Your Name</label>
        <input type="text" id="senderName" placeholder="Enter your name">
        <div class="helper-text">
          This name will appear in recipients' inboxes. Include your signature in the draft template.
        </div>
      </div>
      
      <button onclick="loadData()" id="loadButton" class="load-button">
        <span id="loadText">Load Recipients & Templates</span>
      </button>
      <div id="status"></div>
      
      <div id="recipientList" style="display: none;">
        <div class="section-header">
          <span class="section-title">Recipients</span>
          <button class="action-button">Preview</button>
        </div>
        <div class="recipient-list" id="recipients"></div>
        
        <div class="section-header">
          <span class="section-title">Email Template</span>
        </div>
        <select id="templateSelect">
          <option value="">Loading drafts...</option>
        </select>
        
        <button onclick="sendEmails()" class="send-button">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor">
            <path d="M3.4 20.4l17.45-7.48c.81-.35.81-1.49 0-1.84L3.4 3.6c-.66-.29-1.39.2-1.39.91L2 9.12c0 .5.37.93.87.99L17 12 2.87 13.88c-.5.07-.87.5-.87 1l.01 4.61c0 .71.73 1.2 1.39.91z"/>
          </svg>
          Send Emails
        </button>
      </div>
    </div>

    <script>
      function loadData() {
        var loadButton = document.getElementById('loadButton');
        loadButton.disabled = true;
        document.getElementById('status').innerHTML = '<div style="color: #1a73e8;">Loading...</div>';
        
        google.script.run
          .withSuccessHandler(function(data) {
            loadButton.innerHTML = '<span class="check">✓</span> Load Recipients & Templates';
            loadButton.classList.add('done');
            document.getElementById('recipientList').style.display = 'block';
            
            // Display recipients
            var recipientHtml = '';
            data.rows.forEach(function(row) {
              var initial = row[0].charAt(0).toUpperCase();
              recipientHtml += \`
                <div class="recipient">
                  <div class="recipient-initial">\${initial}</div>
                  <div class="recipient-info">
                    <div class="recipient-name">\${row[0]}</div>
                    <div class="recipient-email">\${row[1]}</div>
                  </div>
                  <div class="more-actions">⋮</div>
                </div>
              \`;
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
            document.getElementById('status').innerHTML = '<div style="color: #d93025;">Error loading drafts: ' + error + '</div>';
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
        
        document.getElementById('status').innerHTML = '<div style="color: #1a73e8;">Sending emails...</div>';
        google.script.run
          .withSuccessHandler(function(result) {
            document.getElementById('status').innerHTML = '<div style="color: #188038;">' + result + '</div>';
          })
          .withFailureHandler(function(error) {
            document.getElementById('status').innerHTML = '<div style="color: #d93025;">Error: ' + error + '</div>';
          })
          .sendMailMerge(templateId, senderName);
      }
    </script>
  `)
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
      .setText('Start Mail Merge')
      .setOnClickAction(CardService.newAction().setFunctionName('showSidebar')));
  
  builder.addSection(section);
  return builder.build();
} 