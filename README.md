# Open Source Mail Merge (OSMM)

Apps Script project share link: https://script.google.com/d/1_2XrHuWGVz8DhP61xyYMA4GKceFtfSzXu_aJej8vu7Rbiw6j9OFYFeVY/edit?usp=sharing 

A free and open-source mail merge add-on for Google Sheets that lets you send personalized emails using Gmail drafts as templates.

## Installation Instructions

1. Create a new Google Apps Script project:
   - Go to [script.google.com](https://script.google.com)
   - Click "New Project"
   - Name your project "Open Source Mail Merge"

2. Enable the Apps Script Manifest:
   - Click on the ⚙️ settings icon (Project Settings)
   - Check "Show "appsscript.json" manifest file in editor"

3. Add the project files:
   - In the script editor, rename the default `Code.gs` file if it exists
   - Create/copy these files with their exact names:
     - `Code.gs`
     - `Sidebar.html`
     - `appsscript.json` (will appear after enabling in step 2)
   - Copy the provided code into each file

4. Save all files:
   - Click the 💾 save icon or press Ctrl+S (Cmd+S on Mac)

5. Run the initial setup:
   - In the script editor, select the `onHomepage` function from the dropdown menu at the top
   - Click the ▶️ Run button
   - Grant the necessary permissions when prompted
   - You might see a warning about the app not being verified - click "Advanced" and then "Go to Open Source Mail Merge (unsafe)"
   - Click through all permission requests

6. Deploy the add-on:
   - Click "Deploy" > "New deployment"
   - Click "Select type" > "Test deployments"
   - In the next screen, select "Google Workspace Add-on"
   - Install Sheets under Configuration
   - You'll get a confirmation that the test deployment is installed
   - Click "Done"

7. Use the add-on:
   - Open any Google Sheet
   - Refresh the page
   - Create a column with the heading "Email Address"
   - Create other columns with variables as needed
   - In your gmail, create a draft that will serve as your template
      - Include a subject, all content, and signature
      - In your gmail draft, tag these variables with {{variable}}
   - Click the little arrow on the bottom right corner of the screen if you are unable to see the workspace quick access toolbar on the right side of the screen (vertical down from your google profile picture)
   - Select the OSMM icon (black envelope)

## Features

- Send personalized emails to multiple recipients using Gmail drafts as templates
- Dynamic variable replacement using `{{variable}}` syntax
- Case-insensitive variable matching
- Automatic email column detection (searches for any column containing "email address")
- Test email functionality before sending to all recipients
- Modern, user-friendly interface

## How to Use

### Prepare Your Spreadsheet

1. Create a Google Sheet with your recipient data
2. Include a column with "email address" in its header (e.g., "Email Address", "Student Email Address", etc.)
3. Add any other columns you want to use as variables in your emails

### Create Your Email Template

1. Create a draft email in Gmail
2. Use variables in your subject or body by surrounding column names with double curly braces
   - Example: `Hello {{name}}, Your order #{{order number}} has shipped!`
3. Variables are case-insensitive, so `{{Name}}` and `{{name}}` work the same

### Send Your Emails

1. Click "Start" in the sidebar
2. Enter your name (will appear as the sender)
3. Select your email template from the dropdown
4. Send a test email first to verify everything works
5. Click "Send Emails" to send to all recipients

## Variables

- Use `{{column name}}` to insert data from any column
- Variables work in both subject line and email body
- Spaces in variable names are allowed: `{{First Name}}`
- Case doesn't matter: `{{first name}}` = `{{First Name}}` = `{{FIRST NAME}}`

## Tips

- Always send a test email first
- Check the "Available Variables" section to see what variables you can use
- Make sure your email address column has "email address" in its header
- Each recipient must have a valid email address

## Limitations

- Uses Gmail's daily sending limits
- Templates must be saved as Gmail drafts
- Maximum 5 most recent drafts shown in template selection

## Permissions Required

The add-on requires these permissions:
- View and manage your spreadsheets in Google Drive
- View and manage your email drafts and send email as you
- Display and run third-party web content in prompts and sidebars inside Google applications

## Support

This is an open-source project. For issues or feature requests, please create an issue in the GitHub repository.

## Setup Instructions

1. **Create a New Google Sheet**
   - Open Google Sheets
   - Create a new spreadsheet
   - Name your first two columns exactly as follows:
     - Column A: "First Name"
     - Column B: "Email Address"

2. **Add the Script**
   - In your Google Sheet, click on `Extensions` > `Apps Script`
   - Copy and paste the provided `Code.gs` content into the script editor (deleting any text already there)
   - Save the script (Ctrl/Cmd + S)
   - Run the script
   - Close the script editor

3. **Authorize the Script**
   - Return to your Google Sheet
   - Refresh the page
   - Click on the new "📧 Mail Merge" menu that appears
   - Click "Start Mail Merge"
   - When prompted, click "Continue" to grant permissions
   - Select your Google account
   - Click "Advanced" and then "Go to Mail Merge (unsafe)"
   - Click "Allow" to grant the necessary permissions

## Creating Email Templates

1. **Create a Draft in Gmail**
   - Go to Gmail
   - Click "Compose"
   - Write your email template
   - You can use `{{First Name}}` in your draft to personalize the email (it will be replaced with each recipient's name)
   - Save as draft (just close the compose window)

## Using the Mail Merge

1. **Prepare Your Recipient List**
   - In your Google Sheet, enter recipient information:
     - First Name in Column A
     - Email Address in Column B
     - Each row represents one recipient

2. **Send Your Mail Merge**
   - In your Google Sheet, click the "OSMM" menu
   - Click "Start Mail Merge"
   - In the sidebar that appears, click "Load Data"
   - Verify your recipients are listed correctly
   - Select your draft template from the dropdown
   - Click "Send Emails"
   - Confirm when prompted

## Troubleshooting

If you don't see your drafts:
1. Make sure you've authorized the script with all necessary permissions
2. Try refreshing the Google Sheet
3. Check that you have saved drafts in your Gmail account

If the menu doesn't appear:
1. Refresh your Google Sheet
2. Make sure you've saved the script
3. Check that you've granted all necessary permissions

## Limitations

- Works only with Gmail drafts
- Limited to Google Sheets
- Must be set up individually for each spreadsheet
- Maximum of 100 emails per day for regular Gmail accounts
- Maximum of 1500 emails per day for Google Workspace accounts

## Privacy & Security

- This script runs entirely within your Google account
- No data is sent to external servers
- All operations are performed through Google's official APIs
- The script requires permission to:
  - Access your Gmail drafts
  - Send emails on your behalf
  - Read and modify the current spreadsheet

