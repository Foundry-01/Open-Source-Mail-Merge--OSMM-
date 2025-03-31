# Open Source Mail Merge (OSMM)

A free and simple mail merge tool that works directly with Google Sheets and Gmail drafts. Send personalized emails to multiple recipients using Gmail draft templates.

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
   - In your Google Sheet, click the "📧 Mail Merge" menu
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

## Support

This is an open-source project. For support:
- Check the troubleshooting section
- Review the code on GitHub
- Submit issues through the GitHub repository

## License

This project is licensed under the MIT License - see the LICENSE file for details. 