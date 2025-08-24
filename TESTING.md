# OSMM Tester Setup (Developer add-on)

Follow these steps to let friends install and test OSMM without copying code.

## 1) Link to a standard Google Cloud project
1. Open the Apps Script editor for OSMM.
2. Project Settings (gear icon) → Google Cloud Platform (GCP) Project → Change project.
3. Choose an existing Cloud project or create a new one. Confirm the link.

## 2) Configure OAuth consent (External, Testing)
1. In Google Cloud Console → APIs & Services → OAuth consent screen.
2. User type: External. Publishing status: Testing.
3. App name: Open Source Mail Merge (OSMM).
4. App logo: upload your OSMM square, transparent PNG.
5. App domain: add the domain where you will host policies/assets (e.g., GitHub Pages domain).
6. Authorized domains: add the same domain.
7. Developer contact email: set your email.
8. App scopes: start minimal. You can add scopes when Apps Script requests them.
9. Test users: Add the Google accounts of your friends (up to 100).
10. Save.

Policy URLs (required):
- Privacy Policy: public HTTPS URL (temporary OK)
- Terms of Service: public HTTPS URL (temporary OK)

## 3) Narrow scopes in manifest (recommended before testing)
Edit `appsscript.json` to minimal scopes to reduce consent friction:
- Keep: `.../spreadsheets.currentonly`, `.../gmail.readonly` (read drafts), `.../gmail.send` (send), `.../script.container.ui`, `.../script.locale`.
- Remove if unused: `https://mail.google.com/`, `.../gmail.modify`, `.../gmail.compose`, `.../gmail.addons.execute`, `.../userinfo.email`.
- Remove Gmail Advanced Service and `webapp` section if not needed.

## 4) Create a Test deployment
1. In Apps Script: Deploy → Test deployments → Select type: Google Workspace Add-on → Configure Sheets.
2. Install the test deployment for your account.

## 5) How testers install the Developer add-on
Share the Apps Script project with testers (Viewer or Editor). Then testers:
1. Open Google Sheets → any spreadsheet.
2. Extensions → Add-ons → Manage add-ons → Developer add-ons tab.
3. Locate "Open Source Mail Merge" → Install.
4. During authorization, they’ll see the Testing consent screen since they’re on Test users.

## 6) Update process
- When you change code, create a new Test deployment version. Testers can revisit Developer add-ons to update.

## 7) Troubleshooting
- Not seeing the add-on under Developer add-ons? Ensure the tester is added on the OAuth consent screen as a Test user, and the project is shared with them.
- Authorization blocked? Check authorized domains and that policy URLs are valid HTTPS.
- Sidebar/UI not appearing? Reload the Sheet and confirm the add-on is installed for Sheets.

## 8) Ready for Marketplace later
When stable, switch OAuth consent to Production and follow Marketplace review steps per Google’s guides.
