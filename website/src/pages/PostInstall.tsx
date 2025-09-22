import React from 'react';
import DocPage from '../components/DocPage';

export default function PostInstall(): JSX.Element {
  return (
    <DocPage
      title="Post Install: Using OSMM"
      toc={[
        { id: 'open-addon', label: '1. Open the Add‑on in Sheets' },
        { id: 'prepare-sheet', label: '2. Prepare Your Sheet' },
        { id: 'create-draft', label: '3. Create a Gmail Draft Template' },
        { id: 'load-and-test', label: '4. Load Recipients & Send a Test' },
        { id: 'send-emails', label: '5. Send Emails' },
        { id: 'variables', label: 'Tips: Variables & CC' },
        { id: 'limits', label: 'Gmail Sending Limits' },
        { id: 'troubleshoot', label: 'Troubleshooting' },
      ]}
    >
      <p className="muted">A quick guide to using Open Source Mail Merge after installation.</p>

      <h2 id="open-addon">1. Open the Add‑on in Sheets</h2>
      <p>
        Open your Google Sheet, refresh the page, and look for the OSMM icon in the right‑hand sidebar. Click it to open the add‑on.
      </p>

      <h2 id="prepare-sheet">2. Prepare Your Sheet</h2>
      <ul>
        <li>
          Create an <strong>Email Address</strong> column
          <ul>
            <li>Case‑insensitive header matching</li>
            <li><code>email</code> also works</li>
          </ul>
        </li>
        <li>Add any other columns you want to use as variables (for example, <code>{'{{First Name}}'}</code>, <code>{'{{Company}}'}</code>).</li>
        <li>Each row represents a single recipient.</li>
      </ul>

      <h2 id="create-draft">3. Create a Gmail Draft Template</h2>
      <ul>
        <li>In Gmail, compose a draft email that will become your template.</li>
        <li>You can reference Sheet columns using <code>{'{{variable}}'}</code> syntax in the subject and body.</li>
        <li>Example: Subject: <code>{'Hello {{First Name}}'}</code>.</li>
      </ul>

      <h2 id="load-and-test">4. Load Recipients &amp; Send a Test</h2>
      <ul>
        <li>In the OSMM sidebar, click <strong>Start</strong> to load recipients.</li>
        <li>Select your Gmail draft from the dropdown.</li>
        <li>Click <strong>Send Test Email</strong> to send a preview to yourself. Verify formatting and variable replacements.</li>
      </ul>

      <h2 id="send-emails">5. Send Emails</h2>
      <ul>
        <li>After confirming the test looks good, click <strong>Send Emails</strong>.</li>
        <li>OSMM sends individual emails to each recipient from your account.</li>
      </ul>

      <h2 id="variables">Tips: Variables, CC &amp; BCC</h2>
      <ul>
        <li>Variables are case‑insensitive and can include spaces: <code>{'{{First Name}}'}</code> equals <code>{'{{first name}}'}</code>.</li>
        <li>Add a column named <code>cc</code> or <code>bcc</code> to include those recipients. Multiple addresses can be separated by commas.</li>
        <li>Use the <strong>Available Variables</strong> panel in the sidebar to see which headers are detected.</li>
      </ul>

      <h2 id="limits">Gmail Sending Limits</h2>
      <p className="muted">Respect Gmail daily limits:</p>
      <ul>
        <li>Personal accounts: ~500/day</li>
        <li>Google Workspace accounts: ~2,000/day</li>
      </ul>

      <h2 id="troubleshoot">Troubleshooting</h2>
      <ul>
        <li>If drafts don’t appear, refresh the Sheet and ensure the draft is saved in Gmail.</li>
        <li>If recipients don’t load, verify your sheet has a valid email column and no duplicate headers.</li>
        <li>Always send a test email before bulk sending.</li>
      </ul>

      <p>
        Need details? See <a className="link" href="/privacy">Privacy</a> and <a className="link" href="/terms">Terms</a>, or visit the
        {' '}<a className="link" href="https://github.com/jdragon3001/Open-Source-Mail-Merge--OSMM-" target="_blank" rel="noreferrer">GitHub repository</a>.
      </p>
    </DocPage>
  );
}


