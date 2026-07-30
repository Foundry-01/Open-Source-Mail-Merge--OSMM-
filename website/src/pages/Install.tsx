import React from 'react';
import DocPage from '../components/DocPage';

const REPO = 'https://github.com/jdragon3001/Open-Source-Mail-Merge--OSMM-';

export default function Install(): JSX.Element {
  return (
    <DocPage
      title="Install OSMM"
      toc={[
        { id: 'what-you-need', label: 'What you need' },
        { id: 'open-editor', label: '1. Open the Apps Script editor' },
        { id: 'add-code', label: '2. Add Code.gs' },
        { id: 'add-sidebar', label: '3. Add Sidebar.html' },
        { id: 'add-manifest', label: '4. Add appsscript.json' },
        { id: 'authorize', label: '5. Authorize OSMM' },
        { id: 'open-osmm', label: '6. Open OSMM from the menu' },
        { id: 'permissions', label: 'What the permissions do' },
        { id: 'updating', label: 'Updating later' },
        { id: 'troubleshoot', label: 'Troubleshooting' },
      ]}
    >
      <p className="muted">
        OSMM installs as a script inside your own Google Sheet. It takes about five minutes,
        costs nothing, and everything stays in your Google account.
      </p>

      <h2 id="what-you-need">What you need</h2>
      <ul>
        <li>A Google account (personal Gmail or Workspace).</li>
        <li>A Google Sheet with your recipient list.</li>
        <li>The OSMM source code, which lives on <a className="link" href={REPO} target="_blank" rel="noreferrer">GitHub</a>.</li>
      </ul>

      <h2 id="open-editor">1. Open the Apps Script editor</h2>
      <ul>
        <li>Open the Google Sheet you want to send from.</li>
        <li>In the menu bar, choose <strong>Extensions → Apps Script</strong>.</li>
        <li>A new tab opens with an empty project and a file called <code>Code.gs</code>.</li>
      </ul>
      <p className="muted">
        The script belongs to this Sheet. Repeat these steps for any other Sheet you want to use,
        or just copy the Sheet — the script comes with it.
      </p>

      <h2 id="add-code">2. Add Code.gs</h2>
      <ul>
        <li>
          Open <a className="link" href={`${REPO}/blob/main/Code.gs`} target="_blank" rel="noreferrer">Code.gs</a> on
          GitHub and copy the whole file.
        </li>
        <li>Back in the editor, select everything in <code>Code.gs</code> and paste over it.</li>
        <li>Click the save icon.</li>
      </ul>

      <h2 id="add-sidebar">3. Add Sidebar.html</h2>
      <ul>
        <li>In the editor's Files list, click <strong>+ → HTML</strong>.</li>
        <li>
          Name it <code>Sidebar</code> exactly. Apps Script adds the <code>.html</code> for you,
          so the file becomes <code>Sidebar.html</code>.
        </li>
        <li>
          Copy <a className="link" href={`${REPO}/blob/main/Sidebar.html`} target="_blank" rel="noreferrer">Sidebar.html</a> from
          GitHub, paste over everything in the new file, and save.
        </li>
      </ul>

      <h2 id="add-manifest">4. Add appsscript.json</h2>
      <p>
        This file tells Google exactly which permissions OSMM needs. Skip it and Apps Script guesses,
        which asks for full access to your entire mailbox instead of just drafts and sending.
      </p>
      <ul>
        <li>In the editor's left sidebar, click <strong>Project Settings</strong> (the gear icon).</li>
        <li>Tick <strong>Show "appsscript.json" manifest file in editor</strong>.</li>
        <li>Go back to <strong>Editor</strong>. A file called <code>appsscript.json</code> is now in the list.</li>
        <li>
          Copy <a className="link" href={`${REPO}/blob/main/appsscript.json`} target="_blank" rel="noreferrer">appsscript.json</a> from
          GitHub, paste over everything in that file, and save.
        </li>
        <li>
          Optional: change <code>"timeZone"</code> to your own, for example
          {' '}<code>"America/Los_Angeles"</code>.
        </li>
      </ul>

      <h2 id="authorize">5. Authorize OSMM</h2>
      <ul>
        <li>In the editor toolbar, pick <code>showSidebar</code> from the function dropdown and click <strong>Run</strong>.</li>
        <li>Google asks for permission. Choose your account.</li>
        <li>
          You will see <em>"Google hasn't verified this app"</em>. That is expected — the app is
          your own copy of the code, not a published one. Click <strong>Advanced</strong>, then
          <strong> Go to (project name) (unsafe)</strong>.
        </li>
        <li>Review the permissions and click <strong>Allow</strong>.</li>
      </ul>

      <h2 id="open-osmm">6. Open OSMM from the menu</h2>
      <ul>
        <li>Go back to your Sheet and refresh the page.</li>
        <li>A new <strong>OSMM</strong> menu appears next to Help.</li>
        <li>Choose <strong>OSMM → Open Mail Merge</strong> to open the sidebar.</li>
      </ul>
      <p>
        You're installed. Next: <a className="link" href="/post-install">how to use OSMM</a>.
      </p>

      <h2 id="permissions">What the permissions do</h2>
      <ul>
        <li><strong>Read this spreadsheet</strong> — to load your recipient list and variables.</li>
        <li><strong>Read your Gmail drafts</strong> — to use a draft as the email template.</li>
        <li><strong>Send email as you</strong> — to send each personalized message.</li>
        <li><strong>Show a sidebar</strong> — to display the OSMM interface.</li>
      </ul>
      <p className="muted">
        Nothing is sent anywhere else. There is no OSMM server. Read the code yourself on
        {' '}<a className="link" href={REPO} target="_blank" rel="noreferrer">GitHub</a>, or see our
        {' '}<a className="link" href="/privacy">Privacy Policy</a>.
      </p>

      <h2 id="updating">Updating later</h2>
      <p>
        To pick up a newer version, copy the latest <code>Code.gs</code>, <code>Sidebar.html</code> and
        {' '}<code>appsscript.json</code> from GitHub and paste over the three files again. If the permissions
        changed, Google will ask you to re-authorize the next time you run it.
      </p>

      <h2 id="troubleshoot">Troubleshooting</h2>
      <ul>
        <li><strong>No OSMM menu.</strong> Refresh the Sheet. If it still doesn't appear, run <code>showSidebar</code> once from the editor to finish authorizing.</li>
        <li><strong>"Could not find sidebar file."</strong> The HTML file is named something other than <code>Sidebar</code>. Rename it.</li>
        <li><strong>Blank sidebar.</strong> <code>Sidebar.html</code> was pasted incompletely. Copy the whole file again.</li>
        <li><strong>No drafts in the dropdown.</strong> Save the draft in Gmail first, then reopen the sidebar.</li>
      </ul>
    </DocPage>
  );
}
