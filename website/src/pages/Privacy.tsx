import React from 'react';

export default function Privacy(): JSX.Element {
  return (
    <article className="doc">
      <h1>Privacy Policy</h1>
      <p>Open Source Mail Merge (OSMM) operates entirely within your Google account. We do not store or transmit your data to external servers.</p>
      <h2>Information We Access</h2>
      <ul>
        <li>Google Sheets: to read recipient information</li>
        <li>Gmail: to read drafts and send email on your behalf</li>
        <li>Your email address: to send test emails and identify the sender</li>
      </ul>
      <h2>Data Storage</h2>
      <p>We do not store any personal information outside your Google Workspace account. All processing occurs via official Google APIs during active use of the add-on.</p>
      <h2>Contact</h2>
      <p>Questions? Open an issue on our GitHub repository.</p>
      <p className="muted">Last updated: {new Date().toLocaleDateString()}</p>
    </article>
  );
}


