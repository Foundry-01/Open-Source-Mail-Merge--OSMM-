import React from 'react';

export default function Terms(): JSX.Element {
  return (
    <article className="doc">
      <h1>Terms of Service</h1>
      <p>By using OSMM, you agree to use the software responsibly and in compliance with Gmail and Google Workspace policies, including sending limits and anti-spam rules.</p>
      <h2>License</h2>
      <p>OSMM is open-source and provided under the Apache-2.0 license, without warranty of any kind.</p>
      <h2>Acceptable Use</h2>
      <ul>
        <li>No unsolicited or spam messages.</li>
        <li>Respect recipient preferences and applicable laws (e.g., CAN-SPAM/GDPR).</li>
      </ul>
      <p className="muted">Last updated: {new Date().toLocaleDateString()}</p>
    </article>
  );
}


