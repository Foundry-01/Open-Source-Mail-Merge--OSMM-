import React from 'react';
import DocPage from '../components/DocPage';

export default function Privacy(): JSX.Element {
  return (
    <DocPage
      title="Privacy Policy"
      toc={[
        { id: 'introduction', label: '1. Introduction' },
        { id: 'data-access-use', label: '2. Data Access & Use' },
        { id: 'third-party', label: '3. Third-Party Sharing & Transmission' },
        { id: 'data-retention', label: '4. Data Retention' },
        { id: 'security', label: '5. Security Measures' },
        { id: 'your-rights', label: '6. Your Rights' },
        { id: 'contact-updates', label: '7. Contact & Policy Updates' },
      ]}
    >
      <h2 id="introduction">1. Introduction</h2>
      <p>
        This Privacy Policy applies to Open Source Mail Merge (OSMM), a Google Workspace Add‑on. It explains how OSMM handles data when installed and used within your Google Workspace account.
      </p>

      <h2 id="data-access-use">2. Data Access &amp; Use</h2>
      <ul>
        <li>
          <strong>Scope access</strong>: OSMM uses Google OAuth scopes to read from Google Sheets (to fetch recipient data and merge fields) and Gmail (to access your drafts and send emails). No additional unnecessary scopes are requested.
        </li>
        <li>
          <strong>Data flow</strong>: All processing occurs entirely within your Google account using Google Apps Script. No data is transmitted or stored externally.
        </li>
        <li>
          <strong>Data stays within your control</strong>: No third-party servers, analytics, or telemetry are used. Even any user-created logs (e.g., to Sheets) remain in your account.
        </li>
      </ul>

      <h2 id="third-party">3. Third-Party Sharing &amp; Transmission</h2>
      <p>OSMM does <strong>not</strong> share any user or project data with third parties or external services.</p>

      <h2 id="data-retention">4. Data Retention</h2>
      <p>OSMM does not retain user data itself. Any logs you choose to create remain in your own Sheets, under your management.</p>

      <h2 id="security">5. Security Measures</h2>
      <p>Processing happens within Google’s secure infrastructure under your account. OSMM does not introduce external vulnerabilities.</p>

      <h2 id="your-rights">6. Your Rights</h2>
      <p>You can uninstall OSMM or revoke its permissions at any time via your Google Account &gt; Security &gt; Third‑party apps.</p>

      <h2 id="contact-updates">7. Contact &amp; Policy Updates</h2>
      <p>Your use of OSMM constitutes acceptance of this policy. We may update it when necessary; users will be informed via GitHub releases or documentation update.</p>

      <p>
        See also: <a className="link" href="/terms">Terms of Service</a>
      </p>

      <p className="muted">Last updated: {new Date().toLocaleDateString()}</p>
    </DocPage>
  );
}


