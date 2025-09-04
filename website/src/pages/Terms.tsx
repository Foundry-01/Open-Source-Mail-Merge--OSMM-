import React from 'react';
import DocPage from '../components/DocPage';

export default function Terms(): JSX.Element {
  return (
    <DocPage
      title="Terms of Service"
      toc={[
        { id: 'acceptance', label: '1. Acceptance of Terms' },
        { id: 'license', label: '2. License Grant' },
        { id: 'no-warranty', label: '3. No Warranty' },
        { id: 'limitation', label: '4. Limitation of Liability' },
        { id: 'responsibilities', label: '5. User Responsibilities' },
        { id: 'arbitration', label: '6. Governing Law & Binding Arbitration' },
        { id: 'changes', label: '8. Changes to Terms' },
      ]}
    >
      <h2 id="acceptance">1. Acceptance of Terms</h2>
      <p>
        By using Open Source Mail Merge (OSMM), you agree to these Terms. If you disagree, do not use or install the Add‑on.
      </p>

      <h2 id="license">2. License Grant</h2>
      <p>
        OSMM is open-source under the MIT License, which grants rights to use, modify, distribute, and sublicense the software. Key aspects include: definitions, scope of usage rights, redistribution terms, modifications, warranty disclaimer, and limitation of liability.
      </p>

      <h2 id="no-warranty">3. No Warranty</h2>
      <p>
        The software is provided “as is”, without warranties of any kind. The developer is not liable for any issues arising from its use.
      </p>

      <h2 id="limitation">4. Limitation of Liability</h2>
      <p>
        In no event shall the author or contributors be liable for any damages, loss of data, or business interruption resulting from use of OSMM.
      </p>

      <h2 id="responsibilities">5. User Responsibilities</h2>
      <ul>
        <li>Ensure compliance with Gmail and Google Workspace’s sending limits and policies.</li>
        <li>Use responsibly and test via the “Send Test” function before mass sends to avoid being flagged as spam.</li>
      </ul>

      <h2 id="arbitration">6. Governing Law &amp; Binding Arbitration</h2>
      <p>
        These Terms are governed by the laws of the State of Indiana. <strong>By using OSMM, you agree to resolve any disputes through binding arbitration rather than in court.</strong>
      </p>
      <p>
        <strong>Arbitration Agreement</strong>: Any dispute, controversy, or claim arising out of or relating to these Terms or the use of OSMM shall be resolved exclusively through binding arbitration administered by the American Arbitration Association (AAA) under its Commercial Arbitration Rules. The arbitration will be conducted in Indiana.
      </p>
      <p>
        <strong>Waiver of Right to Sue</strong>: You expressly waive your right to bring any claim in a court of law and agree that arbitration is your sole remedy for resolving disputes. This includes waiving your right to a jury trial.
      </p>
      <p>
        <strong>Class Action Waiver</strong>: You agree that disputes must be brought on an individual basis only and waive any right to participate in a class action lawsuit or class-wide arbitration.
      </p>

      <h2 id="changes">8. Changes to Terms</h2>
      <p>
        We may update these terms occasionally; active users will be notified via GitHub or project documentation.
      </p>

      <p className="muted">Last updated: {new Date().toLocaleDateString()}</p>
    </DocPage>
  );
}


