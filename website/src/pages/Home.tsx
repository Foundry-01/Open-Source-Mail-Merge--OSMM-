import React from 'react';

export default function Home(): JSX.Element {
  return (
    <>
      <section className="hero">
        <div className="hero-copy">
          <h1>Personalized email at scale — no BCC, no fuss</h1>
          <p>
            OSMM is a free Google Sheets add-on that sends individual, personalized emails
            from your Gmail drafts. Perfect for marketing outreach, fundraising, and sales
            lead generation — all inside your Google account.
          </p>
          <ul className="hero-bullets">
            <li>Send one-to-one emails to large lists — no BCC</li>
            <li>Use variables like {'{{First Name}}'} from your Sheet</li>
            <li>Runs in your account — simple, safe, open-source</li>
          </ul>
          <div className="cta-row">
            <a className="button primary" href="https://workspace.google.com/marketplace/search?q=Open%20Source%20Mail%20Merge" target="_blank" rel="noreferrer">Get the add‑on</a>
            <a className="button" href="https://github.com/jdragon3001/Open-Source-Mail-Merge--OSMM-" target="_blank" rel="noreferrer">Get the code</a>
          </div>
          <div className="cta-under">
            <a className="link small" href="/privacy">Privacy</a>
          </div>
        </div>
        <div className="hero-art">
          <img src="/hero-illustration.svg" alt="Mail Merge illustration" />
        </div>
      </section>

      <section className="section">
        <div className="container">
          <h2>Why OSMM?</h2>
          <div className="grid">
            <div className="card">
              <h3>Personalization that performs</h3>
              <p>Use Sheet columns as variables to tailor each subject and message. Better relevance, better replies.</p>
            </div>
            <div className="card">
              <h3>No BCC — real conversations</h3>
              <p>Every recipient gets an individual email thread from your account. Higher trust and response rates.</p>
            </div>
            <div className="card">
              <h3>Private and simple</h3>
              <p>Runs within your Google account using official APIs. No external servers or data exports.</p>
            </div>
            <div className="card">
              <h3>Free and open-source</h3>
              <p>Transparent code you control. Fork it, adapt it, or use it as-is.</p>
            </div>
          </div>
        </div>
      </section>

      <section className="section alt">
        <div className="container">
          <h2>Common use-cases</h2>
          <div className="grid">
            <div className="card">
              <h3>Sales outreach</h3>
              <p>Personalized prospecting from a Sheet of leads. Start one-to-one conversations at scale.</p>
            </div>
            <div className="card">
              <h3>Fundraising</h3>
              <p>Thank donors, announce campaigns, and follow up with targeted, individualized messages.</p>
            </div>
            <div className="card">
              <h3>Community & education</h3>
              <p>Coordinate events, send updates, and keep communication personal without copy‑pasting.</p>
            </div>
          </div>
        </div>
      </section>

      <section className="section">
        <div className="container how">
          <h2>How it works</h2>
          <ol className="steps">
            <li><strong>Draft in Gmail</strong>: Write your message once and add variables like {'{{First Name}}'}.</li>
            <li><strong>Prepare your Sheet</strong>: Add columns for Email Address and any variables.</li>
            <li><strong>Send</strong>: Open the OSMM add-on in Sheets, select your draft, and send a test — then deliver to everyone.</li>
          </ol>
          <div className="cta-row">
            <a className="button primary" href="https://github.com/jdragon3001/Open-Source-Mail-Merge--OSMM-" target="_blank" rel="noreferrer">Get started on GitHub</a>
            <a className="button" href="/terms">Terms</a>
          </div>
          <p className="muted small">Subject to Gmail daily sending limits. Works with personal and Workspace accounts.</p>
        </div>
      </section>
    </>
  );
}


