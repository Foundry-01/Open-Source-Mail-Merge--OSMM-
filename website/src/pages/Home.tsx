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
          <h2>Built for nonprofits, clubs, and businesses</h2>
          <p className="section-subtitle">Use Gmail's sending limits to their fullest: 500/day (personal) • 2,000/day (Workspace).</p>
          <div className="grid grid-3">
            <div className="card accent">
              <h3>Personalization that converts</h3>
              <p>Use Sheet columns as variables to customize each subject and message. Higher engagement and better response rates.</p>
            </div>
            <div className="card accent">
              <h3>Individual conversations</h3>
              <p>No BCC mass emails. Every recipient gets a personal thread from your account, building trust and encouraging replies.</p>
            </div>
            <div className="card accent">
              <h3>Secure & cost‑effective</h3>
              <p>Runs entirely within your Google account using official APIs. No monthly fees, external servers, or data exports required.</p>
            </div>
          </div>
        </div>
      </section>

      {/** Consolidated to one set of cards; second section removed for clarity **/}

      <section className="section">
        <div className="container">
          <h2>Simple setup, powerful results</h2>
          <div className="how-it-works">
            <div className="step">
              <div className="step-number">1</div>
              <div className="step-content">
                <h3>Draft in Gmail</h3>
                <p>Write your message once and add variables like <code>{'{{First Name}}'}</code> or <code>{'{{Company}}'}</code>. Your template becomes the foundation for personalized outreach.</p>
              </div>
            </div>
            <div className="step">
              <div className="step-number">2</div>
              <div className="step-content">
                <h3>Prepare your Sheet</h3>
                <p>Create columns for Email Address and any variables you used. Each row becomes a personalized email to that recipient.</p>
              </div>
            </div>
            <div className="step">
              <div className="step-number">3</div>
              <div className="step-content">
                <h3>Send & scale</h3>
                <p>Open the OSMM add-on in Sheets, select your draft, send a test email first — then deliver to everyone on your list.</p>
              </div>
            </div>
          </div>
          <div className="cta-section">
            <div className="cta-row">
              <a className="button primary" href="https://workspace.google.com/marketplace/search?q=Open%20Source%20Mail%20Merge" target="_blank" rel="noreferrer">Get the add‑on</a>
              <a className="button" href="https://github.com/jdragon3001/Open-Source-Mail-Merge--OSMM-" target="_blank" rel="noreferrer">View on GitHub</a>
            </div>
            <p className="limits-note">Respects Gmail's daily limits: 500 emails (personal accounts) • 2,000 emails (Workspace accounts)</p>
          </div>
        </div>
      </section>
    </>
  );
}


