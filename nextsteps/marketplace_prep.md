# OSMM Marketplace Prep

Status: In progress

## Current step

Prepare a complete, review-ready plan for Google Workspace Marketplace publishing.

### Substeps (do these now)
- Create comprehensive checklist aligned to Google review criteria
- Inventory current manifest, scopes, UI, and assets
- Identify gaps and decisions (branding, hosting, consent)

## High-level plan

1) Scopes and manifest hygiene
- Audit scopes; remove `mail.google.com` if not strictly required
- Remove unused advanced services and `webapp` block if not used
- Ensure add-on metadata matches listing (name, colors)

2) Branding and assets
- Replace Google icon with OSMM square PNG, transparent background
- Prepare screenshots that demonstrate flows in Sheets sidebar
- Optional: promo banner (2200×1400) and video

3) Legal and hosting
- Public Privacy Policy URL
- Terms of Service URL (public)

4) OAuth consent
- External app, Production status
- Authorized domains include hosting for privacy/TOS/icon
- Scopes justification matches manifest and in-app use

5) Marketplace SDK configuration
- Select Editor add-on for Sheets
- Link to deployment, set universal navigation URL if needed

6) Listing content
- Name: Open Source Mail Merge (OSMM)
- Short and detailed descriptions; pricing info (Free, open-source)
- Developer name, website, support email/URL

7) QA and readiness
- Edge cases (no email column, invalid emails, empty rows)
- Loading indicators, error messages, sign-out behavior

8) Submission and follow-up
- Submit, track feedback, iterate

## Running notes

- Initial review of `appsscript.json` shows broad Gmail scopes including `https://mail.google.com/`; likely reducible to `gmail.send`, `gmail.readonly`, and add-on scopes.
- `logoUrl` uses Google icon; must replace with original OSMM asset.
- `webapp` block exists but add-on is Sheets-only; consider removal if unused.
- Privacy policy exists in repo; must host via public HTTPS and reference on consent + listing.

