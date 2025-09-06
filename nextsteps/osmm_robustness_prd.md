## OS Mail Merge (Google Sheets Add-on) — Robustness & Reliability PRD / Spec

### Context
- The add-on sends personalized emails from Gmail drafts using data in the active Google Sheet (`SpreadsheetApp.getActiveSheet()`).
- Current issues: brittle header detection, inconsistent CC/BCC handling, unsafe placeholder replacement, vague error messages, noisy logging, and long-run timeouts.
- Scope: Apps Script only (`Code.gs`). No site/website changes. Keep active-sheet-only behavior.

### Objectives
- Make sending deterministic and resilient without changing core UX.
- Ensure predictable header detection, robust address parsing/validation, strong error reporting, privacy-safe logging, and quota-aware, resumable sending.

### In Scope
- Header normalization and aliasing (SSOT)
- CC and BCC collection/parsing
- Safe placeholder replacement
- Actionable error messages
- Logging minimization
- Chunked sending + resumability
- Use of display values for templating
- Centralized UI asset resolver

### Out of Scope
- Multi-sheet selection UI
- External services or dependencies
- Changes to draft selection UI/flow

### User Stories
- As a user, I can name the email column with variants (e.g., "e-mail address") and it’s detected reliably.
- As a user, I can place multiple CCs/BCCs separated by commas or semicolons, optionally with display names, and they’re parsed correctly.
- As a user, placeholder replacement is predictable even when headers overlap (e.g., "name" and "first name").
- As a user, failures indicate the exact sheet and row with a clear reason.
- As a user, large sends don’t time out and can be resumed without duplicating sent emails.

### Functional Requirements

#### 1. Header Detection SSOT
- Normalize headers by: lowercase and strip all non-alphanumeric characters.
- Canonical keys and aliases (case-insensitive):
  - email: email, emailaddress, email_address, email-address, e-mail, e-mailaddress, e-mail_address, e-mail-address
  - cc: cc
  - bcc: bcc
- Duplicates:
  - For cc/bcc: treat duplicates as additional recipient columns.
  - For others: warn in preflight summary; first occurrence is used for substitution.

#### 2. CC/BCC Parsing
- Collect indices for both CC and BCC headers.
- For each cell, split list by comma or semicolon, trim each.
- Support display names: extract addr from formats like `Name <addr@x.com>`.
- Validate per-address; store only valid addresses.
- Dedupe addresses within each bucket per row.

#### 3. Primary Address Parsing
- Allow `Name <addr@x.com>` in the primary email column.
- If multiple addresses are present, error with clear remediation.
- Use the same parsing/validation pipeline as CC/BCC; enforce exactly one address.

#### 4. Placeholder Replacement
- Escape keys when building regexes.
- Replace in descending key-length order to avoid substring collisions.
- Case-insensitive match for `{{ key }}` with arbitrary whitespace.
- Future-friendly: leave room for `{{key|default}}` support.

#### 5. Error Messages
- Include sheet name and 1-based row number.
- Include a short snippet of the problematic value.
- Example: `Sheet 'Leads', row 17 (email='Name <bad@>'): Invalid email address`.

#### 6. Logging & Privacy
- Remove bulk data logs.
- Log header detection summary, counts, and per-row failures only.
- Provide a concise preflight summary before sending.

#### 7. Active Sheet Only
- Continue using `SpreadsheetApp.getActiveSheet()`.
- Include active sheet name in summaries and errors to avoid confusion.

#### 8. Display Values for Templating
- Use `getDisplayValues()` for substitution to match user-visible content.
- Keep `getValues()` only if needed for non-templating logic.

#### 9. Email Validation & Deduplication
- Single parsing/validation utility for To/CC/BCC.
- Gentle validation suitable for Gmail; allow common, valid addresses.
- Deduplicate within each bucket; do not cross-add To into CC/BCC.

#### 10. Quota-Aware Sending & Resumability
- Send in chunks (configurable batch size, default 50 per run).
- Maintain progress via `PropertiesService.getUserProperties()` with last processed row index.
- Estimate remaining execution time; stop early with a "resume needed" message.
- Small sleep between sends; consider simple backoff on repeated errors.

#### 11. UI Asset Resolver
- Centralize resolution for sidebar/html assets with a preferred order and consistent error.

### Non-Functional Requirements
- No new external dependencies.
- Maintain or improve current script performance.
- Clear, commented utilities for maintainability.

### Preflight Summary (Before Sending)
- Detected columns: email, cc (n), bcc (n), total recipient rows.
- Duplicate non-cc/bcc headers list (if any) with chosen column index.
- First N (e.g., 5) sample placeholder keys detected.

### Acceptance Criteria
- Header variants map correctly; tests with all listed email aliases pass.
- Duplicate cc/bcc columns all contribute addresses.
- A CC/BCC cell with `"Alice <a@x.com>, bob@y.com; carl@z.com"` parses into three valid addresses.
- Primary `email` cell `"Jane <jane@x.com>"` is accepted; `"a@x.com, b@y.com"` yields a clear error.
- Placeholder collisions (e.g., `{{name}}` and `{{first name}}`) replace correctly without partial overlaps.
- Errors show sheet name and 1-based row.
- Logs contain no full data dumps.
- Large sheets (e.g., 1000 rows) send in batches and can resume after timeouts without duplicates.

### Running Notes
- 2025-09-06: Implemented utilities (header normalization, address parsing, safe templating), updated `getSheetData` to use display values and SSOT detection, added BCC support and robust CC parsing, enhanced errors with sheet/row context, added batching/resume with `PropertiesService`, centralized sidebar asset resolver. No site changes.

### Implementation Plan (High-Level)
1) Utilities module:
   - normalizeHeaderKey(key): string
   - isAliasOfEmail/CC/BCC(key): boolean
   - parseAddressList(cellValue): string[]  // split, trim, extract from `<...>`, validate, dedupe
   - parsePrimaryAddress(cellValue): string // single address or error
   - buildSafeReplacer(keys): (text, row) => string
2) Integrate utilities into `getSheetData()` and send flows; add BCC support.
3) Replace raw logs with summaries; add detailed per-row error context.
4) Switch to `getDisplayValues()` for templating.
5) Add chunking/resume via `PropertiesService`.
6) Add centralized UI asset resolver.

### Risks & Mitigations
- Overly strict validation could reject valid addresses → Use moderate rules and unit-like test cases in comments.
- Execution time for parsing on large sheets → Keep utilities simple and O(n) per row; chunk work.
- State consistency during resume → Store both last processed row and a simple checksum (e.g., header hash) to detect sheet changes.

### Open Questions
- Should `to` be treated as an alias for email? (Default: no; explicit is better. Provide a helpful error if detected.)
- Should we offer a “dry run” that only produces the preflight summary? (Nice-to-have.)


