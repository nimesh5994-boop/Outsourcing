# Working Paper Platform: Requirements and Desired Outcome

Status: **DRAFT for review.** Nothing here is final until the open decisions in section 10 are answered.
Items marked **[proposed]** are my suggestion, not something already agreed. Items marked **[built]** exist and are tested today; **[partial]** is built but incomplete; **[not built]** is a gap.

---

## 1. Purpose

Year-end working papers are assembled by hand: export reports from the client's Xero, paste them into a firm template, rebuild the schedules, tick and tie everything, then review. Most of that time is assembly, not judgement.

This platform exists to **move preparer time from assembling to reviewing**: connect a client's Xero once, pull the data, and get back a working paper in the firm's own format where every figure traces to source, the arithmetic is already done, and anything that needs a human decision is flagged rather than buried.

## 2. Desired outcome

When this platform is working as intended, for a UK limited company on Xero:

1. A preparer connects the client's Xero organisation once, then pulls a year's data with a click. No export-and-upload.
2. A complete working paper downloads in the practice's own template layout, built from live data.
3. **Every number ties to source.** The ETB equals Xero's trial balance plus posted journals to the penny; every lead schedule agrees to it; every difference is shown, not hidden.
4. **The paper arrives pre-reviewed for exceptions**: reconciliation differences, unusual postings, missing evidence, going-concern and statutory-deadline flags are already raised, each with enough detail to act on.
5. The preparer's time goes on judgement: posting adjusting journals, explaining flagged items, completing the manual sections. Posted journals flow through every schedule automatically.
6. A reviewer can see what was pulled, when, from which organisation, what was overridden by an upload, and what a flag is based on.
7. A second practice could be onboarded without code changes and without being able to see anyone else's data.

### Success measures **[proposed, targets to be agreed]**

| Measure | Proposed target | How measured |
|---|---|---|
| Numerical accuracy | 100% of lead-schedule totals agree to source TB (zero unexplained differences) | Automated regression against real client files, run on every change |
| Preparer time to a reviewable paper | Material reduction vs current manual process | Time a sample of real jobs before/after (baseline needed first) |
| Share of the standard template's tabs produced automatically | Agree a list of "must-have" tabs; target 100% of those | Tab-by-tab coverage table (section 5, FR-W) |
| Exceptions caught pre-review | Reviewer finds nothing the platform could have flagged | Reviewer feedback log |
| Pull reliability | A pull completes first time on a normal-size organisation | Pull success/failure counts |

No baseline timings exist today, so the time-saving target cannot be set honestly until a few real jobs are timed manually.

## 3. Users and roles

| Role | Does | Status |
|---|---|---|
| Partner | Everything: users, templates, clients, jobs, Xero connections | [built] |
| Manager | As partner, minus user management | [built] |
| Preparer | Only clients explicitly granted to them | [built] |
| Reviewer (sign-off) | Review and sign off a generated paper | [not built], see FR-R |
| Client (read-only portal) | See and respond to their own queries | [not built], out of scope for now |

Each user belongs to exactly one practice. No cross-practice access exists.

## 4. Scope

**In scope**
- UK limited companies, year-end accounts and tax working papers.
- Xero as the primary source (direct connection), with file upload as a permanent fallback and override.
- Output as an Excel working paper in the practice's own template (or a system default layout).

**Out of scope for now** **[proposed]**
- QuickBooks / Sage direct connections (upload via generic mapping works but is unvalidated).
- Producing statutory accounts or filing anything with HMRC / Companies House.
- Client-facing portal.
- Payroll processing (a structured BrightPay summary upload is supported for reconciliation only).
- Capital allowances computation and disallowable-expense identification (manual inputs by design).

## 5. Functional requirements

### FR-D: Data acquisition

| ID | Requirement | Status |
|---|---|---|
| FR-D1 | Connect a client's Xero organisation once via OAuth; credentials stored encrypted; reconnect prompt when access is revoked | [built], untested against a live Xero org |
| FR-D2 | Pull Trial Balance (current and comparative), Xero's account types, and the organisation's financial year end | [built], untested against a live org |
| FR-D3 | Warn when the job period cannot be what a year-to-date TB represents (part year, wrong year end) | [built] |
| FR-D4 | Pull the full ledger (every posting) where the Xero app is entitled to it | [partial]: built, but Xero gates it to its top developer tier |
| FR-D5 | Pull aged debtors/creditors per contact as at period end | [not built] |
| FR-D6 | Attach contact names to ledger postings (needed by duplicate, contact-coding and DLA checks) | [not built] |
| FR-D7 | Any uploaded export for a report overrides the pulled version; the preparer can always correct by uploading | [built] |
| FR-D8 | Accept Xero exports by upload with automatic report-type and period detection, bulk and PDF-aware | [built] |
| FR-D9 | Keep the raw Xero response for every pull, downloadable, for tracing a figure | [built] |
| FR-D10 | Pull VAT return detail and bank statement balances | [not built], VAT is not exposed by Xero's UK API; bank statements come from the bank |

### FR-W: Working paper content

The reference is a real finished paper for a real client, used as the answer key. Requirement: **generate the sections below from source data, generically, for any client**, with live formulas so posted journals flow through.

| Section | Status |
|---|---|
| Raw data tabs (TB current/comparative, nominal ledger) | [built] as hidden data sheets feeding every schedule |
| Extended trial balance: Xero PTB + journals, Dr/Cr kept separate, comparative for reference | [built], traced from the real file |
| Journals (narrated adjusting journals by account code) | [built] |
| Profit and Loss, Balance Sheet | [built] |
| Bank lead schedule | [built] |
| Prepayments and accruals (per line) | [built] |
| Directors' loan account balance and activity | [built] |
| Creditors over one year (any non-current liability) | [built] |
| Stock (total level; category split is manual) | [built] |
| Intercompany debtor/creditor (via a per-client related-party list) | [built], needs aged data |
| Control accounts: debtors, creditors, VAT, PAYE, CT, others | [built] |
| Fixed assets: category rollforward and asset-level register | [built] |
| Corporation tax computation | [built], capital allowances manual |
| Debtors/creditors reconciliation to aged listings | [built] (needs aged upload) |
| VAT reconciliation (return vs ledger) | [built] (needs VAT uploads) |
| PAYE reconciliation (payroll vs ledger) | [built] (needs BrightPay uploads) |
| Compliance checklist and points forward | [built] |
| Adjusted position matching the firm's own posted journals exactly | [not built], see FR-W2 |

- **FR-W1** The output uses the practice's own template (insert schedules at configured positions, numbering, header cells, thresholds) or a system default. **[built]** Known cost: round-tripping a real template through openpyxl drops embedded images and dropdown lists. **[open, decision needed]**
- **FR-W2** Firms often post journals using an internal "tax calc code" scheme that differs from Xero's own account codes. To reproduce an adjusted position from the firm's journals, a per-practice code translation is needed. **[not built, proposed as a later phase]**
- **FR-W3** A section is only produced when the client has the relevant accounts (no empty stock tab for a client with no stock). **[built]**

### FR-C: Checks and exceptions

All **[built]**: TB self-balance and tie-out (opening + movement = closing), year-on-year variance and P&L variance with reviewer comment column, nominal activity review and suggested reallocation, duplicate transactions, contact coding consistency, weekend manual journals, directors' loan (S455 / monthly threshold), dividend vs reserves, petty cash, loan facilities, stock presence, going-concern indicators, related-party transactions, statutory filing deadlines, possible miscoded capital expenditure, control-account misposting scan.

- **FR-C1** Every check returns a status (ok / review / n/a / error) with a plain-language message and the underlying detail. **[built]**
- **FR-C2** Materiality and variance thresholds are configurable per template. **[built]**
- **FR-C3** Optional AI-assisted notes on flagged items, opt-in per template, never changing a number, failing safe to no note. **[built]**

### FR-R: Review and sign-off **[not built, proposed]**

| ID | Requirement |
|---|---|
| FR-R1 | Per-schedule status: not started / prepared / reviewed, with who and when |
| FR-R2 | Reviewer comments and preparer responses recorded against a schedule or flag |
| FR-R3 | Record of what was pulled, uploaded and overridden, and when, attached to each generated version |
| FR-R4 | Regenerating after new journals or a re-pull keeps a version history rather than overwriting silently |

### FR-A: Administration

- Practice, user, role and per-client access management **[built]**.
- Template upload and configuration per practice **[built]**.
- Per-client settings (template, related parties, report notes) **[built]**.
- Usage and error visibility for the platform owner (pull failures, generation failures) **[not built, proposed]**.

## 6. Non-functional requirements

| Area | Requirement | Status |
|---|---|---|
| **Accuracy** | Figures must agree to source; a regression suite using real client files and real finished-paper figures runs before release | [built] for one real client; **[proposed]** extend to several clients of different shapes (with stock, loans, VAT, payroll) |
| **Traceability** | Each schedule figure is a live formula back to a data sheet, never a pasted value | [built] for core schedules; reconciliation sheets still hold computed values [partial] |
| **Security: credentials** | OAuth tokens encrypted at rest; single-use refresh tokens saved before use; revocation handled | [built] |
| **Security: isolation** | A practice can never see another practice's data; preparers only see granted clients | [built] |
| **Data protection** | Written data-retention and deletion policy for client financial data; deletion of a client/job removes pulled data and credentials | [partial]: deletion cascades work; no written policy or retention schedule |
| **Auditability** | Who connected which organisation, who pulled, who generated, when | [partial]: connection and pull timestamps stored; no audit log |
| **Performance** | A normal-size organisation pulls and generates without timing out on serverless hosting; large ledgers pulled in resumable batches | [built] for batching; large-organisation timings untested |
| **Reliability** | Transient database errors self-heal; Xero rate limits respected with backoff | [built] |
| **Honest degradation** | When Xero refuses an endpoint or data is missing, the paper still builds and says which schedules lack data | [built] |
| **Testing** | Every change runs automated tests including formula evaluation with a real calculation engine; release gated on green | [built] (~445 automated tests) |

## 7. Constraints and dependencies

1. **Xero has no API key.** Access is OAuth 2.0 through an app registered in the Xero developer portal.
2. **Xero's Journals endpoint** (the only full-ledger source) is limited to its top developer tier plus a security assessment and use-case approval. Without it, ledger-based schedules (TB tie-out, DLA activity, nominal matrix, duplicate and contact checks, control-account movements) still need an uploaded Account Transactions export.
3. **Xero connection limits:** 5 connected organisations on the free tier; more requires a paid tier with app certification. This directly limits how many clients can use direct pull, and is the main commercial constraint on the "second practice" outcome.
4. Apps registered on or after 2 March 2026 must use granular scopes only.
5. Hosting: Vercel (serverless, time-limited requests) with Neon Postgres.
6. Pulled-API shapes are built from Xero's documented format and tested against a fake Xero only. **First live pull must be validated on Xero's free Demo Company** before anyone relies on it.

## 8. Current state versus the outcome

| Outcome | Where we are |
|---|---|
| 1. One-click pull | Built for the trial balance; ledger depends on Xero tier; unvalidated against a live org |
| 2. Complete paper in firm template | Built; template round-trip fidelity cost open |
| 3. Every number ties to source | Built and regression-tested on one real client |
| 4. Pre-reviewed exceptions | Built; some checks degrade without ledger/contact data |
| 5. Posted journals flow through | Built, for Xero-native codes; firm code translation not built |
| 6. Reviewer visibility | Partial: pull record and raw reports exist; no sign-off workflow |
| 7. Second practice onboarding | Data isolation built; no billing, no onboarding flow, Xero connection cap |

## 9. Proposed phasing **[proposed]**

1. **Prove the pull (days).** Run a live pull on Xero's Demo Company; fix any mismatch between real responses and the converters. Decide whether Xero's top tier for Journals is worth pursuing, or whether ledger upload stays the standing route.
2. **Close the data gaps (weeks).** Aged debtors/creditors by contact; contact names on postings (needs an invoices scope); validate on two or three more real clients with different shapes.
3. **Review workflow (weeks).** Sign-off states, comments, version history, audit log (FR-R).
4. **Firm-specific fidelity.** Tax-calc code translation (FR-W2); decide the template-injection approach so branding and dropdowns survive.
5. **Productise (only if the commercial decision is made).** Billing, self-service onboarding, written data-protection policy, Xero tier and certification, support tooling. Keep tenant boundaries as they are; do not build multi-tenant features ahead of the decision.

## 10. Open decisions for you

1. **Who is this for**: Affinity's own outsourced work only, or eventually other firms? This decides whether phase 5 matters and whether Xero's connection cap and app certification are blockers.
2. **Journals access**: pursue Xero's top developer tier (cost, security assessment, approval), or accept uploading the Account Transactions export as the ledger source indefinitely?
3. **Must-have tab list**: which tabs of your standard working paper must be produced automatically for the platform to count as done? (Section 5 lists what exists; I assumed the real file's highlighted tabs.)
4. **Review workflow**: is formal sign-off inside the platform required, or does review stay in Excel?
5. **Template fidelity**: accept losing embedded images/dropdowns when injecting into a firm template, or invest in lower-level template injection?
6. **Firm journal codes**: should journals posted in your own tax-calc code scheme be supported (translation table per practice)?
7. **Baseline**: can a few real jobs be timed manually so the time-saving target is measurable rather than assumed?
8. **Data protection**: who owns the retention and deletion policy for client financial data held here?
