# Working Paper Automation

Turns the bookkeeping exports you already collect from clients (trial
balance, nominal activity/GL detail, aged debtors & creditors, VAT return,
bank closing statement, P&L, B/S) into a prefilled Excel working paper pack
with reconciliations already run: TB self-balance and P&L/B&S tie-out,
current-vs-comparative variance analysis flagged by materiality, debtors and
creditors control account reconciliation (with the full aged listing as a
breakdown of the closing balance), bank and VAT cross-checks, control account
rollforwards (balance b/fwd + movements = balance c/fwd), a nominal
activity analysis matrix that allocates transactions to their contra nominal
code and flags anything that needs manual reallocation, a UK Corporation
Tax computation (current rates, marginal relief applied automatically,
checked against whatever's booked as the tax charge), and a fixed asset
register (a category-level cost/depreciation rollforward derived entirely
from TB + nominal activity, plus an asset-level version that rolls a
prior-year register forward, depreciates each asset, and flags new
additions/possible disposals from nominal activity).

A small internal web app: staff upload a client's export files for a job,
confirm (or correct) how columns map to the working paper's fields, and
generate a downloadable `.xlsx` pack.

## Practices, templates, and clients

The data model is **Practice → Template(s) → Client(s) → Job(s)**: a
practice uploads its own working paper template(s) (`.xlsx`), sets one as
the default, and every client created under that practice inherits it
(overridable per client). Each client is then processed job-by-job,
period-by-period, building up history over time.

## Access control

Every user belongs to exactly one practice, with one of three roles:

- **Partner** - full access: manage users, templates, and every client/job
  in the practice.
- **Manager** - the same client/job/template access as a Partner, minus
  user management (can't add, remove, or re-scope other users).
- **Preparer** - only the specific clients a Partner has explicitly granted
  via **Manage Users → client access**; no template or user management.
  Un-granted clients return a 403, not a 404, so a preparer can tell "not
  mine" from "doesn't exist."

There's no cross-practice access at all - a user from one practice gets a
404 (not a 403) on another practice's pages, so practices can't even be
distinguished from "doesn't exist" by someone outside them.

**Signing up** (`/practices`, logged out) creates a new practice and its
first user (a Partner) together - that's the only way a Partner account is
ever created; every other user is added from that Partner's **Manage
Users** page (name, email, a password the Partner sets and shares
out-of-band, and a role - client access for Preparers is set from the same
page afterward).

Sessions are a signed cookie (`app/auth.py`, via `itsdangerous`), not a
server-side session table - deliberately stateless so there's nothing to
clean up on a serverless platform. Verifying the signature requires
**`SECRET_KEY`**, a required environment variable alongside `DATABASE_URL`
(see Running it / Deployment below) - generate one with
`python3 -c "import secrets; print(secrets.token_hex(32))"` and never
commit it to git. Passwords are hashed with `bcrypt`; nothing is ever
compared or stored as plaintext.

The UI splits **Setup** (`/practices/{id}` - templates and their config,
where a practice admin works) from **Clients** (`/practices/{id}/clients` -
the day-to-day job workflow), since they're different jobs done by
different people at different cadences.

A template is normalised once, at upload, not on every job: loading and
re-saving it through openpyxl reduces its stored size substantially
(confirmed on a real 63-sheet template: 24MB down to 11MB, cell data/
formulas/formatting all intact) so every later generation off that
template is fast. The trade-off - openpyxl round-tripping strips embedded
images and dropdown data-validation lists - is paid once per template, not
once per job; the practice re-adds a logo/validations to the template
after upload if it needs them, not every time a job runs.

Each template carries a JSON **customisation config** (edited from its
detail page) controlling which schedules get generated for that template,
where each one should be inserted, the header-cell convention used for the
CLIENT NAME/PERIOD/SCHEDULE TITLE block, and materiality thresholds. This
is the seam for supporting more than one template format without new code -
adding a second template is configuration, not a rebuild.

**Generating into the real template file.** When a client has a template
set, `generate` builds the working paper pack as new sheets inserted
directly into a copy of that practice's actual uploaded `.xlsx`
(`app/excel_builder.py: build_workbook_into_template`) - not the system's
own generic layout. The template's own sheets (cover page, firm notes,
whatever else is already in the file) are never modified: every generated
schedule is a brand new sheet, positioned per that schedule's
`insert_after_sheet` config (or left wherever it's created if unset), with
numbering starting from `numbering.start_at`. A schedule with
`"enabled": false` is skipped entirely - no sheet, no index entry. A
client with no template configured still gets the system's own generic
layout (`build_workbook`) exactly as before. If a generated sheet's name
collides with one already in the template (e.g. the template already has
its own "Index"), openpyxl auto-suffixes the new one rather than erroring
or overwriting anything.

`header_cells` controls which cell each generated schedule's CLIENT NAME/
PERIOD/SCHEDULE TITLE header block lands in - defaults to A1/A2/A3 (the
generic layout's convention) but is overridable per template, since real
templates vary: studying three of the practice's own real working paper
formats (Xero Ltd, Sage/QBO/FreeAgent Ltd, Manual Job Partnership/Sole
Trader) found a shared "Client"/"Year End"/"Subject" labelled-row
convention with values in a different column, and a newer template found
since uses yet another convention ("Name of company:"/"Start period:"/
"End period:" key-value pairs). `header_cells` configures *where the
value goes* (a cell reference per field); it doesn't write label text of
its own - a template's own surrounding labels, if any, are part of the
template's untouched sheets, not something a generated schedule adds.

**Configurable materiality.** `materiality` (`{"default_amount": 500,
"variance_pct_threshold": 0.10}` by default) sets the £ and % thresholds
every check flags against - a practice tightens or loosens it per template
rather than living with the system's £500/10% defaults everywhere.
`main.py`'s `_generate_workbook_steps` reads it once per job
(`template["config"]["materiality"]`, falling back to each module's own
constant when a job has no template) and threads the resulting
`materiality`/`variance_pct_threshold` through every check that flags on a
variance: `recon.py` (TB variance analysis, nominal activity review,
debtors/creditors control recon, bank recon, VAT cross-check),
`control_accounts.py`, `nominal_matrix.py`, `financial_statements.py`
(the Balance Sheet's own balance check), `going_concern.py`,
`fixed_assets.py` (both the category and asset-level rollforwards),
`accruals_prepayments.py`, `related_party_transactions.py`, and
`corporation_tax.py`. It reaches the generated workbook too, not just the
Recon summary: the live Excel formulas on the formula-linked TB Lead
Schedule (the variance flag) and the fixed asset category sheet (the cost/
depreciation diff highlight) bake in the same configured value, so editing
a figure on a `DATA_*` sheet after the fact still flags against the
practice's chosen threshold, not a hardcoded one.

## Formula-linked schedules

The generated workbook doesn't just contain Python-computed numbers - every
figure on the TB Lead Schedule, control account rollforwards, P&L, Balance
Sheet, Corporation Tax computation, category-level fixed asset register, and
nominal activity matrix is a live Excel formula, the way a manually-built
working paper is. Change a figure on one of the hidden `DATA_*` sheets and
everything downstream recalculates - a reviewer can trace any number back to
its source by following the formula chain, not just trust a pasted value.

How it works: every job's raw canonical data (TB current/comparative,
nominal activity, aged debtors/creditors, P&L, B/S) is written once onto
hidden `DATA_*` sheets (`app/data_sheets.py`). Every schedule sheet then
references those ranges - mostly via `SUMPRODUCT`-based exact-match and
OR-across-values sums (`app/xlformulas.py`), not `SUMIFS`, because the
verification library this was tested against (see below) mishandles
`SUMIFS` on certain text criteria; `SUMPRODUCT` is standard, equally-safe
Excel syntax with no known downside, so it's used everywhere instead. A few
schedules cross-reference each other directly - the Balance Sheet's
"current year profit" line and the Corporation Tax sheet's "profit per
accounts" line both point straight at the P&L sheet's own NET PROFIT cell,
rather than each holding their own copy of that number.

Python still owns anything that's a genuine shaping/classification decision
rather than arithmetic - which fixed asset category an account belongs to
(parsing an inconsistently-formatted account name), which contra account a
nominal transaction gets bucketed under and which accounts make the "top N
by value" cut on the matrix, which control accounts exist at all. Those
groupings are computed once in Python and then referenced by formula (e.g.
by account code or by a synthetic per-transaction `row_id`), so the
*number* in every cell is live, even where the *structure* of the schedule
isn't something a spreadsheet formula should be deriving.

Because this sandbox can't run LibreOffice to recalculate a workbook, every
formula-linked schedule is verified with the `formulas` Python library (a
real, independent formula evaluator) against the equivalent Python
computation - not just "the string looks like a formula", but "a formula
engine computes the same number" (`tests/test_formulas.py`, dev-only:
`pip install -r requirements-dev.txt`). Two library-specific formula bugs
were found and worked around this way (documented in `xlformulas.py`):
`SUMIFS` mishandling leading-zero-looking text criteria, and wrapping an
OR'd boolean condition in `>0` before multiplying returning all-TRUE
regardless of the underlying condition.

Not yet formula-linked: the reconciliation check sheets (TB self-balance,
debtors/creditors/bank/VAT recon, nominal review), the asset-level fixed
asset register (prior-year rollforward), and the closing fixed asset
register - these still write Python-computed values today.

## Status

This is a working MVP built and validated against real Xero exports for a
UK client, not a finished product. Xero is the only platform with a
dedicated parser today (see below); QBO/Sage/other exports go through a
generic column-mapping path that works but hasn't been validated against
real QBO/Sage export files yet - if you have sample exports from those,
share them and the mapping aliases can be tightened the same way Xero's
were.

## Why Xero gets special handling

Real Xero exports aren't flat tables - a naive "map these columns" approach
silently produces wrong numbers on them. What was true on real client data:

- **Trial Balance** carries the comparative year as a single net column
  (e.g. `31 Dec 2024`) embedded in the same file, not a separate export.
  It also already categorises every account (`Sales`, `Direct Costs`,
  `Overhead`, `Bank`, `Current Asset`, `Current Liability`, `Equity`, ...),
  so P&L and B/S are derived directly from the TB rather than requiring
  separate uploads.
- **Account Transactions** and **Aged Payables/Receivables Detail** are
  grouped reports: a section header row per account/contact, detail rows,
  then a `Total <name>` subtotal row - and that subtotal row uses
  un-evaluated formulas (e.g. `=E9`) in Xero's own export, which read back
  as blank/zero. Totals are computed by summing the detail rows directly
  instead of trusting the subtotal row.
- Some accounts (typically directly-connected bank feeds) have **no
  nominal code at all** in the export - identified by name only. Rows
  aren't dropped for this; the account name is used as a fallback key.
- The `Account Transactions` report doesn't include bank accounts at all
  (Xero puts those in a separate report) - so bank movement analysis isn't
  available yet; bank is covered by the simpler statement-vs-TB balance
  check instead.

See `app/xero_reports.py` for the parsers and the docstrings for the exact
quirks each one works around.

## Uploading: bulk, auto-detected, PDF-aware

Staff don't pick a report type, platform, or period before uploading -
they just drop in every file they have for the job at once (CSV, XLSX, or
PDF, in any order), and the system works out what each one is:

1. **Xero-native detection first** (`app/document_detection.py`,
   `try_xero_native`): each file is actually run through the Xero-specific
   parsers (see above) - if one parses cleanly, that's a structural
   validation, not a guess, so the file auto-confirms immediately with no
   review step, exactly as before this feature existed.
2. **Otherwise, guess and confirm**: report type is scored against every
   report type's column-alias dictionary (reusing `mapping.ALIASES`);
   platform falls back to a light column-name heuristic (mostly "other",
   since only Xero has a dedicated parser); period is guessed from
   whatever date signal is available (Xero-style title-row text, or the
   latest date in a date-shaped column), falling back to "a second upload
   of this same report type is probably last year's comparative" when the
   report type has no per-row dates at all (a TB, an aged listing, a VAT
   return). None of this is trusted blind - each file that isn't a
   Xero-native match is queued for a one-page confirm (report type/
   platform/period + column mapping, all editable) before it's used,
   chained one file after another until every upload from that batch is
   confirmed.
3. **PDF support** (`app/pdf_extraction.py`, via `pdfplumber`): for a
   client who only has a PDF, the largest table found across the PDF's
   pages is extracted into the same shape as a CSV/XLSX upload and flows
   through the identical detection/mapping/confirm path. Deliberately
   narrow: one dominant table per file, text-based PDFs only (no OCR - a
   scanned/image-only PDF raises a clear error asking for a re-export
   instead of silently extracting nothing).
4. **Every sheet of a multi-sheet workbook is read**, not just the first
   one - a VAT return export with separate Summary and Detail tabs (a real
   failure mode this fixed: only the first tab was ever seen) expands into
   one classified sub-upload per sheet through the identical pipeline.
5. **Fixed sections, not a dropdown**: the job page shows one permanent
   section per report type (TB, Aged Debtors, Nominal Activity, VAT
   Return, ...), each with its own upload form (dropping a file there
   sets its report type directly, no guessing needed for that part) and
   an editable instruction note - free text describing how *this client's*
   exports of that type should be read (e.g. "VAT export has a Detail tab
   - use it to reconcile against nominal activity"). Notes persist per
   client (`client["report_notes"]`), reused job after job, and are shown
   right there when uploading.

## What gets checked

| Check | What it does |
|---|---|
| TB self-balance | Debits = credits for both years |
| Balance Sheet balance check | Net Assets vs Total Equity, with the current year's profit/(loss) explicitly bridged in (it isn't closed to retained earnings in the TB itself) - shown on the B/S sheet, not just the Index. Any account whose Account Type this system doesn't recognise as one of the five B/S categories is named explicitly as a likely cause, since it's excluded from every total, not just one side of the check |
| Variance analysis | Every nominal code, current vs comparative, flagged if it moves >10% and >£500 |
| Nominal activity review | Flags suspense postings, round-sum manual journals, and descriptions that read like pending corrections |
| Debtors/creditors control recon | Aged listing total vs TB control account balance, with the full client-wise listing (every customer/supplier, every ageing bucket, exactly as submitted) attached |
| Control account rollforwards | B/fwd + movements = c/fwd for any balance-sheet account with nominal detail, with the aged listing attached as a breakdown of debtors/creditors closing balances specifically - plus the actual postings behind the movement and, where there's no aged listing, a movement-by-contact view - see below |
| Control accounts - possible miscoding | Postings coded to the wrong balance-sheet control account, found by contact identity - see below |
| Nominal analysis matrix | Every transaction against an account allocated to its contra nominal code; multi-way splits and unallocated amounts flagged for manual review - plus what's actually inside the OTHER catch-all, and suggested allocations for unallocated items based on that contact's own history - see below |
| Bank reconciliation | Statement closing balance vs TB |
| VAT cross-check | VAT return boxes vs P&L turnover and VAT control account |
| VAT Reconciliation (Box 1 & 4) | General Ledger matched transaction-by-transaction against the filed return's Sales and Purchases detail - a dedicated workspace, see below |
| PAYE Reconciliation | BrightPay payroll data matched against the General Ledger - Net Pay per employee per month, HMRC PAYE & NI and Pension Contributions as monthly totals - a dedicated workspace, see below |
| Corporation Tax computation | Current UK rates (small profits/marginal relief/main rate) applied to accounting profit, checked against the booked tax charge |
| Statutory filing deadlines | Companies House accounts filing, CT600 filing, and CT payment deadlines - computed purely from the job's period end, no upload needed - see below |
| Fixed asset register (category) | Cost/depreciation rollforward per asset category, derived from TB + nominal activity, checked against the TB |
| Fixed asset register (asset detail) | Prior-year register rolled forward asset-by-asset, new additions/possible disposals flagged from nominal activity, totals checked against TB |
| Accruals & Prepayments schedule | Every Prepayment-typed and accrual-named account, side by side in one table, b/fwd + movement = c/fwd checked against TB - see below |
| Contact coding consistency | A contact whose postings are mostly on one nominal code but a small minority land on a different one - the "BT: 10 postings to Telephone, 2 to Light & Heat" pattern - flags the minority transactions with the likely correct code |
| Duplicate transaction check | Same contact+date+amount posted more than once, or the same reference/invoice number reused on the same nominal code for the same amount - excludes the natural double-entry legs of one transaction (same reference on different codes, or an invoice and its later payment, which share a reference but have opposite signs) |
| Unusual posting date check | Manual journals (not bank feed or trading transactions, which legitimately happen any day) posted on a weekend |

Every check produces a status (`ok` / `review` / `error` / `n/a`) and a
plain-English message, shown on the workbook's Index sheet and again on
each schedule.

The last three (`app/anomaly_detection.py`) are a different kind of check
from the rest of the table: everything else compares a total against
another total (TB vs aged listing, VAT return vs nominal ledger); these
look across *many* transactions for the same contact to find a pattern a
single row can't reveal on its own. They're deliberately conservative -
each requires a decent transaction sample and a genuinely lopsided pattern
before flagging anything, specifically to avoid drowning a reviewer in
noise from transactions that only look similar (see the double-entry
exclusion above, found by running this against real sample data and
noticing every ordinary invoice was getting flagged as its own duplicate).
They surface candidates for a human to confirm, the same as every other
check in this system - not an auto-fixer.

### VAT Reconciliation workspace

The VAT cross-check above compares totals (VAT return boxes vs the P&L
and the VAT control account) - good for a first sanity check, but it
can't tell you *which* transaction is missing when it doesn't tie out.
The VAT Reconciliation workspace (its own card on the job page, `app/
vat_reconciliation.py`) does that at transaction level: it matches the
General Ledger against the filed VAT return's actual Sales (Box 1) and
Purchases (Box 4) detail, one invoice/bill at a time, and reconciles
each box independently - runnable on its own, with its own "Run VAT
Reconciliation" button, before a full working paper is ready to
generate (and it's also folded into the main results list during
Generate, so it shows on the Index sheet and gets its own tabs in the
workbook either way).

**Three upload zones**, each accepting Excel/CSV/PDF and multi-sheet
workbooks (same auto-expansion as the rest of the app - a file with
separate Summary/Detail tabs becomes two uploads automatically):

- **General Ledger** - the QBO/Xero general ledger for the period.
- **Filed Return, Sales Detail (Box 1)** and **Filed Return, Purchases
  Detail (Box 4)** - each accepts up to 10 files together (e.g. one per
  VAT quarter across the year); every file's rows are combined into one
  dataset for that box, with each row kept tagged with the file it came
  from (`Source File`) so a finding can be traced back to a specific
  return.

**Matching** cascades through three passes per box, strongest first:
invoice/reference number alone (deliberately not gated on amount, so a
real invoice with the *wrong* VAT amount posted against it is still
recognised as the same transaction and its variance reported, rather
than showing up as two unrelated "unmatched" rows); then date + amount
(within the configured tolerance) + contact; then amount + contact
alone, flagged `(verify)` since it's the least certain. Box 1 is
matched first and claims its General Ledger rows before Box 4 runs
against what's left, so a sales transaction can never also get claimed
as a purchase; whatever General Ledger activity neither box's filed
return accounts for is reported once, as its own **General Ledger
Coverage** check - not duplicated as a false "exception" under both
boxes just because a sale doesn't appear in the purchases detail (or
vice versa).

Same "read the GL in depth, surface the assumption" treatment applied to
FAR/control accounts/the nominal matrix, translated here:

- **Full matched-pairs detail**: each box's result now attaches every
  matched GL-row/filed-return-row pair - not just the exceptions - as
  `matched_detail` (a new field on `ReconResult`, generic enough that any
  check can use it: extra_detail was already the exceptions table, this
  is the proof behind "142 matched" so it's never just a trusted number).
  Each pair also shows its **implied VAT rate** (VAT ÷ net, from the
  actual GL posting) - when one doesn't correspond to a standard UK rate
  (0%/5%/20%), that's stated in the box's message. Advisory only, same
  discipline as FAR's system-estimated depreciation: never flips a box
  to "review" by itself, since a partial-exemption or flat-rate-scheme
  client can legitimately post other rates.
- **Suggested box for General Ledger Coverage gaps**
  (`_suggest_box_for_gl_coverage_gap`): for a GL transaction that matched neither box, checks that SAME
  contact's other GL transactions that DID match - this client's own
  pattern, not a guess, the same approach as FAR's capex suggestion,
  control accounts' miscoding suggestion, and the nominal matrix's
  unallocated-transaction suggestion. Only suggested when one box
  clearly dominates that contact's matched history (a real majority,
  over a real sample). Runs as its own check ("VAT Recon - suggested box
  for unmatched General Ledger items") with a candidate table; never
  moves or matches anything itself.

**Settings** (saved per job): **VAT Accounting Basis** (Accrual - the
General Ledger's own transaction/invoice date; Cash - a payment date
column when the export actually has one, falling back to the
transaction date row-by-row where it doesn't) picks which date the
date-based matching pass compares against. **VAT Matching Tolerance**
(Exact Match, or a small £ tolerance) sets how close two amounts need
to be to count as the same transaction - and doubles as the threshold
below which a found match's residual difference isn't worth flagging
as a variance in the first place.

This is the first of a small family of these - see PAYE Reconciliation
next, same shape: its own dedicated section, its own settings,
testable on its own before it's relied on.

### PAYE Reconciliation workspace

BrightPay payroll data matched against the General Ledger. Built
against a real client's exports (BrightPay's Payroll Summary, P32, and
Pensions reports, plus the matching Xero Account Transactions/Trial
Balance), which turned up two things worth knowing before reading the
design: this client's books had ~10 days of the year posted against a
BrightPay export covering the full year (the reconciliation correctly
reports that as "no General Ledger match" for almost everything - a
real gap in the books, not a bug), and there is no formal payroll
accrual journal at all - no Gross Wages or NI control account exists;
net pay is paid straight to each employee (contact = employee name,
no reference numbers), and PAYE/NI and pension are each paid to HMRC/
the provider as one lump sum a month. The engine is built around that
reality rather than an assumed textbook chart of accounts:

- **Net Pay** is reconciled **employee-by-employee-by-month**: each
  BrightPay employee/month is matched against a General Ledger posting
  by contact name (fuzzy - first and last word, so BrightPay's full
  legal name "Jamie Francis Doe" still matches a ledger contact of
  "Jamie Doe") and amount, searching the *whole* ledger rather than
  one assumed account, since a bookkeeper coding things differently
  than expected is exactly the scenario this needs to survive.
- **HMRC PAYE & NI** and **Pension Contributions** are reconciled as
  **monthly company-wide totals** - P32's "Amount due" and each
  month's total employee+employer pension contribution, matched
  against a General Ledger posting whose contact looks like HMRC or a
  pension provider (a short built-in keyword list per check, same
  pattern as `vat_control_names` in the checks table above). There's
  no per-employee trace of either in real books, so matching at
  employee level isn't meaningful for these two.

No reference/invoice numbers exist anywhere in this domain - BrightPay
doesn't generate them and neither does a bank payment run - so
matching is a single pass: contact + amount within tolerance, picking
whichever candidate's date is closest when more than one qualifies
(**Date window**, a setting: how many days after the period end to
search - HMRC and pension payments trail the tax month by days to
weeks, never land the same day), and never letting one General Ledger
posting settle two different months.

Two distinct things can put a same-contact, same-amount posting in the
ledger, and only one of them is a real payment: a **cash settlement**
(money that actually left the bank) or an **accrual booking** (an
invoice or bill recorded as owed, with no payment yet). Found by
inspecting a real export directly: an invoice raised in March and paid
in April shows up as two separate ledger rows on two different dates,
and a second invoice with no payment row at all by year end is
genuinely still outstanding. Confusing the two would let an unpaid
invoice masquerade as a paid one, or split one real payment's two
double-entry legs (same contact, same date, equal-and-opposite
debit/credit - Xero lists a payment once per account it touches) into
what looks like two payments. So before matching, the General Ledger
pool is narrowed using Xero's own `Source` column: rows reading
"Payment", "Receivable Payment", "Spend Money" or "Receive Money" are
kept as cash-settlement candidates; rows like "Payable Invoice" or a
manual journal are excluded as accrual-only. What survives that filter
is then deduplicated by (date, contact, amount) to collapse a single
payment's two double-entry legs into one candidate. **This assumption
is never applied silently** - every PAYE Recon result message states
how many postings were excluded as accrual-only, so a preparer can
check the judgement call rather than just trust it. If a General
Ledger upload doesn't carry a `Source` column at all, the system says
so plainly in the message and falls back to treating every same-
contact posting as a candidate, rather than guessing.

An item that doesn't find a within-tolerance match still gets the
*nearest same-contact posting* reported as a diagnostic (ignoring
amount) - "no match" alone is rarely as useful to a preparer as "here's
the closest thing, look at the difference yourself".

**Three uploads**, auto-detected from BrightPay's own column layout
(no mapping step, the same way a genuine Xero export auto-confirms) -
Payroll Summary, P32, and Pensions. **No separate General Ledger
upload**: this workspace reuses the job's existing Nominal Activity /
General Ledger Detail upload rather than asking for the same file
twice. A client's tax year can span more than one BrightPay export
(a new one each BrightPay tax year); upload as many as needed, they're
combined automatically the same way VAT Reconciliation's filed-return
uploads are.

### AI-assisted reconciliation notes (opt-in)

Off by default. When a practice turns it on for a template
(`ai_reconciliation_notes.enabled` in that template's config), every check
flagged `review`/`error` gets one more attempt at being useful: a short
suggestion from an LLM (`app/reconciliation_agent.py`, via the Anthropic
API), shown as a distinctly-styled "AI-ASSISTED NOTE" block on that
check's sheet, right below the deterministic detail tables.

What it's given is exactly what a reviewer already sees on that sheet -
the flagged check's own message and detail table, any `extra_detail` a
check attaches (e.g. the VAT cross-check's candidate nominal postings),
and the practice's own instruction note for that report type (see
Uploading above) - nothing it wasn't already shown elsewhere. It's told
explicitly to hedge, cite only what's in the data it was given, and
respond with a fixed marker if it has nothing useful to add (in which
case no note is written at all - it doesn't pad output for the sake of
it).

This is deliberately kept separate from the deterministic engine:
`recon.py` and friends make no API calls and never will; `ReconResult`
just carries an extra `ai_note` field that main.py's `generate()` fills
in from outside, after all the real checks have already run and produced
their real numbers. A note is a hint for the reviewer, not a finding -
every number in the workbook is exactly what the deterministic checks
computed, with or without this feature on. Any failure (missing API key,
API error, timeout) degrades to no note - or a diagnosable one-line
message in its place - never to a broken generation.

Requires `ANTHROPIC_API_KEY` as an environment variable (see Running it /
Deployment below) only if the feature is actually turned on for a
template; the app runs fully without it otherwise.

## Compliance checklist

Distilled from a real manual-job review checklist covering fixed assets,
bank, stock, debtors, creditors, VAT, PAYE, pensions, loans, DLA, and
dividends - split into the two kinds of item it actually contains:

**Data-driven checks** (`app/compliance_checks.py`) - answerable from data
already ingested, so the system flags them automatically:

| Check | What it does |
|---|---|
| Directors' loan account review | Flags any calendar month where net DLA withdrawals exceed £10,000 (the point HMRC treats it as a loan needing a benefit-in-kind/interest review, not routine drawings), and - separately - flags an S455 consideration with a drafted year-end balance note whenever the account is in debit (the director owes the company) |
| Dividend vs distributable reserves review | Compares dividends declared this year against retained earnings b/fwd + this year's profit; flags a potential unlawful dividend if declared amounts exceed what's available |
| Petty cash running balance review | Rolls the petty cash account's balance forward transaction by transaction through the year; flags any point it goes negative (physically impossible for cash, so it means a mis-dated/mis-posted entry or cash fronted by the business) |
| Loan facility review | Detects Bounce Back Loan / Hire Purchase / Bank Loan-style accounts and lists the specific checklist points that apply to each (BBL's 12-month interest holiday, HP's within/after-one-year split, agreement/statement received) - a reminder, not a computed check, since there's no repayment schedule or agreement in the data this system has |
| Going concern indicators (`app/going_concern.py`) | Negative net current assets (a working-capital deficit) and negative net assets (balance sheet insolvency), from the already-derived Balance Sheet statement - purely an arithmetic fact about the balance sheet shape, explicitly never presented as a verdict on going concern itself, which needs judgement and forward-looking information this system has no way to derive from historical accounting data alone |
| Related party transactions (`app/related_party_transactions.py`) | Postings to a contact already identified as a related party (from Directors' Loan Account activity) landing on some other account - candidates for FRS 102 Section 33 disclosure. Explicitly not a completeness check - only catches a related party who also transacted through the DLA this year, since that's the only source of related-party identity this system has |

**Manual checklist tab** (`app/excel_builder.py: build_compliance_checklist_sheet`,
config key `compliance_checklist`) - a static, editable pro-forma sheet
for everything in that source checklist that *can't* be answered from
data: "was the HP agreement received?", "does the CT liability agree to
the HMRC online account?", "was a bank statement received for every
account?". ~28 items across fixed assets, bank, stock, debtors,
creditors, VAT, PAYE/wages, pensions, loans, DLA, dividends, government
grants, and corporation tax, each with a blank Status and Notes column for
the preparer to complete. Deliberately doesn't repeat anything the
data-driven checks above already cover.

## Live progress during Generate

Generating a working paper runs ten real steps in sequence - loading
uploads, three check modules, the optional AI notes step, control account
rollforwards, the nominal matrix, Corporation Tax, fixed asset registers,
then building the workbook itself (see `GENERATE_STEPS` /
`_generate_workbook_steps()` in `main.py`, the one place this logic
lives). Clicking **Generate working paper** now opens a small progress
page (`job_generate.html`) instead of just waiting on a blank load: it
opens a Server-Sent Events connection to `GET /jobs/{id}/generate/stream`
and lights up each step as it starts and finishes, redirecting to the job
page the moment the workbook is saved. A step that didn't run (AI notes,
when not enabled for that template) shows as skipped rather than stuck.

The original `POST /jobs/{id}/generate` still exists, unchanged in
behaviour - it just exhausts the same generator without looking at the
intermediate events, so nothing that depended on the old classic
request/redirect contract (including the test suite) needed to change.

**Known unverified risk:** SSE depends on the response actually being
streamed to the browser incrementally rather than buffered until the
whole thing finishes. This is confirmed working locally
(`uvicorn`/Starlette stream it correctly), but Vercel's Python runtime's
streaming behaviour hasn't been verified against the real deployment from
this environment - if it turns out to buffer the whole response, the
progress page will still work (it'll just show everything at once, right
before the redirect, rather than live) rather than break.

**Persisted, not just streamed.** Every step event is written onto
`job["progress"]` as it happens (`storage.save_job()` per event), whether or
not anyone is watching the live stream - so the classic POST route persists
the same trail as the SSE route, and a run that nobody watched live is still
reviewable afterwards. The job detail page reads this back through
`_summarize_progress()` and shows a collapsed **Last generation** panel
under the Generate button: per-step status (done/skipped/still-running) and
duration in seconds, a total run time, and - if the run errored out - the
error message in place of a timing. This is the first item from the Pipeline
Map's action plan (`§05`): a foundation for spotting which step is
consistently slow or where a run tends to die, without having to have had
the progress page open when it happened.

**Retries transient failures automatically.** Only steps 1 (loading
confirmed uploads) and 10 (loading the template and saving the workbook +
job) touch Postgres - everything else runs on data already in memory, so a
failure there is deterministic and retrying it would never change the
outcome. Those two steps run through `_run_step_with_retry()`: up to two
retries with a short, growing backoff before the step's error is allowed to
fail the run, same as before this existed. A "retrying" event fires between
attempts (with the attempt number and the error), persisted alongside every
other progress event, so a run that needed a retry still shows up in the
**Last generation** panel with a `retried Nx` badge next to the step,
instead of either failing outright on a one-off blip or silently hiding
that it happened. This is the third item from the Pipeline Map's action
plan (`§05`) - persist and time were combined into one item above.

## Corporation Tax computation

Built from `app/tax_rates.py` (the current rates, as a config - not
hardcoded inside the calculator) and `app/corporation_tax.py` (the standard
small profits rate / marginal relief / main rate rules, verified against
the well-known £150,000 → 24% effective rate reference point). It takes
accounting profit (from the derived P&L) through preparer-entered
adjustments - disallowable expenses and capital allowances, both editable
via a form on the job page since they need a fixed asset register / expense
review this system doesn't have - to taxable profit, applies the current
rate band automatically, and flags a variance against whatever's booked as
Corporation Tax in the trial balance. It's a computation proforma and
reasonableness check, not a full tax return.

**Keeping the rates current**: a scheduled Routine ("HMRC Corporation Tax
rate watch", quarterly, timed around UK Budget/Autumn Statement season)
checks gov.uk for rate/threshold changes and sends a push+email notification
if it finds one - it reports what changed rather than auto-updating the
repo, so a real change gets a human decision before it reaches `tax_rates.py`.
Every generated CT schedule also states the rates used and when they were
last verified, so a stale config is visible on the workbook itself even
between routine checks.

## Statutory filing deadlines

Built from `app/statutory_deadlines.py`: a real job file always states
somewhere when the accounts and Corporation Tax return are due - this
system had nowhere at all for that until now, despite needing nothing
but the period end date already set when the job was created. Computes
three dates for an ordinary UK private limited company:

- **Companies House accounts filing deadline** - normally 9 months after
  the period end, but governed by the Companies Act 2006 s.442
  Accounting Reference Date rule: if the period end is itself the last
  day of its calendar month (e.g. 31 December, or 28 February in a
  non-leap year), the deadline is the *last day* of the target month 9
  months later - regardless of how many days that month actually has -
  not just the same day-of-month. Both cases are handled explicitly
  (`_add_months_ch_style`), since getting this wrong silently would be
  worse than not having the feature at all.
- **Corporation Tax return (CT600) filing deadline** - 12 months after
  the period end, HMRC's plain anniversary date (no ARD-style special
  case).
- **Corporation Tax payment deadline** - 9 months and 1 day after the
  period end - stated explicitly as assuming the company pays by a
  single instalment (augmented profits under £1.5m, scaled down for a
  short period or associated companies), not the quarterly instalment
  regime a large company must use instead.

Always status `ok` (these are facts, not findings), so it never gets in
the way of what actually needs review - just states the three dates on
the Index sheet and its own schedule. Deliberately doesn't attempt a
first set of accounts (its own 21-months-from-incorporation rule) or the
confirmation statement (due from the incorporation/last-statement
anniversary, not the accounting period) - neither is derivable from the
job's period dates alone, and guessing either would be worse than
leaving it out.

## Nominal analysis matrix

Built from `app/nominal_matrix.py`: for a chosen account (Bank, or the
accounts with the most transaction volume by default), every transaction
is allocated to its contra nominal code as a matrix column - the pattern
used in schedules like "Bank Receipt Analysis" / "Purchases & Expenses"
in the real working paper - so a reviewer sees at a glance where the
money went. Requires a Xero Account Transactions export (its "Related
account" field is what makes contra-account allocation possible at all);
generic-mapped uploads won't have this schedule. Contra accounts beyond
the top 10 by value are folded into a single OTHER column so the matrix
stays readable, and any transaction with no contra code at all is
flagged UNALLOCATED.

Two more features, same "read the GL in depth, surface the assumption"
philosophy as FAR and control accounts above:

**What's actually inside OTHER**: every matrix now attaches, as
`extra_detail`, exactly which contra accounts got folded together - each
with its own net amount and transaction count - so OTHER is never a dead
end for a reviewer. When one of those folded-in accounts is itself
individually material (> £500), the result message says so explicitly:
folding into OTHER is purely a display-column cap, not a materiality
judgement, and a large contra account that only missed the top-10 cut
shouldn't read as immaterial. Advisory only - like FAR's system-estimated
depreciation, this never by itself flags the matrix "review".

**Suggested allocations for unallocated transactions**
(`suggest_unallocated_reallocations`): for a transaction with no contra
code at all, checks that SAME contact's other, already-allocated
transactions on the same account - this client's own transaction
history, not a generic guess, the same "client's own data as vocabulary"
approach as FAR's capex suggestion and control accounts' miscoding
suggestion. Only suggested when one contra account clearly dominates
that contact's history (a real majority, over a real sample) - otherwise
left as a genuine unallocated item with no guess attached. Runs as its
own check across every matrix account with a candidate table of exactly
what's suggested and why; never allocates anything itself.

## Control account rollforwards

Built from `app/control_accounts.py`: for every balance-sheet account
with nominal-ledger detail, balance b/fwd + movements during the year
should equal balance c/fwd per the current year TB - any difference is
exactly what needs journalling or investigating. Debtors/creditors
control accounts also carry the full aged listing as a breakdown of the
closing balance, so a reader sees what the balance is actually made up
of, not just its total. Covers every standard Xero balance-sheet account
type with a rollforward that makes sense generically - Current Asset,
Current Liability, Liability, **Prepayment, Inventory, Non-current
Asset, Non-current Liability** - so a prepayment schedule, a stock
account, or a long-term loan each get the same rollforward as a trade
debtors control account, not just the handful of types this originally
shipped with. Fixed Asset accounts are deliberately excluded here - they
get their own, more sophisticated treatment in `fixed_assets.py` (see
below).

Two features - the same "read the GL in depth, surface the assumption"
philosophy applied to FAR above - make this more robust:

**The actual postings, not just the movement total**: every rollforward
now attaches the real nominal-activity transactions behind its single
"MOVEMENTS DURING YEAR" figure as `extra_detail`, so a preparer can check
real postings instead of trusting an aggregate.

**Movement-by-contact, for control accounts with no aged listing**: only
debtors/creditors get the authoritative closing-balance-by-party
breakdown (there's no independent aged listing for a Directors Loan or
Wages Control account). Every *other* control account now gets a
different, honestly-scoped view instead: the year's net movement grouped
by contact - explicitly labelled as the movement only, not the closing
balance make-up (there's no opening-balance-by-contact data to add to
it), and never checked against the TB the way the aged-listing breakdown
is, since it isn't a claim about the full balance.

**Possible postings coded to the wrong control account**
(`suggest_control_account_miscoding`): scans nominal activity for a
contact known to one control account (from its own postings this year,
plus - for debtors/creditors - the aged listing itself, independent
ground truth that doesn't depend on this year's postings at all) turning
up on a *different* balance-sheet control-account-shaped account above
materiality - e.g. a director's loan repayment coded to "Other
Creditors" instead of "Directors Loan Account". Deliberately never looks
at Bank or P&L accounts: in a double-entry export, a genuine transaction
already shows the same contact under at least one other account (an
invoice's contra is Sales, a receipt's contra is Bank) - that's normal,
not a miscoding, and flagging it would just be noise on every correctly-
coded posting. Runs as its own check with a candidate table of exactly
what matched and why; never reclassifies anything itself.

## Fixed asset register

Built from `app/fixed_assets.py`, in two parts that work independently,
plus two features that make both of them more robust by reading the
General Ledger in depth rather than trusting the TB's own totals:

**Category-level** (always available, no upload needed): real charts of
accounts commonly split each fixed asset category into a cost/additions
code and a depreciation code (naming varies a lot in practice - some use
"COMPUTER EQUIPMENT - COST", others run words together with no separator
at all, e.g. "IT EQUIPMENT COST BROUGHT FORWARD" - both are handled by
stripping structural words like cost/additions/depreciation/brought/
forward wherever they appear, rather than matching a fixed suffix
pattern). Paired-up cost and depreciation accounts become a cost /
accumulated depreciation / NBV rollforward per category, cross-checked
against the TB, brought forward from the comparative TB.

**Asset-level** (needs a prior-year register upload - report type "Fixed
Asset Register", generic column mapping since there's no standard
software export for this): rolls each asset forward using its own
depreciation method (straight line or reducing balance) and rate,
flags nominal-activity transactions against fixed asset cost codes that
aren't yet in the register as candidate new additions, flags credit
movements on those codes as possible disposals to match and remove, and
reconciles the register's total NBV to the TB.

**Additions, shown as real transactions, not just a total**
(`far_additions_detail`): every debit posted during the year to a
category's cost code, listed line-by-line - date, reference, description,
contact, amount - attached below the category rollforward as its
`extra_detail` (both in the fallback sheet and the formula-linked one).
This is the first, no-inference step of making the register more robust:
anything the client's own chart of accounts already says is a fixed
asset shows up as a checkable addition, not just an aggregate a preparer
has to trust.

**System-estimated depreciation** (`DEFAULT_CATEGORY_DEPRECIATION`,
`_infer_category_rate`): each category's brought-forward cost/NBV is run
through a default depreciation rate and method inferred from the
category's own name (computer/IT equipment 25% reducing balance, motor
vehicles 25% reducing balance, fixtures/fittings/furniture 15% straight
line, plant/machinery/equipment 20% reducing balance, leasehold
improvements 10% straight line, buildings 2% straight line, 20% straight
line otherwise), prorated for a period shorter than a full year. Shown as
its own "System est." columns next to what's actually booked, and stated
in the result message every time - deliberately never a reason to flag a
category "review" by itself, since real accounting policies vary far too
much for a keyword-inferred rate to be authoritative. It's a sanity-check
number to look at, not a finding.

**Capital expenditure coded elsewhere**
(`suggest_capital_expenditure_reclassification`): the other direction -
scans nominal activity coded to an ordinary expense account (Xero's
Overhead/Direct Costs/Expense types) for postings above the £500
materiality threshold that look like capital expenditure someone
miscoded, e.g. a laptop put through "IT Costs" instead of "Computer
Equipment", or a van through "Repairs and Maintenance" instead of "Motor
Vehicles". The vocabulary it matches against is built from *this
client's own* fixed asset categories - both the TB's cost/depreciation
account names and, if uploaded, the prior year register's own category
column - topped up with a short list of generic capex nouns (laptop,
van, machinery, furniture, renovation, and similar). This is "nature of
business" read from what the client actually owns already, not a
generic industry guess: a garage's existing "Workshop Equipment"
category is what catches its own miscoded tool purchases; a different
client's categories catch different things. Runs as its own check
("Fixed asset register - possible capital expenditure coded elsewhere")
with a candidate table of exactly what matched and why - never
reclassifies anything itself, since a keyword match on a description is
a starting point for the preparer to check, not proof of capital nature.

## Accruals & Prepayments schedule

Built from `app/accruals_prepayments.py` - a section every real working
paper file has, that this system had no equivalent of at all until now.
Widening `control_accounts.py`'s account-type coverage to include
Prepayment (see "Control account rollforwards" above) already gives a
prepayment account its own individual rollforward tab; this is the
other view a preparer also expects - every prepayment and accrual line
side by side in **one consolidated table**, not scattered one-tab-per-
account, the same two-views relationship fixed_assets.py already has
between its category summary and its asset-level detail.

Two categories, found differently since Xero doesn't have a distinct
"Accrual" account type the way it has a genuine "Prepayment" one:
**Prepayments** are TB accounts of that exact Xero account type - no
ambiguity, the platform already says what these are. **Accruals** are
Current Liability accounts whose name reads like one (specific keyword
phrases - "accrued expenses", "accruals", and similar - the same
reasoning as `control_accounts.py`'s debtor/creditor keyword matching:
a loose "accrue" substring would also catch "Accrued Corporation Tax",
which already has its own dedicated CT schedule and is deliberately
excluded here to avoid double-counting the same balance under two
schedules).

Each line rolls forward - balance b/fwd + movement = balance c/fwd,
checked against the TB, flagged if it doesn't tie - and the actual
nominal-activity postings behind every line's movement are attached as
`extra_detail`, the same "real transactions, not just a trusted total"
treatment as the rest of this session's robustness work.

## Upload safety

Two checks run automatically on every upload, so a wrong file doesn't
silently produce a wrong working paper:

- **Report-type confirmation**: if a file's columns don't look like the
  report type selected (e.g. an aged debtors export uploaded as a VAT
  return), a warning is raised before you map anything.
- **Period check**: a job is created with real start/end dates, not just a
  label. Xero exports state their own period in the title rows (`As at 31
  December 2025`, `For the period 1 January 2025 to 31 December 2025`) -
  extracted and compared against the job's declared period, flagging a
  mismatch if someone uploads the wrong year/quarter.

## Reusable client mapping profiles

For non-Xero (generic) uploads, once you confirm how a file's columns map
to the working paper's fields for a client, that mapping is saved
(`data/clients/<id>/mapping_profiles/`) and auto-applied next time you
upload the same report type/platform for that client - so recurring
monthly/quarterly/annual work doesn't mean remapping columns every time.

## Architecture

```
app/
  models.py           canonical schemas every report type is normalised into
  mapping.py           column-alias suggestions + per-client mapping profiles (generic/non-Xero path)
  parsers.py            generic CSV/XLSX loading + column mapping; DataSource abstraction
  xero_reports.py        Xero-specific report parsers (see above)
  document_detection.py   auto-detects report type/platform/period for an upload
  pdf_extraction.py        pdfplumber-based table extraction for PDF uploads
  recon.py                 the reconciliation/cross-check engine
  vat_reconciliation.py     VAT Reconciliation workspace - General Ledger vs Filed VAT
                             Return, Box 1/Box 4 transaction matching (see above)
  brightpay_reports.py       native BrightPay report parsers (Payroll Summary, P32, Pensions)
  paye_reconciliation.py      PAYE Reconciliation workspace - BrightPay payroll data vs
                               the General Ledger (see above)
  reconciliation_agent.py    opt-in LLM explanation for flagged checks (no API calls from recon.py itself)
  anomaly_detection.py       cross-transaction checks (miscoding, duplicates, unusual posting dates)
  compliance_checks.py        data-driven checklist checks (DLA/S455, dividends, petty cash, loans)
  going_concern.py             negative net current assets / negative net assets indicators
  related_party_transactions.py    DLA-contact postings landing outside the DLA account (FRS 102 s.33 candidates)
  control_accounts.py       control account rollforward + aged breakdown engine
  nominal_matrix.py          nominal activity → contra nominal code analysis matrix (+ formula row-id grouping)
  fixed_assets.py             category-level + asset-level fixed asset register engine
  financial_statements.py      structured P&L / Balance Sheet with the explicit balance check
  corporation_tax.py            UK Corporation Tax computation (marginal relief etc.)
  tax_rates.py                   the current CT rates config, monitored for HMRC changes
  statutory_deadlines.py           Companies House / HMRC filing deadline calculator (period dates only)
  accruals_prepayments.py           consolidated Accruals & Prepayments schedule (one table, not one tab per account)
  data_sheets.py                  writes hidden DATA_* raw-data sheets for the formula engine
  xlformulas.py                    Excel formula-string builders (SUMPRODUCT-based, see above)
  excel_builder.py                  builds the final .xlsx: house-style headers, numbered index,
                                     value-based AND formula-linked schedule builders
  storage.py                         Postgres-backed Practice/Template/Client/Job storage
                                       (entities/files/mapping_profiles/users/client_access tables)
  auth.py                             password hashing, signed-cookie sessions, role/practice/
                                       client authorization helpers
  main.py                             FastAPI app + routes
  templates/, static/                  the (minimal) web UI
api/
  index.py                             Vercel entrypoint - re-exports app.main:app as an ASGI callable
vercel.json                            Vercel build/route config (Python runtime, api/index.py)
```

`DataSource` in `parsers.py` is the seam for live API connectors later
(Xero/QBO/Sage OAuth pulls instead of file uploads) - today the only
implementation is `FileDataSource`, but nothing else in the pipeline would
need to change to add one, since everything downstream just consumes a
canonical DataFrame.

## Running it

Storage is Postgres-backed (see Architecture), so a `DATABASE_URL` is
required even for local runs - point it at a local Postgres or a free Neon
branch. Tables are created automatically on first connection (see
`SCHEMA_STATEMENTS` in `storage.py`); there's no separate migration step.
`SECRET_KEY` is also required, to sign session cookies (see Access
control above). `ANTHROPIC_API_KEY` is only needed if you turn on
AI-assisted reconciliation notes for a template (see What gets checked
above) - the app runs fine without it otherwise.

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
export DATABASE_URL="postgresql://user:password@host:5432/dbname"
export SECRET_KEY="$(python3 -c 'import secrets; print(secrets.token_hex(32))')"
uvicorn app.main:app --reload
```

Then open http://127.0.0.1:8000 - create a client, start a job with the
period dates, upload the reports you have, confirm any mapping that needs
review, and generate.

## Deployment (Vercel + Neon)

The app runs as a single FastAPI ASGI app behind Vercel's Python runtime:
`api/index.py` re-exports `app.main:app`, and `vercel.json` routes every
request to it. Because serverless invocations don't share a filesystem,
all state (practices, templates, clients, jobs, uploaded/generated files)
lives in Postgres rather than on disk - see `storage.py`.

To deploy:

1. Create a Postgres database (a [Neon](https://neon.tech) project works
   well - serverless, scales to zero, has a generous free tier). Use a
   pooled connection string (Neon's `-pooler` host) since each serverless
   invocation opens its own connection.
2. In the Vercel project settings, set the environment variables
   `DATABASE_URL` (that connection string) and `SECRET_KEY` (a random
   value, e.g. `python3 -c "import secrets; print(secrets.token_hex(32))"`
   - used to sign session cookies, see Access control above). Tick
   Production for both, at minimum. Never commit either to git. Add
   `ANTHROPIC_API_KEY` too if any practice will turn on AI-assisted
   reconciliation notes - otherwise skip it, nothing needs it.
3. `vercel deploy` (or connect the repo in the Vercel dashboard for
   git-push deploys). No other build step is needed - `requirements.txt`
   is installed automatically by the Python runtime.

On first request after deploy, `storage.py` creates its tables
(`entities`, `files`, `mapping_profiles`) if they don't already exist, so
there's nothing else to provision.

**Known constraint:** uploaded/template/generated files are stored as
Postgres `BYTEA` rather than in a separate object store, to keep the
credential footprint to just `DATABASE_URL`. This is simplest-thing-that-
works for typical working-paper file sizes (tens of KB to a few MB), but
two limits are worth knowing about: Vercel serverless functions cap
response payload size (currently 4.5MB on the default plan), and very
large template/output workbooks will hit that before Postgres itself
becomes a problem. If that happens, move `save_file`/`load_file` in
`storage.py` to an object store (Vercel Blob or S3) - callers already
just pass/receive bytes, so the swap is contained to those two functions.

## Tests

```bash
pytest tests/ -v
```

`tests/test_pipeline.py` runs the full pipeline (parse → reconcile → build
workbook) against `sample_data/`, which mirrors the real Xero export
structures (grouped reports, embedded comparative TB, un-evaluated formula
subtotals) for a fictional client - no real client data is in this repo.

`tests/test_formulas.py` additionally verifies the formula-linked schedules
recalculate correctly, using the `formulas` library (a real, independent
Excel formula evaluator) rather than trusting that a formula string merely
looks right - see "Formula-linked schedules" above. It's skipped
automatically unless that dev-only dependency is installed:

```bash
pip install -r requirements-dev.txt
pytest tests/ -v
```

`tests/test_fixed_assets.py` covers the fixed asset register robustness
features directly (in-memory DataFrames, no database): the transaction-
level additions detail behind a category's total, the system-estimated
depreciation columns (specific-before-generic keyword matching, period
proration, never itself flagging a category "review"), and the capex-
miscoding suggestion (client's-own-category vocabulary, the generic
fallback vocabulary for a client with no fixed asset categories yet,
threshold filtering, and correctly ignoring postings already coded to a
fixed asset account).

`tests/test_control_accounts.py` covers the control account rollforward
robustness features directly (in-memory DataFrames, no database): the
movement transaction detail, the movement-by-contact breakdown appearing
only when there's no aged-listing breakdown to conflict with, the
control-account miscoding suggestion - including the regression case
that matters most, that a normal double-entry contra-leg (a receipt
through Bank, an invoice's Sales leg) is never mistaken for a miscoding
- and account-type discovery covering Prepayment/Inventory/Non-current
Asset/Non-current Liability while still excluding Fixed Asset (which
gets its own dedicated treatment elsewhere).

`tests/test_nominal_matrix.py` covers the nominal analysis matrix
robustness features directly (in-memory DataFrames, no database): the
OTHER-bucket breakdown, the regression case that UNALLOCATED itself
folding into OTHER (with enough distinct contra accounts) is never
double-reported alongside its own already-stated message clause, the
material-OTHER note never flipping status to "review" by itself, and the
unallocated-transaction suggestion (a dominant contact history producing
a suggestion, a 50/50 split or too little history correctly producing
none).

`tests/test_statutory_deadlines.py` covers the filing deadline
calculator's date arithmetic directly (pure functions, no data or
database at all): the Companies House Accounting Reference Date special
case for a period end that IS the last day of its calendar month
(31 December, 28 February in a non-leap year) landing on the last day of
the target month rather than the same day-of-month, the plain case where
it isn't, and that HMRC's CT600/CT payment deadlines don't share that
special case.

`tests/test_accruals_prepayments.py` covers the consolidated Accruals &
Prepayments schedule directly (in-memory DataFrames, no database):
discovering Prepayment-typed and accrual-named accounts, excluding an
account already covered by its own dedicated schedule ("Accrued
Corporation Tax") to avoid double-counting the same balance, flagging a
line that doesn't tie to the TB, and the transaction-level detail behind
each line's movement.

`tests/test_going_concern.py` covers the Going Concern indicators check
directly (in-memory DataFrames, no database): a healthy balance sheet
producing no flags, both indicators firing together, a working-capital
deficit that doesn't imply negative net assets overall producing only
one flag - and specifically the sign-convention regression that matters
most, since `financial_statements.build_bs_statement`'s own "Current
liabilities" line is displayed flipped to a positive number (not
negative the way "NET ASSETS" is genuinely signed) - a real bug found
and fixed while building this feature, which would otherwise have
silently inverted the whole analysis.

`tests/test_related_party_transactions.py` covers the Related Party
Transactions check directly (in-memory DataFrames, no database): a
posting to a DLA-identified contact landing outside the DLA account
correctly flagged, the double-entry contra-leg of that same posting
correctly deduplicated rather than counted twice, an unrelated contact
and a below-threshold posting both correctly ignored, and "n/a" when
there's no DLA account or no named DLA contacts to build a related-party
list from at all.

`tests/test_document_detection.py` covers the auto-detection/PDF-extraction
unit logic directly (no database needed): Xero-native matching, the
column-corruption regression (see git history - a failed Xero-native
parse attempt used to permanently corrupt a file's cached column headers
for every later use on the same upload), report-type/platform/period
guessing, and PDF table extraction (including a scanned/no-table PDF
raising a clear error). Its PDF cases build a real test PDF via
`reportlab`, a dev-only dependency (`requirements-dev.txt`) - they skip
automatically if it isn't installed.

`tests/test_recon_vat.py` covers the VAT cross-check's candidate-
reconciling-items enhancement in isolation (in-memory DataFrames, no
database) - flagged vs. ok, with/without nominal activity, matching vs.
non-matching postings.

`tests/test_vat_reconciliation.py` covers the VAT Reconciliation
workspace's matching engine directly (in-memory DataFrames, no
database): the three-pass cascade in isolation (reference match
ignoring amount tolerance and still surfacing the variance, the date/
amount/contact and amount/contact fallback passes, tolerance absorbing
small differences without hiding genuine ones), cash vs accrual basis
changing which pass a match resolves through, each General Ledger row
claimed at most once, Box 1/Box 4 sharing one pool without cross-box
false positives, and multi-file `Source File` tagging surviving into
the unmatched-items table - plus the matched-detail/implied-VAT-rate
advisory (a non-standard rate stated in the message without flipping
status) and the GL-coverage-gap box suggestion (a dominant contact
history producing a suggestion, too little or too mixed history
correctly producing none).

`tests/test_brightpay_reports.py` covers the BrightPay native report
parsers directly (in-memory CSV text, no database): one row per employee
per month for Payroll Summary and Pensions with TOTAL rows and the
trailing annual-summary block correctly excluded, one row per tax month
for P32, native-detection returning the right report type per file (and
`None` for an unrelated file), and a clear error rather than silent
garbage when a file's shape doesn't match.

`tests/test_paye_reconciliation.py` covers the PAYE Reconciliation
workspace's matching engine directly (in-memory DataFrames, no
database): fuzzy employee-name matching (BrightPay's full legal name vs
a shorter General Ledger contact), the double-entry-leg deduplication
found by testing against a real export, zero-net-pay months correctly
skipped, the date window as a hard requirement (not a soft preference -
a regression test for a real bug caught while writing these tests),
tolerance absorbing small differences, HMRC/pension keyword-contact
matching, monthly aggregation across employees for pension, the
closest-same-contact diagnostic on an unmatched item, the cash-
settlement-vs-accrual-only filter (an unpaid invoice correctly excluded
rather than falsely matched, an invoice paid on a separate later date
correctly matching only via its payment row, and the graceful fallback
- with its own stated-in-the-message warning - when a General Ledger
upload doesn't carry a `Source` column at all).

`tests/test_reconciliation_agent.py` covers the AI-assisted-notes agent
directly, with the Anthropic client mocked throughout - no test in this
suite ever calls the real API. Verifies the contract that must hold no
matter what a model says: no API key degrades to a diagnosable message
rather than silence, an API error degrades to a note rather than a raise,
and the "nothing useful to add" marker becomes an empty note rather than
clutter.

`tests/test_storage_and_routes.py` exercises the Postgres-backed storage
layer, the full practice -> template -> client -> job -> upload ->
generate -> download HTTP flow (including a bulk upload of mixed Xero/
CSV/PDF files walked through the auto-detect confirm chain, a multi-sheet
workbook expanding into one upload per sheet, the SSE progress stream
reporting all ten generate steps in order and leaving the job in the same
generated state as the classic POST route, the VAT Reconciliation
workspace end to end (General Ledger + two multi-file Filed Return
uploads, settings, the standalone Run button, and the same results
landing in the generated workbook's own tabs), the PAYE Reconciliation
workspace end to end (a Xero-native General Ledger upload plus three
auto-detected BrightPay uploads, settings, the standalone Run button,
results landing in the generated workbook's own tabs), and - with a
mocked Anthropic client - an AI-assisted note actually reaching the downloaded
workbook end to end), and the access-control model (signup, login, wrong
password, unauthenticated redirect, a preparer scoped to only their
granted clients, a manager blocked from user management, and cross-
practice access denied) against a real (throwaway) Postgres schema
- the part of the app the other tests never touch, since they all work
with in-memory DataFrames/fixtures directly. It's skipped automatically
if no test database is reachable; point `TEST_DATABASE_URL` at one (or
run Postgres locally on the default port with a `wpa_test`/`wpa_test`
role and database) to run it.

## Known limitations / roadmap

- **The template config now fully drives generation.** Schedule
  enable/disable, `insert_after_sheet` positioning, `numbering.start_at`,
  `header_cells` (which cell each schedule's CLIENT NAME/PERIOD/TITLE block
  lands in), and `materiality` (per-template variance/materiality
  thresholds) are all wired in and generate into a copy of the practice's
  real uploaded template file (see "Practices, templates, and clients"
  above). `materiality` reaches every check that flags on a variance -
  both their Python-computed status and the live Excel formulas the
  formula-linked schedules write - rather than the fixed £500/10%
  defaults regardless of what a template configures.
- **Formula-linked output covers the core schedules, not everything yet.**
  TB Lead Schedule, control accounts, P&L/B&S, category-level fixed assets,
  Corporation Tax, and the nominal matrix are all live-formula (see
  "Formula-linked schedules" above), and this now works identically
  whether the base workbook is a fresh one or a loaded practice template.
  The recon check sheets, the asset-level fixed asset register, and the
  closing register still write Python-computed values - converting those
  is the remaining piece.
- **Inserting into a real template file has a known fidelity cost.**
  Round-tripping a real client's `.xlsx` through openpyxl (load, add sheets,
  save) strips embedded images (e.g. a firm's logo) and dropdown data
  validation lists - confirmed by testing against a real 63-sheet working
  paper file, where a plain load-then-save with zero changes dropped it
  from 24MB to 11MB. Cell values, formulas, and most cell-level formatting
  survive intact; visual branding elements and validation dropdowns don't.
  Worth deciding deliberately (accept the cost, or move to lower-level
  zip/XML sheet injection to avoid it) before this ships for real use, not
  something to discover after the fact.
- **Balance Sheet balance check** assumes the standard account-type
  categorisation (Fixed Asset/Current Asset/Bank/Current Liability/
  Liability/Equity) holds for every account - a genuinely miscategorised
  account in the source TB will show up as an unexplained gap in the
  check, which is the intended behaviour (it's real, not a false
  positive). When the miscategorisation is an Account Type this system
  doesn't recognise as one of the five B/S categories at all (rather than
  a P&L account wrongly typed as a B/S one, which still balances
  correctly on the wrong side and can't be told apart from a genuine
  reclassification), `financial_statements.build_bs_statement` now names
  the specific account(s) responsible - they're excluded from every total
  on the statement, not just one side of the check, so a nonzero one is
  almost always the concrete cause. Surfaced in the check's own message
  and as a highlighted detail table on the generated Balance Sheet sheet
  (`StatementResult.unrecognized_detail`).
- **Bank movement analysis** (a receipts/payments-style matrix, like the
  bank schedules in a full audit file) needs a separate Xero bank
  transactions export - not built yet, no sample data to build it against.
- **QBO/Sage** go through generic column mapping only - no dedicated parser
  yet, since there's no real sample export from either to validate against.
- **VAT return and bank statement** have no dedicated Xero parser (VAT
  returns aren't really a Xero "export" as such, and bank statements come
  from the bank, not the bookkeeping platform) - both go through generic
  mapping today.
- **Live API connectors** (pulling directly from Xero/QBO/Sage instead of
  uploading files) - the `DataSource` abstraction is ready for this, not
  built.
- Payroll data (P30/P32/payslips) for a wages control account rollforward
  is a PDF-form-parsing problem, out of scope for now - a structured
  payroll summary upload would be the more tractable first step if needed.
- **Corporation Tax computation** doesn't compute capital allowances or
  identify disallowable expenses itself - both are manual inputs by design,
  since that needs a fixed asset register and a line-by-line expense review
  this system doesn't have. It also doesn't yet handle augmented profits
  that differ from taxable profits (dividends received from non-51%-group
  companies) - defaults to treating them as equal, which is correct for the
  common case but not universal.
- **Fixed asset register** can't distinguish a genuine in-year purchase
  from an opening-balance data-migration journal when flagging "new
  additions" from nominal activity (both look like a debit to a fixed
  asset cost code) - real data showed this: entries in an "Opening
  Balance" section got flagged as additions needing review, which is the
  right call (they do need reviewing) but not always for the reason the
  label implies. It also doesn't try to match a specific flagged disposal
  to a specific register line automatically - that needs the asset
  description reviewed by a preparer.
