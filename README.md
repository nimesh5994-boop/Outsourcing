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

## What gets checked

| Check | What it does |
|---|---|
| TB self-balance | Debits = credits for both years |
| TB → P&L/B&S tie-out | Rough top-level sanity check (see Limitations) |
| Variance analysis | Every nominal code, current vs comparative, flagged if it moves >10% and >£500 |
| Nominal activity review | Flags suspense postings, round-sum manual journals, and descriptions that read like pending corrections |
| Debtors/creditors control recon | Aged listing total vs TB control account balance |
| Control account rollforwards | B/fwd + movements = c/fwd for any balance-sheet account with nominal detail, with the aged listing attached as a breakdown of debtors/creditors closing balances specifically |
| Nominal analysis matrix | Every transaction against an account allocated to its contra nominal code; multi-way splits and unallocated amounts flagged for manual review |
| Bank reconciliation | Statement closing balance vs TB |
| VAT cross-check | VAT return boxes vs P&L turnover and VAT control account |
| Corporation Tax computation | Current UK rates (small profits/marginal relief/main rate) applied to accounting profit, checked against the booked tax charge |
| Fixed asset register (category) | Cost/depreciation rollforward per asset category, derived from TB + nominal activity, checked against the TB |
| Fixed asset register (asset detail) | Prior-year register rolled forward asset-by-asset, new additions/possible disposals flagged from nominal activity, totals checked against TB |

Every check produces a status (`ok` / `review` / `error` / `n/a`) and a
plain-English message, shown on the workbook's Index sheet and again on
each schedule.

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

## Fixed asset register

Built from `app/fixed_assets.py`, in two parts that work independently:

**Category-level** (always available, no upload needed): real charts of
accounts commonly split each fixed asset category into a cost/additions
code and a depreciation code (naming varies a lot in practice - some use
"COMPUTER EQUIPMENT - COST", others run words together with no separator
at all, e.g. "IT EQUIPMENT COST BROUGHT FORWARD" - both are handled by
stripping structural words like cost/additions/depreciation/brought/
forward wherever they appear, rather than matching a fixed suffix
pattern). Paired-up cost and depreciation accounts become a cost /
accumulated depreciation / NBV rollforward per category, cross-checked
against the TB.

**Asset-level** (needs a prior-year register upload - report type "Fixed
Asset Register", generic column mapping since there's no standard
software export for this): rolls each asset forward using its own
depreciation method (straight line or reducing balance) and rate,
flags nominal-activity transactions against fixed asset cost codes that
aren't yet in the register as candidate new additions, flags credit
movements on those codes as possible disposals to match and remove, and
reconciles the register's total NBV to the TB.

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
  recon.py                 the reconciliation/cross-check engine
  control_accounts.py       control account rollforward + aged breakdown engine
  nominal_matrix.py          nominal activity → contra nominal code analysis matrix
  excel_builder.py             builds the final .xlsx, house-style header blocks + numbered index
  storage.py                    filesystem-backed client/job storage
  main.py                        FastAPI app + routes
  templates/, static/             the (minimal) web UI
```

`DataSource` in `parsers.py` is the seam for live API connectors later
(Xero/QBO/Sage OAuth pulls instead of file uploads) - today the only
implementation is `FileDataSource`, but nothing else in the pipeline would
need to change to add one, since everything downstream just consumes a
canonical DataFrame.

## Running it

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Then open http://127.0.0.1:8000 - create a client, start a job with the
period dates, upload the reports you have, confirm any mapping that needs
review, and generate.

## Tests

```bash
pytest tests/ -v
```

`tests/test_pipeline.py` runs the full pipeline (parse → reconcile → build
workbook) against `sample_data/`, which mirrors the real Xero export
structures (grouped reports, embedded comparative TB, un-evaluated formula
subtotals) for a fictional client - no real client data is in this repo.

## Known limitations / roadmap

- **TB → P&L/B&S tie-out** is a rough top-level check today, not a true
  per-account tie-out (P&L and B/S use different sign conventions than the
  raw TB debit/credit). Worth tightening once there's a second real
  client's data to validate against.
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
