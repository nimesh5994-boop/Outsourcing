"""Best-effort table extraction from a PDF export - for the case where a
client's accounting platform (or the client themselves) only handed over
a PDF, not an Excel/CSV export.

Deliberately narrow scope: finds the single largest table across every
page and treats its first row as the header. That's the right call for
the reports this app cares about (a TB, an aged listing, a nominal
activity report) - one dominant tabular report per file - but won't help
with a PDF containing several separate tables of similar size, or with a
scanned/image-only PDF (no OCR here; the error message says so plainly
rather than pretending to have extracted something it didn't).

Tries pdfplumber's default line-based table detection first (`lines`
strategy - it looks for actual ruled borders), then falls back to a
text-position-based strategy (`text`/`text`) only if that finds nothing
usable. Found live on a real client's PDF export: an Aged Payables Detail
report with no visible gridlines around its data rows (only the Total/
Percentage summary rows happened to sit on a ruled line) - the line
strategy found just those two 1-row fragments and this raised "no table
found" on a perfectly real, text-based, non-scanned export. The text
strategy correctly finds the full 43-row table on the same file by
clustering text positions instead of requiring drawn borders. Line
strategy stays first since it's the more precise signal when a PDF does
have real gridlines (less prone to merging adjacent narrow columns into
one), so this only reaches for the looser strategy when the stricter one
comes up empty.
"""
import io

import pandas as pd
import pdfplumber

_STRATEGIES = [
    None,  # pdfplumber's own default (line-based)
    {"vertical_strategy": "text", "horizontal_strategy": "text"},
]


def _pick_header_row(table: list[list], max_scan: int = 10) -> int:
    """The detected table's own first row isn't always the real header - a
    title/client-name/period line above the actual column headers can get
    swept into the same detected region (found on the same real PDF the
    text-strategy fallback above was added for: 'Aged Payables Detail' /
    the client's name / 'As at ...' each landed as their own 1-cell row
    ahead of the genuine multi-column header). The real header row is the
    one with the most non-blank cells in the first few rows - a title row
    has exactly one, the header row has one per column."""
    def non_blank_count(row) -> int:
        return sum(1 for c in row if c is not None and str(c).strip())

    scan_rows = table[: min(max_scan, len(table))]
    return max(range(len(scan_rows)), key=lambda i: non_blank_count(scan_rows[i]))


def _pick_header_row_cell_count(table: list[list], max_scan: int = 10) -> int:
    """How many filled cells the row _pick_header_row would choose has -
    used to reject a candidate table that isn't real multi-column data
    before committing to it (see extract_table_from_pdf)."""
    idx = _pick_header_row(table, max_scan)
    return sum(1 for c in table[idx] if c is not None and str(c).strip())


def extract_table_from_pdf(content: bytes) -> pd.DataFrame:
    best_table = None
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for table_settings in _STRATEGIES:
            for page in pdf.pages:
                tables = page.extract_tables(table_settings) if table_settings else page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:  # need a header + at least one data row
                        continue
                    if _pick_header_row_cell_count(table) < 2:
                        # Found live on a real client's PDF export: a
                        # report that is genuinely empty of data rows (a
                        # client with zero outstanding aged receivables,
                        # say) still has a title/client-name/period/footer
                        # block on the page, and the text-position strategy
                        # clusters that into a "table" of several 1-cell
                        # rows (e.g. ["Aged Receiv"], [""], ["Shpendi's
                        # Ltd"], ...) - it passes the row-count check above
                        # but every row, including the one _pick_header_row
                        # would choose as "the header", has only a single
                        # filled cell. A real tabular report's header row
                        # has one cell per column. Reject it rather than
                        # returning a nonsense one-column DataFrame.
                        continue
                    if best_table is None or len(table) > len(best_table):
                        best_table = table
            if best_table is not None:
                break  # a stricter/earlier strategy found something usable - don't also try the looser one

    if best_table is None:
        raise ValueError(
            "No table could be found in this PDF. It may be a scanned image "
            "rather than a text-based export (OCR isn't supported here, so "
            "re-export it as Excel/CSV/a text-based PDF instead), or the "
            "report may genuinely have no data rows to show - if so, this "
            "upload can be skipped."
        )

    header_idx = _pick_header_row(best_table)
    header_row, body_rows = best_table[header_idx], best_table[header_idx + 1:]
    header = [str(cell).strip() if cell else f"column_{i + 1}" for i, cell in enumerate(header_row)]

    cleaned_rows = []
    for row in body_rows:
        row = list(row) + [None] * (len(header) - len(row))  # pdfplumber can under/over-detect cells per row
        row = row[: len(header)]
        cleaned_rows.append([str(cell).strip() if cell is not None else "" for cell in row])

    df = pd.DataFrame(cleaned_rows, columns=header)
    return df.loc[:, [c for c in df.columns if c]]  # drop unlabelled stray columns


def extract_leading_text(content: bytes, max_lines: int = 6) -> str:
    """Best-effort read of the first few non-blank lines of text from a
    PDF's first page, independent of table detection - for sniffing a
    report's own title (e.g. "Aged Payables Detail") in cases where the
    table extraction above has already dropped those rows as pre-header
    noise (see _pick_header_row). Used by document classification for
    report types that are structurally identical past their own title
    line - see xero_reports.classify_aged_report_party - since the
    columns extract_table_from_pdf hands back carry no other signal to
    go on. Never raises: a title-sniff is advisory only, so any failure
    here (encrypted PDF, no extractable text) just means no override
    happens, not a broken upload."""
    try:
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            if not pdf.pages:
                return ""
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return ""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    return " ".join(lines[:max_lines])
