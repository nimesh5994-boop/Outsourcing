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
"""
import io

import pandas as pd
import pdfplumber


def extract_table_from_pdf(content: bytes) -> pd.DataFrame:
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        best_table = None
        for page in pdf.pages:
            for table in page.extract_tables():
                if not table or len(table) < 2:  # need a header + at least one data row
                    continue
                if best_table is None or len(table) > len(best_table):
                    best_table = table

    if best_table is None:
        raise ValueError(
            "No table could be found in this PDF. It may be a scanned image "
            "rather than a text-based export - OCR isn't supported here, so "
            "re-export it as Excel/CSV/a text-based PDF instead."
        )

    header_row, body_rows = best_table[0], best_table[1:]
    header = [str(cell).strip() if cell else f"column_{i + 1}" for i, cell in enumerate(header_row)]

    cleaned_rows = []
    for row in body_rows:
        row = list(row) + [None] * (len(header) - len(row))  # pdfplumber can under/over-detect cells per row
        row = row[: len(header)]
        cleaned_rows.append([str(cell).strip() if cell is not None else "" for cell in row])

    df = pd.DataFrame(cleaned_rows, columns=header)
    return df.loc[:, [c for c in df.columns if c]]  # drop unlabelled stray columns
