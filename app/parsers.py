"""File loading + normalisation into canonical DataFrames.

`DataSource` is the seam for future live API connectors: today the only
implementation is `FileDataSource` (reads an uploaded CSV/XLSX), but a
`XeroApiDataSource` / `QboApiDataSource` etc. could be dropped in later
without touching mapping, reconciliation or the Excel builder — they all
just consume a canonical DataFrame, however it was produced.
"""
import io
from abc import ABC, abstractmethod
from pathlib import Path

import pandas as pd

from app.models import REPORT_SCHEMAS


class DataSource(ABC):
    @abstractmethod
    def raw_columns(self) -> list[str]:
        ...

    @abstractmethod
    def raw_dataframe(self) -> pd.DataFrame:
        ...


class FileDataSource(DataSource):
    """Loads a CSV or XLSX export exactly as downloaded from the platform -
    either from a local path, or from raw bytes already held in memory (the
    case once uploaded files live in Postgres rather than on disk: a
    filesystem path wouldn't survive from one serverless invocation to the
    next, so the upload route reads the file into bytes once and every
    later step - mapping, parsing, generation - works from those bytes
    instead of re-opening a path that a different instance's ephemeral
    filesystem never had)."""

    def __init__(self, source: str | Path | bytes, filename: str | None = None, sheet_name: str | None = None):
        if isinstance(source, (bytes, bytearray)):
            if filename is None:
                raise ValueError("filename is required when source is raw bytes (needed to tell CSV from XLSX)")
            self._buffer: bytes | None = bytes(source)
            self._suffix = Path(filename).suffix.lower()
        else:
            self.file_path = Path(source)
            self._buffer = None
            self._suffix = self.file_path.suffix.lower()
        # None = the sheet pandas defaults to (the first one) - an .xlsx
        # with several sheets (e.g. a VAT return export with separate
        # "Summary" and "Detail" tabs) is otherwise silently reduced to
        # whichever sheet happens to load first; see excel_sheet_names()
        # and how main.py's upload route expands a multi-sheet file into
        # one classified sub-upload per sheet instead.
        self._sheet_name = sheet_name
        self._df = None

    def _load(self) -> pd.DataFrame:
        if self._df is None:
            handle = io.BytesIO(self._buffer) if self._buffer is not None else self.file_path
            if self._suffix == ".pdf":
                from app.pdf_extraction import extract_table_from_pdf

                content = self._buffer if self._buffer is not None else self.file_path.read_bytes()
                self._df = extract_table_from_pdf(content)
            elif self._suffix in (".xlsx", ".xls"):
                sheet = self._sheet_name if self._sheet_name is not None else 0
                self._df = pd.read_excel(handle, sheet_name=sheet, dtype=str)
            else:
                self._df = pd.read_csv(handle, dtype=str, keep_default_na=False)
            self._df.columns = [str(c).strip() for c in self._df.columns]
        return self._df

    def raw_columns(self) -> list[str]:
        return list(self._load().columns)

    def raw_dataframe(self) -> pd.DataFrame:
        return self._load()


def excel_sheet_names(content: bytes) -> list[str]:
    """Every sheet name in an .xlsx/.xls file, in workbook order - used to
    decide whether an upload needs expanding into one sub-upload per sheet
    rather than silently only ever reading the first one."""
    import openpyxl

    wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


NUMERIC_FIELDS = {
    "debit", "credit", "current", "bucket_1", "bucket_2", "bucket_3", "bucket_4", "older", "total",
    "box1", "box2", "box3", "box4", "box5", "box6", "box7", "box8", "box9",
    "closing_balance", "amount", "cost", "depreciation_rate", "accumulated_depreciation_b_fwd",
    "net_amount", "vat_amount",
}

DATE_FIELDS = {"date", "date_acquired", "payment_date"}


def _to_numeric(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("£", "", regex=False)
        .str.strip()
    )
    # accounting-style negatives: (123.45) -> -123.45
    is_paren = cleaned.str.startswith("(") & cleaned.str.endswith(")")
    cleaned = cleaned.where(~is_paren, "-" + cleaned.str.strip("()"))
    cleaned = cleaned.replace({"": "0", "-": "0", "nan": "0"})
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def apply_mapping(source: DataSource, report_type: str, mapping: dict) -> pd.DataFrame:
    """mapping: {source_column: canonical_field or None}. Returns a
    DataFrame with exactly the canonical columns for report_type, numeric
    fields coerced to float."""
    raw = source.raw_dataframe()
    canonical_fields = list(REPORT_SCHEMAS[report_type].keys())
    out = pd.DataFrame(index=raw.index)

    for source_col, canonical_field in mapping.items():
        if canonical_field and source_col in raw.columns:
            out[canonical_field] = raw[source_col]

    for field in canonical_fields:
        if field not in out.columns:
            out[field] = 0.0 if field in NUMERIC_FIELDS else ""

    for field in canonical_fields:
        if field in NUMERIC_FIELDS:
            out[field] = _to_numeric(out[field])
        else:
            out[field] = out[field].astype(str).replace("nan", "").str.strip()

    for field in canonical_fields:
        if field in DATE_FIELDS:
            out[field] = pd.to_datetime(out[field], errors="coerce", dayfirst=True)

    if report_type == "trial_balance":
        out["balance"] = out["debit"] - out["credit"]
    if report_type == "nominal_activity":
        out["net"] = out["debit"] - out["credit"]

    return out[[c for c in out.columns]]
