"""File loading + normalisation into canonical DataFrames.

`DataSource` is the seam for future live API connectors: today the only
implementation is `FileDataSource` (reads an uploaded CSV/XLSX), but a
`XeroApiDataSource` / `QboApiDataSource` etc. could be dropped in later
without touching mapping, reconciliation or the Excel builder — they all
just consume a canonical DataFrame, however it was produced.
"""
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
    """Loads a CSV or XLSX export exactly as downloaded from the platform."""

    def __init__(self, file_path: str | Path):
        self.file_path = Path(file_path)
        self._df = None

    def _load(self) -> pd.DataFrame:
        if self._df is None:
            if self.file_path.suffix.lower() in (".xlsx", ".xls"):
                self._df = pd.read_excel(self.file_path, dtype=str)
            else:
                self._df = pd.read_csv(self.file_path, dtype=str, keep_default_na=False)
            self._df.columns = [str(c).strip() for c in self._df.columns]
        return self._df

    def raw_columns(self) -> list[str]:
        return list(self._load().columns)

    def raw_dataframe(self) -> pd.DataFrame:
        return self._load()


NUMERIC_FIELDS = {
    "debit", "credit", "current", "bucket_1", "bucket_2", "bucket_3", "bucket_4", "older", "total",
    "box1", "box2", "box3", "box4", "box5", "box6", "box7", "box8", "box9",
    "closing_balance", "amount", "cost", "depreciation_rate", "accumulated_depreciation_b_fwd",
}

DATE_FIELDS = {"date", "date_acquired"}


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
