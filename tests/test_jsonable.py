"""Unit tests for main.py's _jsonable/_df_to_records/_df_columns - the
conversion every standalone section's results go through before being
stored as Postgres JSONB (see storage._put_entity, which only accepts
plain JSON types, not pandas/numpy scalars). No database needed;
importing app.main doesn't connect to one."""
import pandas as pd

from app.main import _df_columns, _df_to_records, _jsonable, _recon_result_to_dict
from app.recon import ReconResult


def test_jsonable_converts_a_real_timestamp():
    assert _jsonable(pd.Timestamp("2025-06-15")) == "2025-06-15"


def test_jsonable_converts_bare_nat_to_none():
    # Regression test for a real bug: pd.NaT is its own singleton type
    # (NaTType), NOT a pd.Timestamp subclass, so `isinstance(value,
    # pd.Timestamp)` never catches it - a bare NaT (e.g. a date field left
    # blank because the source row's date genuinely didn't parse, or a
    # combined-match row referencing a filed/GL row with no date) reached
    # json.dumps() unconverted and crashed the whole save_job() call with
    # an unhandled 500, found live when a VAT reconciliation result
    # contained an unmapped date column.
    assert _jsonable(pd.NaT) is None


def test_jsonable_converts_nan_and_numpy_scalars():
    assert _jsonable(float("nan")) is None
    import numpy as np
    assert _jsonable(np.float64(1.5)) == 1.5
    assert _jsonable(np.int64(7)) == 7
    assert _jsonable(np.bool_(True)) is True


def test_jsonable_passes_through_plain_values():
    assert _jsonable(None) is None
    assert _jsonable("hello") == "hello"
    assert _jsonable(42) == 42


def test_df_to_records_survives_a_bare_nat_column():
    df = pd.DataFrame([
        {"Date": pd.NaT, "Amount": 100.0},
        {"Date": pd.Timestamp("2025-01-01"), "Amount": 200.0},
    ])
    records = _df_to_records(df)
    assert records[0]["Date"] is None
    assert records[1]["Date"] == "2025-01-01"


def test_df_columns_preserves_the_dataframes_own_order():
    # Regression test for a real bug: Postgres jsonb does NOT preserve an
    # object's key insertion order on a save/reload round trip (it
    # re-orders keys by length, then lexicographically), so a standalone
    # section's results table rendered with a visibly scrambled column
    # order (TOTAL/DIFF ahead of Date, contra-account columns out of the
    # rank-by-value order the check itself put them in) even though every
    # row's values still lined up correctly under their own equally-
    # scrambled keys - found live via a screenshot of the nominal matrix's
    # standalone section, not caught by any test that only checked cell
    # values. _df_columns captures the DataFrame's real column order
    # separately so job_detail.html can render columns in that order
    # rather than trusting a stored record's own key order.
    df = pd.DataFrame([{"Date": "2025-01-01", "Reference": "R1", "Description": "d", "Contact": "c", "TOTAL": 1, "DIFF": 0}])
    assert _df_columns(df) == ["Date", "Reference", "Description", "Contact", "TOTAL", "DIFF"]


def test_df_columns_empty_for_none_or_empty_dataframe():
    assert _df_columns(None) == []
    assert _df_columns(pd.DataFrame()) == []


def test_recon_result_to_dict_includes_column_order_for_every_table():
    detail = pd.DataFrame([{"Z first": 1, "A second": 2}])
    extra = pd.DataFrame([{"Z first": 3, "A second": 4}])
    matched = pd.DataFrame([{"Z first": 5, "A second": 6}])
    result = ReconResult("Test check", "ok", "message", detail, extra, "extra label", matched, "matched label")
    out = _recon_result_to_dict(result)
    assert out["detail_columns"] == ["Z first", "A second"]
    assert out["extra_detail_columns"] == ["Z first", "A second"]
    assert out["matched_detail_columns"] == ["Z first", "A second"]
