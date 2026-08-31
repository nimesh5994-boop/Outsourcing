"""Unit tests for main.py's _jsonable/_df_to_records - the conversion
every standalone section's results go through before being stored as
Postgres JSONB (see storage._put_entity, which only accepts plain JSON
types, not pandas/numpy scalars). No database needed; importing app.main
doesn't connect to one."""
import pandas as pd

from app.main import _df_to_records, _jsonable


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
