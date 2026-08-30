"""Unit tests for the nominal activity matrix robustness features added
on top of build_matrix - see app/nominal_matrix.py:

  - extra_detail: what's actually inside the "OTHER" catch-all column
    (each folded-in contra account, its own net amount and transaction
    count), so OTHER is never a dead end for a reviewer.
  - the materiality note in the message when a contra account was folded
    into OTHER purely because of the top-N column cap, not because it's
    immaterial - advisory only, never itself flips status to "review".
  - suggest_unallocated_reallocations: for an unallocated transaction,
    suggest a likely contra account based on that SAME contact's other,
    already-allocated transactions on the same account - only when one
    contra account clearly dominates that contact's history.

Existing build_matrix/build_all_matrices behaviour (contra-account
pivoting, multi-code-split flagging) is covered by tests/test_pipeline.py
and tests/test_formulas.py - not repeated here."""
import pandas as pd

from app import nominal_matrix as nm


def _row(date, code, name, debit=0.0, credit=0.0, contra_code="", contra_name="",
         contra_needs_review=False, contact="", description="", reference=""):
    return {
        "date": pd.Timestamp(date), "account_code": code, "account_name": name,
        "reference": reference, "description": description, "contact": contact,
        "debit": debit, "credit": credit,
        "contra_code": contra_code, "contra_name": contra_name, "contra_needs_review": contra_needs_review,
    }


# --- OTHER-bucket breakdown + materiality note ---------------------------

def test_other_bucket_breakdown_lists_the_folded_in_contra_accounts():
    # 12 distinct contra accounts, all above the top-10 cap, descending size
    rows = [
        _row("2025-01-01", "1200", "BANK", credit=1000.0 - i * 10, contra_code=str(4000 + i),
             contra_name=f"Expense {i}", contact=f"Supplier {i}", reference=f"R{i}")
        for i in range(12)
    ]
    nominal = pd.DataFrame(rows)
    result = nm.build_matrix("1200", "BANK", nominal)
    assert not result.extra_detail.empty
    assert len(result.extra_detail) == 2  # the 2 accounts that missed the top-10 cut
    assert set(result.extra_detail["Contra account"]) == {"4010 - Expense 10", "4011 - Expense 11"}
    assert "folded into" in result.message.lower()
    assert "top 10 columns" in result.message


def test_material_other_note_never_flips_status_to_review():
    rows = [
        _row("2025-01-01", "1200", "BANK", credit=1000.0 - i * 10, contra_code=str(4000 + i),
             contra_name=f"Expense {i}", contact=f"Supplier {i}", reference=f"R{i}")
        for i in range(12)
    ]
    nominal = pd.DataFrame(rows)
    result = nm.build_matrix("1200", "BANK", nominal)
    # every transaction here is cleanly allocated (no unallocated, no
    # multi-code splits) - the material-OTHER note is advisory only and
    # must not turn an otherwise-clean matrix into "review"
    assert result.status == "ok"


def test_other_bucket_breakdown_excludes_unallocated_even_when_it_falls_into_other():
    # with enough distinct contra accounts, UNALLOCATED itself can get
    # folded into OTHER by the same top-N cap - it must not then be
    # double-reported as if it were just another hidden contra account,
    # since its amount is already called out on its own in the message
    rows = [
        _row("2025-01-01", "1200", "BANK", credit=1000.0 - i * 10, contra_code=str(4000 + i),
             contra_name=f"Expense {i}", contact=f"Supplier {i}", reference=f"R{i}")
        for i in range(12)
    ]
    rows.append(_row("2025-06-01", "1200", "BANK", credit=250.0, contact="Acme Ltd", reference="ACME-U"))
    nominal = pd.DataFrame(rows)
    result = nm.build_matrix("1200", "BANK", nominal)
    assert not any(str(c).startswith("UNALLOCATED") for c in result.extra_detail["Contra account"])
    assert result.message.count("£-250") == 1  # only in the unallocated clause, not repeated in the OTHER note
    assert "is unallocated to any contra code" in result.message
    assert "2 account(s)" in result.message  # the 2 genuinely-hidden expense accounts, not counting UNALLOCATED as a 3rd


def test_no_extra_detail_when_fewer_than_max_contra_columns():
    rows = [
        _row("2025-01-01", "1200", "BANK", credit=100.0, contra_code="4000", contra_name="Office Costs", contact="Acme"),
        _row("2025-02-01", "1200", "BANK", credit=200.0, contra_code="4001", contra_name="Travel", contact="Acme"),
    ]
    nominal = pd.DataFrame(rows)
    result = nm.build_matrix("1200", "BANK", nominal)
    assert result.extra_detail.empty
    assert "folded into" not in result.message.lower()


# --- suggest_unallocated_reallocations -----------------------------------

def test_suggests_contra_account_from_dominant_contact_history():
    rows = [
        _row("2025-01-01", "1200", "BANK", credit=100.0, contra_code="4000", contra_name="Office Costs", contact="Acme Ltd", reference="R1"),
        _row("2025-02-01", "1200", "BANK", credit=120.0, contra_code="4000", contra_name="Office Costs", contact="Acme Ltd", reference="R2"),
        _row("2025-03-01", "1200", "BANK", credit=90.0, contra_code="4000", contra_name="Office Costs", contact="Acme Ltd", reference="R3"),
        _row("2025-04-01", "1200", "BANK", credit=150.0, contact="Acme Ltd", reference="R4"),  # unallocated
    ]
    nominal = pd.DataFrame(rows)
    result = nm.suggest_unallocated_reallocations(nominal, account_codes=["1200"])
    assert result.status == "review"
    assert len(result.detail) == 1
    row = result.detail.iloc[0]
    assert row["Suggested contra account"] == "4000 - Office Costs"
    assert row["Contact"] == "Acme Ltd"


def test_no_suggestion_without_enough_history():
    rows = [
        _row("2025-01-01", "1200", "BANK", credit=100.0, contra_code="4000", contra_name="Office Costs", contact="Acme Ltd", reference="R1"),
        _row("2025-02-01", "1200", "BANK", credit=150.0, contact="Acme Ltd", reference="R2"),  # unallocated
        _row("2025-03-01", "1200", "BANK", credit=75.0, contact="New Supplier Ltd", reference="R3"),  # unallocated, no history at all
    ]
    nominal = pd.DataFrame(rows)
    result = nm.suggest_unallocated_reallocations(nominal, account_codes=["1200"])
    assert result.status == "ok"


def test_no_suggestion_when_contact_history_is_mixed_not_dominant():
    rows = [
        _row("2025-01-01", "1200", "BANK", credit=100.0, contra_code="4000", contra_name="Office Costs", contact="Acme Ltd", reference="R1"),
        _row("2025-02-01", "1200", "BANK", credit=100.0, contra_code="4001", contra_name="Travel", contact="Acme Ltd", reference="R2"),
        _row("2025-03-01", "1200", "BANK", credit=150.0, contact="Acme Ltd", reference="R3"),  # unallocated, 50/50 split history
    ]
    nominal = pd.DataFrame(rows)
    result = nm.suggest_unallocated_reallocations(nominal, account_codes=["1200"])
    assert result.status == "ok"


def test_na_without_contra_column():
    nominal = pd.DataFrame([{"date": pd.Timestamp("2025-01-01"), "account_code": "1200", "account_name": "BANK",
                              "reference": "R1", "description": "", "contact": "Acme", "debit": 0.0, "credit": 100.0}])
    result = nm.suggest_unallocated_reallocations(nominal)
    assert result.status == "n/a"
    assert nm.suggest_unallocated_reallocations(None).status == "n/a"
    assert nm.suggest_unallocated_reallocations(pd.DataFrame()).status == "n/a"
