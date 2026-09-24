"""Interco Debtor/Creditor detection - which Aged Debtors/Creditors
contacts are related parties, per the client's own maintained name list
(see main.py's save_related_party_names). Not derivable from the aged
report or Trial Balance alone - a related-company relationship is
knowledge the preparer already has, not a pattern in the numbers."""
import pandas as pd


def find_interco_rows(aged_report: pd.DataFrame | None, party_field: str, related_party_names: list[str] | None) -> pd.DataFrame:
    if aged_report is None or aged_report.empty or not related_party_names:
        return pd.DataFrame(columns=[party_field, "total"])
    names = {n.strip().lower() for n in related_party_names if n.strip()}
    if not names:
        return pd.DataFrame(columns=[party_field, "total"])
    match = aged_report[party_field].astype(str).str.strip().str.lower().isin(names)
    return aged_report[match]
