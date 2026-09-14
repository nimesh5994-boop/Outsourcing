"""UK Corporation Tax computation.

This is a proforma + rate calculator, not a full tax return: it takes
accounting profit through to taxable profit via manually-entered
adjustments (disallowable expenses, capital allowances - both need data
this system doesn't have, like a fixed asset register or a disallowable-
expense review of the P&L), then applies the current small profits
rate / marginal relief / main rate automatically, and flags a variance
against whatever tax charge is already booked in the trial balance as a
reasonableness check.

Marginal relief formula (standard HMRC method):
    relief = (upper_limit - augmented_profits) * (taxable_profits / augmented_profits) * standard_fraction
    tax    = taxable_profits * main_rate - relief

Thresholds scale with the number of associated companies (divided by
1 + associated_companies) and with a short accounting period (pro-rated
by days/365) per the standard rules.
"""
from dataclasses import dataclass, field

import pandas as pd

from app.tax_rates import CT_RATES, CTRates
from app.xero_reports import PL_ACCOUNT_TYPES

MATERIALITY_AMOUNT = 500.0


@dataclass
class CTComputation:
    accounting_profit: float
    disallowable_additions: float
    capital_allowances: float
    taxable_profit: float
    augmented_profits: float
    associated_companies: int
    period_days: int
    lower_limit: float
    upper_limit: float
    band: str  # "small profits rate" | "marginal relief" | "main rate"
    rate_applied: float  # effective rate, taxable_profit ? tax / taxable_profit : 0
    marginal_relief: float
    tax_charge: float
    booked_tax_charge: float | None
    variance: float | None
    status: str  # "ok" | "review" | "n/a"
    message: str
    rates_used: CTRates = field(default_factory=lambda: CT_RATES)


def compute(
    accounting_profit: float,
    disallowable_additions: float = 0.0,
    capital_allowances: float = 0.0,
    augmented_profits: float | None = None,
    associated_companies: int = 0,
    period_days: int = 365,
    booked_tax_charge: float | None = None,
    materiality: float = MATERIALITY_AMOUNT,
    rates: CTRates = CT_RATES,
) -> CTComputation:
    taxable_profit = accounting_profit + disallowable_additions - capital_allowances
    if augmented_profits is None:
        augmented_profits = taxable_profit

    scale = (1 + max(0, associated_companies))
    period_fraction = min(1.0, max(0.0, period_days / 365.0))
    lower_limit = round(rates.lower_limit / scale * period_fraction, 2)
    upper_limit = round(rates.upper_limit / scale * period_fraction, 2)

    marginal_relief = 0.0
    if taxable_profit <= 0:
        band, rate_applied, tax_charge = "no tax due (loss/nil profit)", 0.0, 0.0
    elif augmented_profits <= lower_limit:
        band, rate_applied = "small profits rate", rates.small_profits_rate
        tax_charge = taxable_profit * rates.small_profits_rate
    elif augmented_profits >= upper_limit:
        band, rate_applied = "main rate", rates.main_rate
        tax_charge = taxable_profit * rates.main_rate
    else:
        band = "marginal relief"
        marginal_relief = (upper_limit - augmented_profits) * (taxable_profit / augmented_profits) * rates.marginal_relief_fraction
        tax_charge = taxable_profit * rates.main_rate - marginal_relief
        rate_applied = (tax_charge / taxable_profit) if taxable_profit else 0.0

    tax_charge = round(tax_charge, 2)
    marginal_relief = round(marginal_relief, 2)

    variance = None
    status, message = "n/a", f"Computed at the {band} - no booked tax charge supplied to compare against."
    if booked_tax_charge is not None:
        variance = round(tax_charge - booked_tax_charge, 2)
        if abs(variance) <= materiality:
            status = "ok"
            message = f"Computed tax charge (£{tax_charge:,.2f} at the {band}) agrees to the booked charge within materiality."
        else:
            status = "review"
            message = (
                f"Computed tax charge (£{tax_charge:,.2f} at the {band}) differs from the booked charge "
                f"(£{booked_tax_charge:,.2f}) by £{abs(variance):,.2f} - review the adjustments used, "
                f"associated companies count, or whether the booked figure needs updating."
            )

    return CTComputation(
        accounting_profit=round(accounting_profit, 2),
        disallowable_additions=round(disallowable_additions, 2),
        capital_allowances=round(capital_allowances, 2),
        taxable_profit=round(taxable_profit, 2),
        augmented_profits=round(augmented_profits, 2),
        associated_companies=associated_companies,
        period_days=period_days,
        lower_limit=lower_limit,
        upper_limit=upper_limit,
        band=band,
        rate_applied=round(rate_applied, 4),
        marginal_relief=marginal_relief,
        tax_charge=tax_charge,
        booked_tax_charge=booked_tax_charge,
        variance=variance,
        status=status,
        message=message,
        rates_used=rates,
    )


def find_tax_provision_account(tb_current: pd.DataFrame | None) -> tuple[str, str] | None:
    """The balance-sheet Corporation Tax provision/payable account (not the
    P&L tax charge line - excluded by requiring a non-P&L account_type),
    so a control-account-style rollforward (see control_accounts.py) can
    check whether the provision itself - b/fwd + this year's charge
    posted - payments made = c/fwd - actually ties to the trial balance,
    the same tie-out every other balance-sheet control account already
    gets. Takes the first match if more than one account name contains
    "corporation tax" on the balance sheet side - a genuine edge case
    (most charts of accounts have exactly one), not handled further."""
    if tb_current is None or tb_current.empty:
        return None
    is_bs_account = ~tb_current["account_type"].astype(str).str.lower().isin(PL_ACCOUNT_TYPES)
    matches = tb_current[tb_current["account_name"].astype(str).str.lower().str.contains("corporation tax", na=False) & is_bs_account]
    if matches.empty:
        return None
    row = matches.iloc[0]
    return str(row["account_code"]), row["account_name"]
