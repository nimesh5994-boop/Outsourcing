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

from app.tax_rates import CT_RATES, CTRates


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
    materiality: float = 500.0,
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
