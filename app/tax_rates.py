"""UK Corporation Tax rates configuration.

Rates have been stable since 1 April 2023 (Finance Act 2021's re-introduction
of the small profits rate + marginal relief regime). Kept as a plain config
here - not hardcoded inside the calculator - specifically so a rate change
is a one-line edit to this file, not a hunt through the computation logic.

A scheduled check (see the "HMRC CT rate watch" Routine) periodically
compares this against gov.uk and notifies if something's changed - but
that's a prompt, not code, so it can't enforce this file gets updated.
Preparers should still treat CT_RATES.as_at as a "verify before relying on
this" date, not a guarantee.
"""
from dataclasses import dataclass


@dataclass(frozen=True)
class CTRates:
    small_profits_rate: float
    main_rate: float
    lower_limit: float          # augmented profits threshold - small profits rate applies below this
    upper_limit: float          # augmented profits threshold - main rate applies above this
    marginal_relief_fraction: float  # "standard fraction" in HMRC's marginal relief formula
    as_at: str                  # date these figures were last verified against gov.uk
    source: str


CT_RATES = CTRates(
    small_profits_rate=0.19,
    main_rate=0.25,
    lower_limit=50_000.0,
    upper_limit=250_000.0,
    marginal_relief_fraction=3 / 200,
    as_at="2026-08-23",
    source="https://www.gov.uk/corporation-tax-rates",
)
