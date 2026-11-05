"""IRRBB calculations: supervisory shocks, delta EVE, delta NII, limit status.

Simplified gap-based engine for a single currency:
- Shocks follow the BCBS d368 standardised functional form, with decay x = 4 years.
- Delta EVE reprices each bucket's net gap as a zero-coupon flow at its midpoint.
- Delta NII is the 12-month effect on buckets repricing within a year, constant balance sheet.
- Shocked rates respect a maturity-dependent floor: -150 bp at t = 0, rising 3 bp a year to 0%.
"""
from dataclasses import dataclass

import numpy as np
import pandas as pd

from .data import Inputs

SCENARIOS = ["Parallel up", "Parallel down", "Steepener", "Flattener", "Short up", "Short down"]
NII_SCENARIOS = ["Parallel up", "Parallel down"]
DECAY_YEARS = 4.0

WITHIN, EARLY_WARNING, LIMIT_BREACH, SOT_BREACH = "Within limit", "Early warning", "Limit breach", "SOT breach"
STATUS_ORDER = [WITHIN, EARLY_WARNING, LIMIT_BREACH, SOT_BREACH]


def shock_bp(scenario, t, parallel, short, long):
    """Rate shock in basis points at maturity t (years)."""
    t = np.asarray(t, dtype=float)
    short_part = short * np.exp(-t / DECAY_YEARS)
    long_part = long * (1 - np.exp(-t / DECAY_YEARS))
    shapes = {
        "Parallel up": parallel + 0 * t,
        "Parallel down": -parallel + 0 * t,
        "Steepener": -0.65 * short_part + 0.9 * long_part,
        "Flattener": 0.8 * short_part - 0.6 * long_part,
        "Short up": short_part,
        "Short down": -short_part,
    }
    return shapes[scenario]


def rate_floor(t):
    """Post-shock floor (decimal): -1.5% at t = 0, +3 bp a year, capped at 0%."""
    return np.minimum(0.0, -0.015 + 0.0003 * np.asarray(t, dtype=float))


def shocked_rates(base, shock, t):
    """Apply a bp shock to decimal rates; a base rate already below the floor is kept."""
    return np.maximum(base + shock / 10000, np.minimum(base, rate_floor(t)))


def delta_eve(gap, t, base, shocked):
    return float(np.sum(gap * (np.exp(-shocked * t) - np.exp(-base * t))))


def delta_nii(gap, t, base, shocked):
    within_year = t < 1
    return float(np.sum((gap * (shocked - base) * (1 - t))[within_year]))


def status_for(value_pct, early_warning, limit, sot):
    """Status of a change in % of Tier 1. Only losses consume a limit."""
    loss = max(-value_pct, 0.0)
    if loss > sot:
        return SOT_BREACH
    if loss > limit:
        return LIMIT_BREACH
    if loss > early_warning:
        return EARLY_WARNING
    return WITHIN


@dataclass
class Results:
    gap: pd.DataFrame  # adds Net_Gap and Cumulative_Gap to the input sheet
    eve: pd.Series  # currency millions, by scenario
    nii: pd.Series
    eve_pct: pd.Series  # % of Tier 1
    nii_pct: pd.Series
    prior_eve: pd.Series
    prior_nii: pd.Series
    prior_eve_pct: pd.Series
    prior_nii_pct: pd.Series
    prior_tier1: float
    prior_date: pd.Timestamp
    trend: pd.DataFrame  # Date, Worst_EVE_Pct, Worst_NII_Pct; ends at the reporting date

    @property
    def worst_eve(self):
        return self.eve_pct.idxmin()

    @property
    def worst_nii(self):
        return self.nii_pct.idxmin()


def _prior_series(row, prefix, scenarios):
    return pd.Series({s: float(row[f"{prefix} {s}"]) for s in scenarios}, dtype=float)


def calculate(inputs: Inputs) -> Results:
    gap = inputs.gap.copy()
    gap["Net_Gap"] = gap["Assets"] - gap["Liabilities"] + gap["Derivatives_Net"]
    gap["Cumulative_Gap"] = gap["Net_Gap"].cumsum()

    t = gap["Midpoint_Years"].to_numpy()
    net = gap["Net_Gap"].to_numpy()
    # Flat extrapolation beyond the curve's first and last tenors.
    base = np.interp(t, inputs.curve["Tenor_Years"], inputs.curve["Zero_Rate_Pct"]) / 100

    eve, nii = {}, {}
    for scenario in SCENARIOS:
        shock = shock_bp(scenario, t, inputs.parallel_bp, inputs.short_bp, inputs.long_bp)
        shocked = shocked_rates(base, shock, t)
        eve[scenario] = delta_eve(net, t, base, shocked)
        if scenario in NII_SCENARIOS:
            nii[scenario] = delta_nii(net, t, base, shocked)
    eve, nii = pd.Series(eve), pd.Series(nii)

    prior = inputs.history.iloc[-1]
    prior_eve = _prior_series(prior, "EVE", SCENARIOS)
    prior_nii = _prior_series(prior, "NII", NII_SCENARIOS)
    prior_tier1 = float(prior["Tier1"])

    eve_pct, nii_pct = eve / inputs.tier1 * 100, nii / inputs.tier1 * 100

    rows = [
        {
            "Date": row["Date"],
            "Worst_EVE_Pct": min(row[f"EVE {s}"] for s in SCENARIOS) / row["Tier1"] * 100,
            "Worst_NII_Pct": min(row[f"NII {s}"] for s in NII_SCENARIOS) / row["Tier1"] * 100,
        }
        for _, row in inputs.history.iterrows()
    ]
    rows.append({"Date": inputs.reporting_date, "Worst_EVE_Pct": eve_pct.min(), "Worst_NII_Pct": nii_pct.min()})

    return Results(
        gap=gap,
        eve=eve,
        nii=nii,
        eve_pct=eve_pct,
        nii_pct=nii_pct,
        prior_eve=prior_eve,
        prior_nii=prior_nii,
        prior_eve_pct=prior_eve / prior_tier1 * 100,
        prior_nii_pct=prior_nii / prior_tier1 * 100,
        prior_tier1=prior_tier1,
        prior_date=prior["Date"],
        trend=pd.DataFrame(rows),
    )
