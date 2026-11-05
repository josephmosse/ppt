"""Load and validate the Excel input workbook."""
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

EVE_COLUMNS = [f"EVE {s}" for s in ("Parallel up", "Parallel down", "Steepener", "Flattener", "Short up", "Short down")]
NII_COLUMNS = [f"NII {s}" for s in ("Parallel up", "Parallel down")]

REQUIRED_COLUMNS = {
    "Settings": ["Parameter", "Value"],
    "Repricing_Gap": ["Bucket", "Midpoint_Years", "Assets", "Liabilities", "Derivatives_Net"],
    "Curve": ["Tenor_Years", "Zero_Rate_Pct"],
    "Limits": ["Metric", "Early_Warning_Pct", "Limit_Pct", "SOT_Pct"],
    "History": ["Date", "Tier1", *EVE_COLUMNS, *NII_COLUMNS],
    "Assumptions": ["Item", "Assumption"],
    "Actions": ["Action", "Owner", "Due", "Status"],
}
REQUIRED_SETTINGS = [
    "Institution", "Reporting date", "Currency", "Prepared by", "Committee", "Tier 1 capital",
    "Parallel shock bp", "Short shock bp", "Long shock bp",
]
NUMERIC_COLUMNS = {
    "Repricing_Gap": ["Midpoint_Years", "Assets", "Liabilities", "Derivatives_Net"],
    "Curve": ["Tenor_Years", "Zero_Rate_Pct"],
    "Limits": ["Early_Warning_Pct", "Limit_Pct", "SOT_Pct"],
    "History": ["Tier1", *EVE_COLUMNS, *NII_COLUMNS],
}


class InputError(ValueError):
    """The workbook does not match the expected structure or values."""


@dataclass
class Inputs:
    institution: str
    reporting_date: pd.Timestamp
    currency: str
    prepared_by: str
    committee: str
    tier1: float
    parallel_bp: float
    short_bp: float
    long_bp: float
    gap: pd.DataFrame
    curve: pd.DataFrame
    limits: pd.DataFrame  # indexed by metric: "EVE", "NII"
    history: pd.DataFrame
    assumptions: pd.DataFrame
    actions: pd.DataFrame


def _number(settings, key):
    try:
        return float(settings[key])
    except (TypeError, ValueError):
        raise InputError(f"Settings: '{key}' must be a number, found {settings[key]!r}.") from None


def _check_numeric(sheets):
    for sheet, columns in NUMERIC_COLUMNS.items():
        for column in columns:
            values = pd.to_numeric(sheets[sheet][column], errors="coerce")
            bad = sheets[sheet][values.isna()]
            if not bad.empty:
                # +2: Excel rows are 1-based and row 1 holds the headers.
                raise InputError(f"{sheet}: column '{column}' needs a number in row {bad.index[0] + 2}.")
            sheets[sheet][column] = values


def load_inputs(path) -> Inputs:
    path = Path(path)
    if not path.exists():
        raise InputError(f"Input workbook not found: {path}")
    workbook = pd.read_excel(path, sheet_name=None)

    missing = [sheet for sheet in REQUIRED_COLUMNS if sheet not in workbook]
    if missing:
        raise InputError(f"Missing sheet(s): {', '.join(missing)}.")
    sheets = {name: workbook[name].dropna(how="all").reset_index(drop=True) for name in REQUIRED_COLUMNS}
    for sheet, columns in REQUIRED_COLUMNS.items():
        absent = [c for c in columns if c not in sheets[sheet].columns]
        if absent:
            raise InputError(f"{sheet}: missing column(s) {', '.join(absent)}.")
    _check_numeric(sheets)

    settings = dict(zip(sheets["Settings"]["Parameter"], sheets["Settings"]["Value"]))
    absent = [key for key in REQUIRED_SETTINGS if key not in settings or pd.isna(settings[key])]
    if absent:
        raise InputError(f"Settings: missing value(s) for {', '.join(absent)}.")

    gap = sheets["Repricing_Gap"]
    if gap.empty or not gap["Midpoint_Years"].is_monotonic_increasing or (gap["Midpoint_Years"] <= 0).any():
        raise InputError("Repricing_Gap: Midpoint_Years must be positive and in ascending order.")
    curve = sheets["Curve"]
    tenors = curve["Tenor_Years"]
    if len(curve) < 2 or not (tenors.is_monotonic_increasing and tenors.is_unique):
        raise InputError("Curve: needs at least two rows, with Tenor_Years strictly ascending.")

    limits = sheets["Limits"].set_index("Metric")
    absent = [m for m in ("EVE", "NII") if m not in limits.index]
    if absent:
        raise InputError(f"Limits: missing row(s) for metric {', '.join(absent)}.")
    for metric in ("EVE", "NII"):
        row = limits.loc[metric]
        if not row["Early_Warning_Pct"] <= row["Limit_Pct"] <= row["SOT_Pct"]:
            raise InputError(f"Limits: {metric} must satisfy Early warning <= Limit <= SOT.")

    tier1 = _number(settings, "Tier 1 capital")
    if tier1 <= 0:
        raise InputError("Settings: 'Tier 1 capital' must be positive.")
    try:
        reporting_date = pd.Timestamp(settings["Reporting date"])
    except (TypeError, ValueError):
        raise InputError("Settings: 'Reporting date' must be a date.") from None

    history = sheets["History"].copy()
    history["Date"] = pd.to_datetime(history["Date"], errors="coerce")
    if history["Date"].isna().any():
        raise InputError("History: every row needs a valid Date.")
    if (history["Tier1"] <= 0).any():
        raise InputError("History: Tier1 must be positive.")
    # Earlier periods only: the current period always comes from the calculation.
    history = history[history["Date"] < reporting_date].sort_values("Date").reset_index(drop=True)
    if history.empty:
        raise InputError("History: needs at least one period before the reporting date.")

    actions = sheets["Actions"].copy()
    actions["Due"] = pd.to_datetime(actions["Due"], errors="coerce")

    return Inputs(
        institution=str(settings["Institution"]),
        reporting_date=reporting_date,
        currency=str(settings["Currency"]),
        prepared_by=str(settings["Prepared by"]),
        committee=str(settings["Committee"]),
        tier1=tier1,
        parallel_bp=_number(settings, "Parallel shock bp"),
        short_bp=_number(settings, "Short shock bp"),
        long_bp=_number(settings, "Long shock bp"),
        gap=gap,
        curve=curve,
        limits=limits,
        history=history,
        assumptions=sheets["Assumptions"],
        actions=actions,
    )
