"""Create data/irrbb_input.xlsx with ILLUSTRATIVE figures for a fictional bank. Replace them with your own."""
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from irrbb.calc import NII_SCENARIOS, SCENARIOS, calculate  # noqa: E402
from irrbb.data import load_inputs  # noqa: E402

OUTPUT = ROOT / "data" / "irrbb_input.xlsx"
REPORTING_DATE = pd.Timestamp("2026-09-30")
TIER1 = 1850.0
# Quarter-by-quarter scale on today's sensitivities: oldest first, giving a plausible path.
HISTORY_SCALES = [0.70, 0.74, 0.78, 0.83, 0.80, 0.86, 0.91]
HISTORY_TIER1 = [1720.0, 1738.0, 1755.0, 1772.0, 1790.0, 1810.0, 1830.0]

SETTINGS = [
    ("Institution", "Sample Bank (illustrative data)"),
    ("Reporting date", REPORTING_DATE),
    ("Currency", "EUR"),
    ("Prepared by", "ALM team"),
    ("Committee", "ALCO"),
    ("Tier 1 capital", TIER1),
    ("Parallel shock bp", 200),
    ("Short shock bp", 250),
    ("Long shock bp", 100),
]
# Derivatives_Net: payer swaps add to short gaps (receive floating) and reduce long gaps (pay fixed); they net to zero.
GAP = [
    ("O/N", 0.003, 2400, 2300, 0), ("1M", 0.04, 1900, 1700, 0), ("3M", 0.17, 2600, 1900, 900),
    ("6M", 0.375, 1500, 1300, 300), ("1Y", 0.75, 1400, 1200, 0), ("2Y", 1.5, 1300, 1450, 0),
    ("3Y", 2.5, 1250, 1400, 0), ("5Y", 4.0, 2100, 2100, -300), ("7Y", 6.0, 1600, 1100, -300),
    ("10Y", 8.5, 1900, 700, -400), ("15Y", 12.5, 700, 250, -200), ("20Y", 17.5, 300, 100, 0), (">20Y", 25.0, 100, 0, 0),
]
CURVE = [(0.25, 2.05), (0.5, 2.05), (1, 2.10), (2, 2.20), (3, 2.30), (5, 2.50), (7, 2.65), (10, 2.85), (15, 3.00), (20, 3.05), (30, 3.05)]
LIMITS = [("EVE", 10.0, 12.0, 15.0), ("NII", 3.0, 4.0, 5.0)]
ASSUMPTIONS = [
    ("Sensitivity method", "Net repricing gap by bucket, repriced at bucket midpoint; single currency"),
    ("Non-maturity deposits", "Behavioural repricing profile in the gap; average repricing maturity 3.2 years"),
    ("Loan prepayment", "6% constant prepayment rate, reflected in the asset profile"),
    ("Own equity", "Excluded from liabilities in the EVE measure"),
    ("NII horizon", "12 months, constant balance sheet, instantaneous shock"),
    ("Post-shock floor", "\u2212150 bp at overnight, rising 3 bp a year to 0%"),
]
ACTIONS = [
    ("Review the payer swap programme for the 5Y to 10Y buckets", "Head of ALM", "2026-11-30", "In progress"),
    ("Back-test the non-maturity deposit model", "Model Risk", "2026-12-15", "Open"),
    ("Refresh the prepayment assumption with Q3 data", "ALM team", "2026-12-31", "Open"),
    ("Update the IRRBB limit framework", "CRO office", "2026-09-30", "Done"),
]


def write_workbook(history_rows):
    gap_columns = ["Bucket", "Midpoint_Years", "Assets", "Liabilities", "Derivatives_Net"]
    sheets = {
        "Settings": pd.DataFrame(SETTINGS, columns=["Parameter", "Value"]),
        "Repricing_Gap": pd.DataFrame(GAP, columns=gap_columns),
        "Curve": pd.DataFrame(CURVE, columns=["Tenor_Years", "Zero_Rate_Pct"]),
        "Limits": pd.DataFrame(LIMITS, columns=["Metric", "Early_Warning_Pct", "Limit_Pct", "SOT_Pct"]),
        "History": pd.DataFrame(history_rows),
        "Assumptions": pd.DataFrame(ASSUMPTIONS, columns=["Item", "Assumption"]),
        "Actions": pd.DataFrame(
            [(a, o, pd.Timestamp(d), s) for a, o, d, s in ACTIONS], columns=["Action", "Owner", "Due", "Status"]
        ),
    }
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT, engine="openpyxl", datetime_format="yyyy-mm-dd") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False)
            sheet = writer.sheets[name]
            sheet.freeze_panes = "A2"
            for index, column in enumerate(frame.columns, start=1):
                longest = max([len(str(column)), *(len(str(v)) for v in frame[column])])
                sheet.column_dimensions[sheet.cell(1, index).column_letter].width = min(longest + 2, 70)


def history_row(date, tier1, scale, results=None):
    row = {"Date": date, "Tier1": tier1}
    for scenario in SCENARIOS:
        row[f"EVE {scenario}"] = round(results.eve[scenario] * scale, 1) if results else 0.0
    for scenario in NII_SCENARIOS:
        row[f"NII {scenario}"] = round(results.nii[scenario] * scale, 1) if results else 0.0
    return row


def main():
    dates = pd.date_range(end=REPORTING_DATE, periods=len(HISTORY_SCALES) + 1, freq="QE")[:-1]
    # History is scaled from the current sensitivities, so write once to calculate, then again in full.
    write_workbook([history_row(dates[0], TIER1, 0)])
    results = calculate(load_inputs(OUTPUT))
    write_workbook([history_row(d, t, s, results) for d, t, s in zip(dates, HISTORY_TIER1, HISTORY_SCALES)])
    print(f"Wrote {OUTPUT}")
    print((results.eve_pct.round(2)).to_string())
    print((results.nii_pct.round(2)).to_string())


if __name__ == "__main__":
    main()
