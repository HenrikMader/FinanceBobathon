"""
Budget Report Generator — Acme GmbH Q1 2024
============================================
This script reads the Q1 transaction data and produces a summary Excel report.

STATUS: Work in progress — several things are broken or missing.
        Use IBM Bob to find and fix the issues, then extend the script.
"""

import csv
from pathlib import Path
from datetime import datetime

# Try to import openpyxl — needed for the Excel output in Part 3
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ── Configuration ─────────────────────────────────────────────────────────────

DATA_DIR   = Path(__file__).parent.parent / "data"
OUTPUT_DIR = Path(__file__).parent.parent / "output"

TRANSACTIONS_FILE = DATA_DIR / "transactions_q1.csv"
OUTPUT_FILE       = OUTPUT_DIR / "q1_budget_report.xlsx"

# Monthly budget targets per cost center (EUR)
# Source: annual plan divided by 4 quarters, then by 3 months
MONTHLY_TARGETS = {
    "CC100": 200_000 / 3,
    "CC200": 133_333 / 3,
    "CC300": 600_000 / 3,
    "CC400": 100_000 / 3,
    # BUG: CC500 is missing — Finance cost center has no target defined
}

VALID_COST_CENTERS = ["CC100", "CC200", "CC300", "CC400", "CC500"]


# ── Step 1: Load transactions ─────────────────────────────────────────────────

def load_transactions(filepath: Path) -> list[dict]:
    """Load CSV transactions and do basic type conversion."""
    transactions = []
    with open(filepath, newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            # BUG: date is read as a string but never converted to a date object
            # This causes the monthly grouping below to fail silently
            row["amount_eur"] = float(row["amount_eur"])
            transactions.append(row)
    return transactions


# ── Step 2: Filter & validate ─────────────────────────────────────────────────

def filter_approved(transactions: list[dict]) -> list[dict]:
    """Keep only approved transactions."""
    # BUG: comparison is case-sensitive but CSV has mixed casing ("Yes"/"yes")
    return [t for t in transactions if t["approved"] == "yes"]


def validate_cost_centers(transactions: list[dict]) -> list[dict]:
    """Warn about unknown cost centers and return only valid rows."""
    valid = []
    for t in transactions:
        if t["cost_center"] not in VALID_COST_CENTERS:
            print(f"  WARNING: Unknown cost center '{t['cost_center']}' in {t['transaction_id']}")
        else:
            valid.append(t)
    return valid


# ── Step 3: Aggregate by cost center & month ─────────────────────────────────

def aggregate(transactions: list[dict]) -> dict:
    """
    Returns a nested dict:
        { cost_center: { month_number: total_amount } }
    """
    result = {}
    for t in transactions:
        cc    = t["cost_center"]
        # BUG: t["date"] is still a string — this will crash with AttributeError
        month = t["date"].month
        amt   = t["amount_eur"]

        if cc not in result:
            result[cc] = {}
        result[cc][month] = result[cc].get(month, 0) + amt

    return result


# ── Step 4: Find anomalies ────────────────────────────────────────────────────

def find_duplicates(transactions: list[dict]) -> list[tuple]:
    """
    Find potential duplicate bookings: same vendor, same amount, within 3 days.
    Returns a list of (txn_a, txn_b) pairs.
    """
    # TODO: Implement this function.
    #       Hint: sort by vendor and amount first, then check date proximity.
    #       Bob can help you write this!
    duplicates = []
    return duplicates


def find_over_budget(aggregated: dict) -> list[dict]:
    """
    Compare monthly actuals against MONTHLY_TARGETS.
    Returns a list of dicts describing each overage.
    """
    # TODO: Implement this function.
    #       For each cost center and month, check if spend > target * 1.10 (10% threshold).
    #       Return records with keys: cost_center, month, actual, target, overage_pct
    overages = []
    return overages


# ── Step 5: Write Excel report ────────────────────────────────────────────────

def write_report(aggregated: dict, overages: list[dict], duplicates: list[tuple]):
    """Write a simple Excel report with the aggregated results."""
    if not HAS_OPENPYXL:
        print("openpyxl is not installed — cannot write Excel report.")
        print("Run:  pip install openpyxl")
        return

    OUTPUT_DIR.mkdir(exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Q1 Spend Summary"

    # Header
    ws["A1"] = "Acme GmbH — Q1 2024 Spend Summary"
    ws["A1"].font = Font(bold=True, size=14)

    headers = ["Cost Center", "January (€)", "February (€)", "March (€)", "Q1 Total (€)"]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E79")
        cell.alignment = Alignment(horizontal="center")

    row = 4
    for cc in sorted(aggregated.keys()):
        months = aggregated[cc]
        jan = months.get(1, 0)
        feb = months.get(2, 0)
        mar = months.get(3, 0)
        total = jan + feb + mar

        ws.cell(row=row, column=1, value=cc)
        ws.cell(row=row, column=2, value=round(jan, 2))
        ws.cell(row=row, column=3, value=round(feb, 2))
        ws.cell(row=row, column=4, value=round(mar, 2))
        ws.cell(row=row, column=5, value=round(total, 2))

        for col in range(2, 6):
            ws.cell(row=row, column=col).number_format = '#,##0.00'

        row += 1

    # TODO: Add a second sheet "Anomalies" that lists the overages and duplicate
    #       bookings found in steps 4. Bob can help you implement this.

    wb.save(OUTPUT_FILE)
    print(f"✓ Report saved to {OUTPUT_FILE}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("Loading transactions...")
    txns = load_transactions(TRANSACTIONS_FILE)
    print(f"  Loaded {len(txns)} transactions")

    print("Filtering approved transactions...")
    txns = filter_approved(txns)
    print(f"  {len(txns)} approved transactions remain")

    print("Validating cost centers...")
    txns = validate_cost_centers(txns)

    print("Aggregating by cost center and month...")
    aggregated = aggregate(txns)   # <-- will crash here because of the date bug

    print("Checking for anomalies...")
    duplicates = find_duplicates(txns)
    overages   = find_over_budget(aggregated)
    print(f"  Found {len(duplicates)} potential duplicates")
    print(f"  Found {len(overages)} over-budget situations")

    print("Writing report...")
    write_report(aggregated, overages, duplicates)

    print("\nDone.")


if __name__ == "__main__":
    main()
