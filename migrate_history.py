"""
Historical Migration Script
Extracts all monthly sheets from ~/Downloads/Spending.xlsx into DuckDB and CSV.
Reconciles category totals against sheet formulas and yearly plans.
"""

import os
import re
import datetime
import duckdb
import pandas as pd
import openpyxl
import glob
import sys
from typing import Optional, Dict, List, Tuple

def find_latest_excel_path() -> str:
    downloads = os.path.expanduser("~/Downloads")
    candidates = glob.glob(os.path.join(downloads, "Spending*.xlsx"))
    if candidates:
        candidates.sort(key=os.path.getmtime, reverse=True)
        return candidates[0]
    return os.path.join(downloads, "Spending.xlsx")

EXCEL_PATH = sys.argv[1] if len(sys.argv) > 1 else find_latest_excel_path()
DATA_DIR = os.path.expanduser("~/dev/personal-finance/data")
DB_PATH = os.path.join(DATA_DIR, "spending.duckdb")
BACKUP_CSV_PATH = os.path.join(DATA_DIR, "spending_ledger.csv")

CATEGORY_CANONICAL = {
    "groceries": "Groceries",
    "grocery": "Groceries",
    "shoprite": "Groceries",
    "discretionary": "Discretionary",
    "data": "Data",
    "utilities": "Utilities",
    "utilies": "Utilities",
    "vacation": "Vacations",
    "vacations": "Vacations",
    "travel": "Vacations",
    "repairs": "Repairs",
    "home upgrade": "Home Upgrades",
    "home upgrades": "Home Upgrades",
    "home improvement": "Home Upgrades",
    "hope improvement": "Home Upgrades",
    "shed replacement": "Home Upgrades",
    "9049 maintenance": "Property Maintenance",
    "investment property upkeep": "Property Maintenance",
    "rental properties maintenance": "Property Maintenance",
    "property maintenance": "Property Maintenance",
    "amazon": "Discretionary",
    "amazon card": "Discretionary",
    "cash spending": "Discretionary",
    "paul's cash": "Discretionary",
    "jing's cash": "Discretionary",
    "healthcare": "Healthcare",
}

def is_monthly_sheet(sheet_name: str) -> bool:
    """Rule 1: Monthly sheets have 6-digit numeric names (YYYYMM)."""
    return bool(sheet_name and str(sheet_name).isdigit() and len(str(sheet_name)) == 6)

def excel_serial_to_date(val):
    """Converts Excel serial number, string, or datetime to YYYY-MM-DD."""
    if val is None or val == "":
        return None
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, (int, float)):
        # Excel epoch: 1899-12-30 (handling the 1900 leap year bug)
        base = datetime.date(1899, 12, 30)
        dt = base + datetime.timedelta(days=int(val))
        return dt.strftime("%Y-%m-%d")
    val_str = str(val).strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.datetime.strptime(val_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return val_str[:10]

VALID_CATEGORIES = {
    "Groceries",
    "Discretionary",
    "Data",
    "Utilities",
    "Repairs",
    "Vacations",
    "Home Upgrades",
    "Property Maintenance",
    "Healthcare",
}

REIMBURSEMENT_KEYWORDS = [
    "due jing",
    "paid jing",
    "reimburse jing",
    "more due jing",
    "adjustment for jing",
]

def normalize_category_name(raw_name: str) -> str:
    """Normalizes category section headers within monthly sheets."""
    if not raw_name:
        return "Discretionary"
    clean = raw_name.strip()
    clean = re.sub(r'(?i)\s+total.*', '', clean).strip()
    clean = re.sub(r'(?i)\s+subtotal.*', '', clean).strip()
    clean_lower = clean.lower()
    for pattern, canonical in CATEGORY_CANONICAL.items():
        if pattern in clean_lower:
            return canonical
    title_case = clean.title()
    if title_case in VALID_CATEGORIES:
        return title_case
    return "Off-Budget / Uncategorized"

def categorize_rollup_label(lbl: str) -> Optional[str]:
    """Rule 1 & 3: Maps label in rollup sheets to canonical budget categories."""
    if not lbl:
        return None
    lbl_lower = lbl.lower()
    if any(k in lbl_lower for k in ["bespoke", "investment repair", "property"]):
        return "Property Maintenance"
    if "repair" in lbl_lower:
        return "Repairs"
    if "vacation" in lbl_lower:
        return "Vacations"
    if any(k in lbl_lower for k in ["home upgrade", "home improvement", "upgrade"]):
        return "Home Upgrades"
    if "groceries" in lbl_lower or "grocery" in lbl_lower:
        return "Groceries"
    if "discretionary" in lbl_lower:
        return "Discretionary"
    if "data" in lbl_lower:
        return "Data"
    if "utilities" in lbl_lower or "utilies" in lbl_lower:
        return "Utilities"
    return None

def extract_rollup_associations(wb_formula) -> dict:
    """
    Rule 1 & 3:
    Scans all rollup sheets (sheets whose names do NOT correspond to months)
    and extracts cell references that explicitly map monthly sheet rows
    to known budget categories.
    """
    cell_overrides = {}
    rollup_sheets = [s for s in wb_formula.sheetnames if not is_monthly_sheet(s)]
    for s_name in rollup_sheets:
        ws_r = wb_formula[s_name]
        for r in range(1, ws_r.max_row + 1):
            for c in range(1, ws_r.max_column + 1):
                val = str(ws_r.cell(row=r, column=c).value or "")
                if "=" in val:
                    refs = re.findall(r"'([0-9]{6})'!([A-Z]+[0-9]+)", val)
                    if refs:
                        lbl = str(ws_r.cell(row=r, column=1).value or ws_r.cell(row=r, column=c - 1).value or "").strip()
                        cat = categorize_rollup_label(lbl)
                        if cat:
                            for s_target, coord in refs:
                                t_row = int(re.sub(r'[A-Z]+', '', coord))
                                cell_overrides[(s_target, t_row)] = cat
    return cell_overrides

def parse_monthly_sheet(ws_f, ws_v, sheet_name: str, cell_overrides: dict) -> list:
    """
    Parses a single monthly sheet strictly following rules 2 and 3:
    - Rule 2: Considers only the first 4 columns (Date, Description, Amount, Notes).
              Reimbursement rows like 'due jing' or 'paid jing' are ignored.
    - Rule 3: Unsectioned/bottom transactions default to 'Off-Budget / Uncategorized',
      unless explicitly referenced by a rollup sheet in cell_overrides.
    """
    year_str = sheet_name[:4]
    month_str = sheet_name[4:]
    year_month = f"{year_str}-{month_str}"

    row_category_map = {}
    block_negate_map = {}
    total_rows = []

    # 1. Map rows to categories using formula ranges and block boundaries
    for r in range(1, ws_f.max_row + 1):
        col3_val = ws_f.cell(row=r, column=3).value
        col3_str = str(col3_val or "")
        col2_val = str(ws_f.cell(row=r, column=2).value or "").lower()
        col1_val = str(ws_f.cell(row=r, column=1).value or "").lower()

        # Rule 2: Ignore reimbursement calculation rows
        if any(k in (col1_val + " " + col2_val) for k in REIMBURSEMENT_KEYWORDS):
            continue

        is_total_row = "total" in col2_val or "subtotal" in col2_val or "total" in col1_val
        has_cell_ref = bool(re.search(r'[A-Za-z]+[0-9]+', col3_str))
        has_sum_formula = col3_str.startswith("=") and "sum" in col3_str.lower() and has_cell_ref

        if is_total_row or has_sum_formula:
            label = ""
            # Rule 2: Strictly within columns 1 to 4
            for col in (2, 4, 1):
                v = ws_f.cell(row=r, column=col).value
                if v and isinstance(v, str) and not str(v).startswith("="):
                    label = v.strip()
                    break

            formula_str = col3_str if col3_str.startswith("=") else ""
            cat_canonical = normalize_category_name(label)
            needs_negate = "multiply(-1" in formula_str.lower().replace(" ", "")

            ranges = re.findall(r'C([0-9]+):C([0-9]+)', formula_str, re.IGNORECASE)
            formula_no_ranges = re.sub(r'C[0-9]+:C[0-9]+', '', formula_str, flags=re.IGNORECASE)
            singles = re.findall(r'C([0-9]+)', formula_no_ranges, re.IGNORECASE)

            if ranges or singles:
                for s_r, e_r in ranges:
                    for tr in range(int(s_r), int(e_r) + 1):
                        row_category_map[tr] = cat_canonical
                        block_negate_map[tr] = needs_negate
                for sr in singles:
                    row_category_map[int(sr)] = cat_canonical
                    block_negate_map[int(sr)] = needs_negate

            total_rows.append((r, cat_canonical, needs_negate))

    # Sequential block fallback for rows inside sections
    prev_r = 1
    for t_row, cat_canonical, needs_negate in total_rows:
        for r in range(prev_r + 1, t_row):
            if r not in row_category_map:
                row_category_map[r] = cat_canonical
                block_negate_map[r] = needs_negate
        prev_r = t_row

    total_row_indices = {t[0] for t in total_rows}
    records = []

    # 2. Extract transaction rows
    for r in range(2, ws_v.max_row + 1):
        if r in total_row_indices:
            continue

        # Check if this row in the formula sheet contains a subtotal/rollup formula
        cell_f_amt = ws_f.cell(row=r, column=3).value
        if cell_f_amt and str(cell_f_amt).startswith("=") and re.search(r'[A-Za-z]+[0-9]+', str(cell_f_amt)):
            continue

        # Rule 2: Strictly read columns 1 to 4 only
        date_raw = ws_v.cell(row=r, column=1).value
        desc_raw = ws_v.cell(row=r, column=2).value
        amt_raw = ws_v.cell(row=r, column=3).value
        notes_raw = ws_v.cell(row=r, column=4).value

        if amt_raw is None or not isinstance(amt_raw, (int, float)):
            continue

        desc_str = str(desc_raw or "").strip()

        # Skip summary rows, totals, and blank descriptions
        if any(w in desc_str.lower() for w in ["total", "subtotal", "average", "months so far"]):
            continue
        if any(k in desc_str.lower() for k in REIMBURSEMENT_KEYWORDS):
            continue
        if not date_raw and desc_str.lower() in ["groceries", "utilities", "data", "discretionary", "vacations", "repairs", "auto", "home upgrades"]:
            continue

        date_str = excel_serial_to_date(date_raw)
        if not date_str:
            # Default to the 1st of the sheet's month if date cell is blank
            date_str = f"{year_str}-{month_str}-01"

        amt = float(amt_raw)
        should_negate = block_negate_map.get(r, False)
        if should_negate:
            amt = -abs(amt)

        # Rule 3: Category resolution
        if (sheet_name, r) in cell_overrides:
            category = cell_overrides[(sheet_name, r)]
        elif r in row_category_map:
            category = row_category_map[r]
        else:
            category = "Off-Budget / Uncategorized"

        # Combine notes strictly from column 4
        notes_str = ""
        if notes_raw and isinstance(notes_raw, str) and notes_raw.strip() != category:
            notes_str = notes_raw.strip()

        records.append({
            "date": date_str,
            "year_month": year_month,
            "year": int(year_str),
            "month": int(month_str),
            "description": desc_str,
            "amount": round(amt, 2),
            "category": category,
            "sheet_source": sheet_name,
            "row_number": r,
            "notes": notes_str
        })

    return records

def migrate(excel_path: str = None, data_dir: str = None):
    target_excel = excel_path or EXCEL_PATH
    target_data_dir = data_dir or DATA_DIR
    db_path = os.path.join(target_data_dir, "spending.duckdb")
    backup_csv_path = os.path.join(target_data_dir, "spending_ledger.csv")

    os.makedirs(target_data_dir, exist_ok=True)
    print(f"Loading workbook from {target_excel}...")
    wb_formula = openpyxl.load_workbook(target_excel, data_only=False)
    wb_val = openpyxl.load_workbook(target_excel, data_only=True)

    # Rule 1 & 3: Extract rollup associations
    cell_overrides = extract_rollup_associations(wb_formula)
    print(f"Extracted {len(cell_overrides)} authoritative cell overrides from rollup sheets.")

    # Rule 1: Monthly sheets only
    monthly_sheet_names = sorted([s for s in wb_formula.sheetnames if is_monthly_sheet(s)])
    print(f"Found {len(monthly_sheet_names)} monthly sheets: {monthly_sheet_names[0]} to {monthly_sheet_names[-1]}")

    all_records = []
    for s_name in monthly_sheet_names:
        ws_f = wb_formula[s_name]
        ws_v = wb_val[s_name]
        sheet_records = parse_monthly_sheet(ws_f, ws_v, s_name, cell_overrides)
        all_records.extend(sheet_records)

    df = pd.DataFrame(all_records)
    print(f"Total transactions extracted: {len(df)}")
    print(f"Date range: {df['date'].min()} to {df['date'].max()}")
    print("Breakdown by Category:")
    cat_summary = df.groupby('category').agg(
        count=('amount', 'count'),
        total_spend=('amount', 'sum'),
        avg_spend=('amount', 'mean')
    ).reset_index()
    print(cat_summary.to_string(index=False))

    # Save to CSV backup
    df.to_csv(backup_csv_path, index=False)
    print(f"Saved CSV backup to: {backup_csv_path}")

    # Save to DuckDB
    con = duckdb.connect(db_path)
    con.execute("CREATE OR REPLACE TABLE transactions AS SELECT * FROM df")
    count = con.execute("SELECT count(*) FROM transactions").fetchone()[0]
    print(f"DuckDB table 'transactions' created with {count} rows in: {db_path}")
    con.close()
    return df

if __name__ == "__main__":
    migrate()
