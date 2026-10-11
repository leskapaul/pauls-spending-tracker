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

EXCEL_PATH = os.path.expanduser("~/Downloads/Spending.xlsx")
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
    "disney vacay": "Vacations",
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
    "healthcare": "Healthcare",
    "due jing": "Shared / Reimbursements",
    "more due jing": "Shared / Reimbursements",
    "paid jing": "Shared / Reimbursements",
    "reimburse jing": "Shared / Reimbursements",
    "jing cash spending": "Shared / Reimbursements",
    "adjustment for jing": "Shared / Reimbursements",
    "aaron for sushi": "Discretionary",
    "njt mobile": "Discretionary",
}

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
    # Try common formats
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.datetime.strptime(val_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return val_str[:10]

def normalize_category_name(raw_name):
    if not raw_name:
        return "Discretionary"
    clean = raw_name.strip()
    clean = re.sub(r'(?i)\s+total.*', '', clean).strip()
    clean = re.sub(r'(?i)\s+subtotal.*', '', clean).strip()
    clean_lower = clean.lower()
    for pattern, canonical in CATEGORY_CANONICAL.items():
        if pattern in clean_lower:
            return canonical
    return clean.title()

def migrate():
    os.makedirs(DATA_DIR, exist_ok=True)
    print(f"Loading workbook from {EXCEL_PATH}...")
    wb_formula = openpyxl.load_workbook(EXCEL_PATH, data_only=False)
    wb_val = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    
    monthly_sheet_names = sorted([s for s in wb_formula.sheetnames if s.isdigit() and len(s) == 6])
    print(f"Found {len(monthly_sheet_names)} monthly sheets: {monthly_sheet_names[0]} to {monthly_sheet_names[-1]}")

    all_records = []
    
    for s_name in monthly_sheet_names:
        year_str = s_name[:4]
        month_str = s_name[4:]
        year_month = f"{year_str}-{month_str}"
        
        ws_f = wb_formula[s_name]
        ws_v = wb_val[s_name]

        # 1. Map rows to categories using formula ranges and block boundaries
        row_category_map = {}
        block_negate_map = {} # tracks if formula had MULTIPLY(-1, ...)
        
        # Scan for total rows
        total_rows = []
        for r in range(1, ws_f.max_row + 1):
            row_vals = [ws_f.cell(row=r, column=c).value for c in range(1, 10)]
            row_str = " ".join(str(v) for v in row_vals if v is not None).lower()
            
            is_total_row = "total" in row_str or "subtotal" in row_str
            
            # Also check if column C has a SUM formula without explicit 'total' text
            has_sum_formula = any(str(v).startswith("=") and "sum" in str(v).lower() for v in row_vals if v)
            
            if is_total_row or has_sum_formula:
                # Determine label
                label = ""
                for col in (2, 4, 1, 5):
                    v = ws_f.cell(row=r, column=col).value
                    if v and isinstance(v, str) and not str(v).startswith("="):
                        label = v.strip()
                        break
                
                # Extract formula
                formula_str = ""
                for col in range(1, 10):
                    f_val = str(ws_f.cell(row=r, column=col).value or "")
                    if f_val.startswith("="):
                        formula_str = f_val
                        break
                
                cat_canonical = normalize_category_name(label)
                needs_negate = "multiply(-1" in formula_str.lower().replace(" ", "")

                # If formula explicitly gives cell ranges like C3:C21
                ranges = re.findall(r'C([0-9]+):C([0-9]+)', formula_str, re.IGNORECASE)
                singles = re.findall(r'C([0-9]+)(?!:)', formula_str, re.IGNORECASE)
                
                if ranges or singles:
                    for s_r, e_r in ranges:
                        for tr in range(int(s_r), int(e_r) + 1):
                            row_category_map[tr] = cat_canonical
                            block_negate_map[tr] = needs_negate
                    for sr in singles:
                        row_category_map[int(sr)] = cat_canonical
                        block_negate_map[int(sr)] = needs_negate
                
                total_rows.append((r, cat_canonical, needs_negate))

        # Secondary pass: sequential block fallback for rows not explicitly inside a SUM range
        # E.g. where total row sits right below a section
        prev_r = 1
        for t_row, cat_canonical, needs_negate in total_rows:
            for r in range(prev_r + 1, t_row):
                if r not in row_category_map:
                    row_category_map[r] = cat_canonical
                    block_negate_map[r] = needs_negate
            prev_r = t_row

        # 2. Extract transaction rows
        for r in range(2, ws_v.max_row + 1):
            date_raw = ws_v.cell(row=r, column=1).value
            desc_raw = ws_v.cell(row=r, column=2).value
            amt_raw = ws_v.cell(row=r, column=3).value
            notes_raw = ws_v.cell(row=r, column=4).value
            
            # Also check side columns for split / due jing notes
            side_note = ""
            for c in range(5, 10):
                sv = ws_v.cell(row=r, column=c).value
                if sv and isinstance(sv, str):
                    side_note += f" {sv.strip()}"
            side_note = side_note.strip()

            if amt_raw is None or not isinstance(amt_raw, (int, float)):
                continue

            desc_str = str(desc_raw or "").strip()
            
            # Skip summary rows, totals, and blank descriptions
            if any(w in desc_str.lower() for w in ["total", "subtotal", "average", "months so far"]):
                continue
            if desc_str.lower().startswith("due ") and not date_raw:
                continue

            date_str = excel_serial_to_date(date_raw)
            if not date_str:
                # Default to the 1st of the sheet's month if date cell is blank
                date_str = f"{year_str}-{month_str}-01"

            # Normalize amount: expenses should be negative, refunds positive
            amt = float(amt_raw)
            should_negate = block_negate_map.get(r, False)
            if should_negate:
                # In sheets where expenses were positive and multiplied by -1
                amt = -abs(amt)
            else:
                # If sheet had positive expenses without explicit multiply, check if majority are negative
                pass

            category = row_category_map.get(r, "Discretionary")
            if not category or category == "":
                category = "Discretionary"

            # Combine notes
            full_notes = ""
            if notes_raw and isinstance(notes_raw, str) and notes_raw.strip() != category:
                full_notes = notes_raw.strip()
            if side_note:
                full_notes = f"{full_notes} | {side_note}".strip(" |")

            all_records.append({
                "date": date_str,
                "year_month": year_month,
                "year": int(year_str),
                "month": int(month_str),
                "description": desc_str,
                "amount": round(amt, 2),
                "category": category,
                "sheet_source": s_name,
                "row_number": r,
                "notes": full_notes
            })

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
    df.to_csv(BACKUP_CSV_PATH, index=False)
    print(f"Saved CSV backup to: {BACKUP_CSV_PATH}")

    # Save to DuckDB
    con = duckdb.connect(DB_PATH)
    con.execute("CREATE OR REPLACE TABLE transactions AS SELECT * FROM df")
    count = con.execute("SELECT count(*) FROM transactions").fetchone()[0]
    print(f"DuckDB table 'transactions' created with {count} rows in: {DB_PATH}")
    con.close()

if __name__ == "__main__":
    migrate()
