"""
Core Personal Finance Engine
Handles column normalization, regex categorization, bookend deduplication,
and DuckDB ledger operations.
"""

import os
import re
import datetime
from typing import Dict, List, Tuple, Optional, Any
import yaml
import duckdb
import pandas as pd


class RuleConfig:
    def __init__(self, config_path: str):
        self.config_path = config_path
        with open(config_path, "r") as f:
            self.raw_config = yaml.safe_load(f)

        self.column_aliases: Dict[str, List[str]] = self.raw_config.get("column_aliases", {})
        self.default_category: str = self.raw_config.get("default_category", "Discretionary")
        self.categories: Dict[str, List[str]] = self.raw_config.get("categories", {})
        self.budget_plans: Dict[str, Dict[str, float]] = self.raw_config.get("budget_plans", {})
        self.exclude_patterns: List[str] = self.raw_config.get("exclude_patterns", [])

        # Precompile regexes
        self.compiled_regexes: List[Tuple[str, re.Pattern]] = []
        for cat_name, patterns in self.categories.items():
            for p in patterns:
                self.compiled_regexes.append((cat_name, re.compile(p, re.IGNORECASE)))

        self.compiled_exclude_regexes: List[re.Pattern] = [
            re.compile(p, re.IGNORECASE) for p in self.exclude_patterns
        ]

    def is_excluded(self, description: str) -> bool:
        if not description:
            return False
        clean_desc = str(description).strip()
        for pattern in self.compiled_exclude_regexes:
            if pattern.search(clean_desc):
                return True
        return False

    def categorize(self, description: str) -> str:
        if not description:
            return self.default_category
        clean_desc = str(description).strip()
        for cat_name, pattern in self.compiled_regexes:
            if pattern.search(clean_desc):
                return cat_name
        return self.default_category


def clean_merchant_string(desc: str) -> str:
    """Simplifies merchant description for robust comparison."""
    if not desc:
        return ""
    # Remove dates, phone numbers, state codes, reference IDs
    s = str(desc).lower()
    s = re.sub(r'\b\d{4}-\d{2}-\d{2}\b', '', s)
    s = re.sub(r'[\d\-#*]+', ' ', s)
    s = re.sub(r'\b(in|des|id|indn|co|ppd|web)\b', ' ', s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip()


class TransactionNormalizer:
    def __init__(self, rule_config: RuleConfig):
        self.config = rule_config

    def resolve_column(self, actual_columns: List[str], target_field: str) -> Optional[str]:
        aliases = self.config.column_aliases.get(target_field, [])
        all_candidates = [target_field] + aliases
        for col in actual_columns:
            for candidate in all_candidates:
                if col.strip().lower() == candidate.lower():
                    return col
        return None

    def parse_date(self, val: Any) -> Optional[str]:
        if val is None or pd.isna(val) or val == "":
            return None
        if isinstance(val, (datetime.date, datetime.datetime)):
            return val.strftime("%Y-%m-%d")
        if isinstance(val, (int, float)):
            try:
                import openpyxl.utils.datetime as o_dt
                return o_dt.from_excel(val).strftime("%Y-%m-%d")
            except Exception:
                base = datetime.date(1899, 12, 30)
                return (base + datetime.timedelta(days=int(val))).strftime("%Y-%m-%d")
        val_str = str(val).strip()
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%Y/%m/%d", "%d/%m/%Y"):
            try:
                return datetime.datetime.strptime(val_str, fmt).strftime("%Y-%m-%d")
            except ValueError:
                pass
        # Fallback to dateutil if present
        try:
            from dateutil.parser import parse
            return parse(val_str).strftime("%Y-%m-%d")
        except Exception:
            return None

    def parse_amount(self, val: Any) -> float:
        if val is None or pd.isna(val) or val == "":
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).strip().replace("$", "").replace(",", "")
        # Handle accounting parenthesis: (12.34) -> -12.34
        if s.startswith("(") and s.endswith(")"):
            s = f"-{s[1:-1]}"
        try:
            return float(s)
        except ValueError:
            return 0.0

    def normalize_dataframe(
        self,
        df: pd.DataFrame,
        source_name: str = "",
        invert_signs: Optional[bool] = None
    ) -> pd.DataFrame:
        cols = list(df.columns)
        date_col = self.resolve_column(cols, "date")
        desc_col = self.resolve_column(cols, "description")
        amt_col = self.resolve_column(cols, "amount")
        cat_col = self.resolve_column(cols, "category")

        if not date_col or not desc_col or not amt_col:
            raise ValueError(
                f"Missing required columns in CSV. Found: {cols}. "
                f"Resolved: date={date_col}, description={desc_col}, amount={amt_col}"
            )

        records = []
        for _, row in df.iterrows():
            d = self.parse_date(row[date_col])
            desc = str(row[desc_col] or "").strip()
            amt = self.parse_amount(row[amt_col])
            if not d or not desc or amt == 0.0:
                continue

            existing_cat = str(row[cat_col]).strip() if cat_col and pd.notna(row[cat_col]) else ""
            is_excl = self.config.is_excluded(desc)
            if existing_cat:
                category = existing_cat
            elif is_excl:
                category = "Excluded"
            else:
                category = self.config.categorize(desc)

            dt = datetime.datetime.strptime(d, "%Y-%m-%d")
            records.append({
                "include": not is_excl,
                "date": d,
                "year_month": dt.strftime("%Y-%m"),
                "year": dt.year,
                "month": dt.month,
                "description": desc,
                "amount": round(amt, 2),
                "category": category,
                "sheet_source": source_name,
                "notes": "Auto-excluded by rule" if is_excl else ""
            })

        norm_df = pd.DataFrame(records)
        if norm_df.empty:
            return norm_df

        # Expense sign handling:
        # Standard convention across our ledger: Expenses are NEGATIVE, Credits/Refunds are POSITIVE.
        if invert_signs is True:
            norm_df["amount"] = -norm_df["amount"]
        elif invert_signs is False:
            pass  # Leave amounts exactly as provided
        else:
            # Auto-detection:
            # 1. Check if SoFi or similar inverted export in source name
            is_sofi = "sofi" in source_name.lower()
            # 2. Check if the majority (>60%) of amounts are positive
            positive_count = (norm_df["amount"] > 0).sum()
            total_count = len(norm_df)
            majority_positive = total_count > 0 and (positive_count / total_count) > 0.60
            
            if is_sofi or majority_positive:
                norm_df["amount"] = -norm_df["amount"]

        return norm_df


class DeduplicationEngine:
    def __init__(self, bookend_day_window: int = 3):
        self.window = bookend_day_window

    def deduplicate(
        self, incoming_df: pd.DataFrame, existing_df: pd.DataFrame
    ) -> Tuple[pd.DataFrame, pd.DataFrame, List[Dict[str, Any]]]:
        """
        Deduplicates incoming transactions against the existing ledger.
        Returns:
          - new_txs: rows that are completely new
          - exact_dups: rows that are exact matches (same date, same amount, same description)
          - bookend_suspects: rows that closely match within +/- window days with identical amount
        """
        if incoming_df.empty:
            return incoming_df, pd.DataFrame(), []

        if existing_df.empty:
            return incoming_df, pd.DataFrame(), []

        new_rows = []
        exact_dups = []
        bookend_suspects = []

        # Prepare existing lookup structures
        # 1. Exact lookup: (date, amount, cleaned_desc)
        existing_exact_set = set()
        for _, row in existing_df.iterrows():
            key = (row["date"], round(float(row["amount"]), 2), clean_merchant_string(row["description"]))
            existing_exact_set.add(key)

        # 2. Existing records with date objects for window search
        existing_records = []
        for _, row in existing_df.iterrows():
            try:
                dt = datetime.datetime.strptime(row["date"], "%Y-%m-%d").date()
                existing_records.append({
                    "date": dt,
                    "date_str": row["date"],
                    "amount": round(float(row["amount"]), 2),
                    "description": row["description"],
                    "clean_desc": clean_merchant_string(row["description"]),
                    "category": row.get("category", "")
                })
            except Exception:
                continue

        for _, inc_row in incoming_df.iterrows():
            inc_date_str = inc_row["date"]
            inc_amt = round(float(inc_row["amount"]), 2)
            inc_desc = inc_row["description"]
            inc_clean = clean_merchant_string(inc_desc)

            # Check 1: Exact match
            exact_key = (inc_date_str, inc_amt, inc_clean)
            if exact_key in existing_exact_set:
                exact_dups.append(inc_row.to_dict())
                continue

            # Check 2: Bookend / Window match (+/- window days with matching amount)
            inc_date = datetime.datetime.strptime(inc_date_str, "%Y-%m-%d").date()
            is_bookend = False
            matched_candidate = None

            for ex in existing_records:
                if abs(ex["amount"] - inc_amt) < 0.01:
                    day_diff = abs((inc_date - ex["date"]).days)
                    if 0 < day_diff <= self.window:
                        # Check description similarity
                        # If one cleaned string is a substring of the other or word tokens overlap
                        inc_tokens = set(inc_clean.split())
                        ex_tokens = set(ex["clean_desc"].split())
                        if inc_clean == ex["clean_desc"] or (inc_tokens and inc_tokens.issubset(ex_tokens)) or (ex_tokens and ex_tokens.issubset(inc_tokens)):
                            is_bookend = True
                            matched_candidate = ex
                            matched_candidate["day_diff"] = day_diff
                            break

            if is_bookend and matched_candidate:
                bookend_suspects.append({
                    "incoming": inc_row.to_dict(),
                    "matched_existing": matched_candidate,
                    "reason": f"Matches existing transaction within {matched_candidate['day_diff']} days with identical amount (${inc_amt:.2f})"
                })
            else:
                new_rows.append(inc_row.to_dict())

        return pd.DataFrame(new_rows), pd.DataFrame(exact_dups), bookend_suspects


class LedgerRepository:
    def __init__(self, db_path: str):
        self.db_path = db_path
        os.makedirs(os.path.dirname(db_path), exist_ok=True)
        self._ensure_schema()

    def _get_connection(self):
        return duckdb.connect(self.db_path)

    def _ensure_schema(self):
        con = self._get_connection()
        con.execute("""
            CREATE TABLE IF NOT EXISTS transactions (
                date VARCHAR,
                year_month VARCHAR,
                year INTEGER,
                month INTEGER,
                description VARCHAR,
                amount DOUBLE,
                category VARCHAR,
                sheet_source VARCHAR,
                row_number BIGINT,
                notes VARCHAR
            )
        """)
        con.close()

    def get_all_transactions(self) -> pd.DataFrame:
        con = self._get_connection()
        df = con.execute("SELECT * FROM transactions ORDER BY date DESC").df()
        con.close()
        return df

    def append_transactions(self, df: pd.DataFrame, backup_csv_path: Optional[str] = None):
        if df.empty:
            return
        con = self._get_connection()
        # Add row_number if missing
        if "row_number" not in df.columns:
            df["row_number"] = None
        con.register("incoming_batch", df)
        con.execute("""
            INSERT INTO transactions (date, year_month, year, month, description, amount, category, sheet_source, row_number, notes)
            SELECT date, year_month, year, month, description, amount, category, sheet_source, row_number, notes
            FROM incoming_batch
        """)
        con.close()

        if backup_csv_path:
            full_df = self.get_all_transactions()
            full_df.to_csv(backup_csv_path, index=False)

    def add_single_transaction(
        self,
        date_str: str,
        description: str,
        amount: float,
        category: str,
        notes: str = "",
        source: str = "Manual Entry",
        backup_csv_path: Optional[str] = None
    ):
        """Adds a single manual/cash transaction to the ledger and CSV backup."""
        dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        record = pd.DataFrame([{
            "date": date_str,
            "year_month": dt.strftime("%Y-%m"),
            "year": dt.year,
            "month": dt.month,
            "description": description.strip(),
            "amount": round(amount, 2),
            "category": category.strip(),
            "sheet_source": source,
            "row_number": None,
            "notes": notes.strip()
        }])
        self.append_transactions(record, backup_csv_path=backup_csv_path)

    def update_all_transactions(self, df: pd.DataFrame, backup_csv_path: Optional[str] = None):
        """Overwrites transactions table with edited DataFrame (for manual inline edits in UI)."""
        con = self._get_connection()
        save_df = df.copy()
        # Drop temporary UI columns
        drop_cols = [c for c in ["date_dt", "include"] if c in save_df.columns]
        if drop_cols:
            save_df = save_df.drop(columns=drop_cols)
        con.register("updated_ledger_view", save_df)
        con.execute("CREATE OR REPLACE TABLE transactions AS SELECT * FROM updated_ledger_view")
        con.close()
        if backup_csv_path:
            save_df.to_csv(backup_csv_path, index=False)

    def get_fire_summary(self, budget_plans: Dict[str, Dict[str, float]]) -> Dict[str, Any]:
        """Calculates monthly and annual deficit vs plan matching your FIRE spreadsheet."""
        con = self._get_connection()
        df = con.execute("""
            SELECT year_month, category, round(sum(amount), 2) as spend
            FROM transactions
            GROUP BY year_month, category
            ORDER BY year_month, category
        """).df()
        con.close()

        monthly_plans = budget_plans.get("monthly", {})
        months = sorted(df["year_month"].unique())

        monthly_breakdowns = []
        for ym in months:
            m_df = df[df["year_month"] == ym]
            cat_spend = dict(zip(m_df["category"], m_df["spend"]))

            row_data = {"year_month": ym}
            total_monthly_spend = sum(cat_spend.values())
            row_data["total_spend"] = total_monthly_spend

            total_deficit = 0.0
            for cat, plan in monthly_plans.items():
                actual = cat_spend.get(cat, 0.0)
                # Deficit formula: -(actual + plan) when actual is negative
                # If plan is 750 and spend is -550, deficit is +(750 - 550) = +200 (surplus)
                deficit = -1 * (actual + plan)
                row_data[f"{cat}_spend"] = actual
                row_data[f"{cat}_deficit"] = deficit
                total_deficit += deficit

            row_data["monthly_deficit_vs_plan"] = total_deficit
            monthly_breakdowns.append(row_data)

        return {
            "monthly_report": pd.DataFrame(monthly_breakdowns),
            "category_totals": df
        }

    def get_yearly_fire_rollup(self, year: int, budget_plans: Dict[str, Any]) -> Dict[str, Any]:
        """
        Computes the annual FIRE rollup matching the layout and formulas of tab '2026' in the spreadsheet:
        1. Months So Far in the specified year.
        2. Baseline Monthly Categories (Groceries, Discretionary, Data, Utilities) averaged over months so far,
           compared against monthly plan, computing individual deficits and the overall Monthly Deficit vs Plan.
        3. Periodic Annual Buckets (Repairs, Vacations) comparing YTD total against full-year plan.
        4. Projected Annual Deficit vs Plan = (Monthly Deficit vs Plan * 12) + Repairs Deficit + Vacations Deficit.
        5. Capital / Non-operating Projects (Home Upgrades, Bespoke Property Maintenance, Auto, Healthcare)
           tracked separately so they do not distort operating deficit.
        """
        year = int(year)
        con = self._get_connection()
        months_so_far = con.execute("SELECT count(distinct year_month) FROM transactions WHERE year = ?", [year]).fetchone()[0] or 1
        
        cat_df = con.execute("""
            SELECT category, round(sum(amount), 2) as total_spend, count(*) as tx_count
            FROM transactions
            WHERE year = ?
            GROUP BY category
            ORDER BY total_spend ASC
        """, [year]).df()
        con.close()

        cat_spend_map = dict(zip(cat_df["category"], cat_df["total_spend"]))
        cat_count_map = dict(zip(cat_df["category"], cat_df["tx_count"]))

        monthly_plans = budget_plans.get("monthly", {})
        annual_plans = budget_plans.get("annual", {})

        # 1. Baseline Monthly Categories
        baseline_rows = []
        monthly_deficit_sum = 0.0
        for cat in ["Groceries", "Discretionary", "Data", "Utilities"]:
            actual_total = cat_spend_map.get(cat, 0.0)
            actual_avg = round(actual_total / months_so_far, 2)
            plan = monthly_plans.get(cat, 0.0)
            # In Excel: =MULTIPLY(-1, SUM(B2, C2)) where actual is negative and plan is positive
            deficit = round(-1.0 * (actual_avg + plan), 2)
            monthly_deficit_sum += deficit
            status = "Surplus" if deficit <= 0 else "Deficit"
            baseline_rows.append({
                "Category": cat,
                "Actual Total": actual_total,
                "Actual Avg / Mo": actual_avg,
                "Plan / Mo": plan,
                "Deficit / Mo": deficit,
                "Status": status
            })

        # 2. Periodic Annual Buckets
        annual_rows = []
        annual_periodic_deficit_sum = 0.0
        for cat in ["Repairs", "Vacations"]:
            actual_total = cat_spend_map.get(cat, 0.0)
            plan = annual_plans.get(cat, 0.0)
            # In Excel: =MULTIPLY(-1, SUM(B10, C10))
            deficit = round(-1.0 * (actual_total + plan), 2)
            annual_periodic_deficit_sum += deficit
            status = "Surplus" if deficit <= 0 else "Deficit"
            annual_rows.append({
                "Category": cat,
                "Actual YTD Total": actual_total,
                "Annual Plan": plan,
                "Deficit vs Plan": deficit,
                "Status": status
            })

        # Projected Annual Deficit = (monthly deficit * 12) + periodic deficits
        projected_annual_deficit = round((monthly_deficit_sum * 12) + annual_periodic_deficit_sum, 2)

        # 3. Capital / Non-operating Projects
        excluded_cats = set(monthly_plans.keys()) | set(annual_plans.keys())
        capital_rows = []
        total_capital = 0.0
        for cat, tot in cat_spend_map.items():
            if cat not in excluded_cats:
                capital_rows.append({
                    "Category": cat,
                    "Total Spent": tot,
                    "Transactions": cat_count_map.get(cat, 0)
                })
                total_capital += tot

        return {
            "year": year,
            "months_so_far": months_so_far,
            "baseline_df": pd.DataFrame(baseline_rows),
            "monthly_deficit_vs_plan": round(monthly_deficit_sum, 2),
            "annual_periodic_df": pd.DataFrame(annual_rows),
            "annual_periodic_deficit": round(annual_periodic_deficit_sum, 2),
            "projected_annual_deficit": projected_annual_deficit,
            "capital_df": pd.DataFrame(capital_rows),
            "total_capital_spend": round(total_capital, 2)
        }
