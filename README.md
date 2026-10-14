# 💸 Personal Finance & FIRE Tracker

> Because linking 56+ spreadsheet tabs with `=ARRAYFORMULA` by hand in retirement is cruel and unusual punishment.

A fast, 100% local, zero-cloud personal finance engine and interactive Streamlit dashboard built to track early retirement (FIRE) burn rates, plan deficits, and multi-year spending without paying a dime in SaaS fees.

---

### 🚀 What Is This?

Years of spending data previously trapped across dozens of monthly Excel tabs and manual Java utilities, now liberated into a high-performance **DuckDB** single-stream ledger.

* **Local & Private**: No cloud sync, no subscription fees, no Plaid API breakages. Your data stays on your Mac.
* **Auto Sign Taming**: Banks that export expenses as positive numbers (looking at you, SoFi) are auto-detected and normalized.
* **Smart Deduplication**: Catches statement cycle bookend overlaps and pending-to-posted date drift without dropping genuine recurring charges.
* **Tab 2026 FIRE Accounting**:
  * **Core Living Baseline**: Averages Groceries, Utilities, Data, and Discretionary spend over the months elapsed to compute your true monthly operating deficit/surplus.
  * **Periodic Annual Buckets**: Tracks lumpy annual targets (Vacations, Repairs) against full-year budgets to project your annual FIRE deficit.
  * **Capital Exclusions**: Keeps major capital upgrades (e.g. 200A electric panels, sheds, fireplaces) in their own lane so they don't artificially blow up your baseline retirement budget.
* **Cash & Manual Logging**: Instant UI forms to log cash spending and side-reimbursements on the fly.

---

### 🛠️ Tech Stack

* **Language**: Python 3.9+
* **Database**: [DuckDB](https://duckdb.org/) (sub-millisecond local analytical queries)
* **Dashboard**: [Streamlit](https://streamlit.io/)
* **Config**: YAML (`config/rules.yaml`) for regex rules, budget targets, and institution aliases
* **Testing**: Pytest (12 automated unit tests)

---

### ⚡ Quick Start

```bash
# 1. Fire up the dashboard
./run.sh

# 2. Or run directly with virtual environment
.venv/bin/streamlit run app.py
```
Open **`http://localhost:8501`** in your browser.

#### Running Tests
```bash
.venv/bin/pytest -v
```

---

### 📂 Repository Anatomy

* `app.py`: Streamlit 3-tab UI (FIRE Analytics & Deficits, Unified Searchable Ledger, Ingestion Studio).
* `engine.py`: Normalization, regex categorization, bookend deduplication, and DuckDB repository.
* `migrate_history.py`: One-click ingestion script extracting historical sheets into DuckDB and CSV.
* `config/rules.yaml`: User-configurable budget plans, column aliases, and exclusion patterns.
* `data/`: Local DuckDB database (`spending.duckdb`) and flat mirror backup (`spending_ledger.csv`).
