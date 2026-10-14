"""
Personal Finance & FIRE Spending Dashboard
Interactive Streamlit application for single-stream expense analysis,
FIRE plan vs. actual deficit tracking, and bank CSV ingestion with bookend deduplication.
"""

import os
import datetime
import streamlit as st
import pandas as pd
import duckdb
from engine import RuleConfig, TransactionNormalizer, DeduplicationEngine, LedgerRepository

# Page Configuration
st.set_page_config(
    page_title="FIRE Spending & Single-Stream Ledger",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS to ensure Deploy button and toolbar are completely hidden
st.markdown("""
<style>
.stDeployButton {
    display: none !important;
}
[data-testid="stToolbar"] {
    display: none !important;
}
footer {
    display: none !important;
}
</style>
""", unsafe_allow_html=True)


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config", "rules.yaml")
DB_PATH = os.path.join(BASE_DIR, "data", "spending.duckdb")
BACKUP_CSV_PATH = os.path.join(BASE_DIR, "data", "spending_ledger.csv")

@st.cache_resource
def get_engine_components():
    config = RuleConfig(CONFIG_PATH)
    normalizer = TransactionNormalizer(config)
    dedup = DeduplicationEngine(bookend_day_window=3)
    repo = LedgerRepository(DB_PATH)
    return config, normalizer, dedup, repo

config, normalizer, dedup, repo = get_engine_components()

def render_manual_transaction_form(key_prefix: str, repo: LedgerRepository, config: RuleConfig, default_date: datetime.date):
    """Renders a clean form to log manual expenses (like cash spending) directly into the ledger."""
    with st.form(f"{key_prefix}_manual_form", clear_on_submit=True):
        st.markdown("##### 💵 Log Manual / Cash Expense")
        c1, c2, c3 = st.columns([2, 3, 2])
        with c1:
            m_date = st.date_input("Date", value=default_date, key=f"{key_prefix}_date")
        with c2:
            m_desc = st.text_input("Description / Payee", placeholder="e.g. Paul Cash Spending / Farmers Market", key=f"{key_prefix}_desc")
        with c3:
            m_amt = st.number_input("Amount ($)", min_value=0.01, step=5.0, format="%.2f", key=f"{key_prefix}_amt")

        c4, c5 = st.columns([2, 5])
        with c4:
            m_type = st.radio("Type", ["Expense (-)", "Credit / Refund (+)"], horizontal=True, key=f"{key_prefix}_type")
        with c5:
            all_cat_options = sorted(list(config.categories.keys()) + ["Discretionary", "Shared / Reimbursements"])
            default_cat_idx = all_cat_options.index("Discretionary") if "Discretionary" in all_cat_options else 0
            m_cat = st.selectbox("Category", options=all_cat_options, index=default_cat_idx, key=f"{key_prefix}_cat")

        m_notes = st.text_input("Notes / Split Details (Optional)", placeholder="e.g. 10 gas, 40 brewery, 20 wood, 100 dinner", key=f"{key_prefix}_notes")

        submitted = st.form_submit_button("➕ Save Transaction to Ledger", type="primary")
        if submitted:
            if not m_desc.strip():
                st.error("Please provide a description.")
            else:
                final_amt = -abs(m_amt) if m_type.startswith("Expense") else abs(m_amt)
                repo.add_single_transaction(
                    date_str=m_date.strftime("%Y-%m-%d"),
                    description=m_desc.strip(),
                    amount=final_amt,
                    category=m_cat,
                    notes=m_notes.strip(),
                    source="Manual Entry",
                    backup_csv_path=BACKUP_CSV_PATH
                )
                st.success(f"Successfully added '{m_desc.strip()}' (${final_amt:.2f}) to {m_cat}!")
                st.cache_resource.clear()
                st.rerun()

# Sidebar
st.sidebar.title("💰 FIRE Finance")
st.sidebar.markdown("*Single-stream transaction ledger & FIRE analytics*")

# Load Ledger Data
all_df = repo.get_all_transactions()

if all_df.empty:
    st.warning("No transactions found in database. Please run migration or ingest CSVs.")
    st.stop()

# Ensure date column is datetime
all_df["date_dt"] = pd.to_datetime(all_df["date"], errors="coerce")
min_date = all_df["date_dt"].min().date()
max_date = all_df["date_dt"].max().date()

st.sidebar.markdown("---")
st.sidebar.metric("Total Transactions", f"{len(all_df):,}")
total_spend = all_df["amount"].sum()
st.sidebar.metric("All-Time Spend", f"${abs(total_spend):,.2f}")
st.sidebar.caption(f"Historical span: {min_date} to {max_date}")

# Main Tabs
tab_overview, tab_ledger, tab_ingest = st.tabs([
    "📊 FIRE Analytics & Plan Deficits",
    "📜 Single-Stream Ledger",
    "📥 Ingest & Deduplicate CSVs"
])

# -------------------------------------------------------------
# TAB 1: FIRE Analytics & Plan Deficit Tracking
# -------------------------------------------------------------
with tab_overview:
    col_t1, col_t2 = st.columns([3, 1])
    with col_t1:
        st.header("FIRE Expense Rollup & Deficit vs. Plan")
        st.markdown("Tracked identically to your Excel annual summary tabs (`2026`, `2025`, etc.).")
    with col_t2:
        available_years = sorted(all_df["year"].unique(), reverse=True)
        selected_year = st.selectbox("Select Year", available_years, index=0)

    # Get annual rollup for the selected year
    yearly_rollup = repo.get_yearly_fire_rollup(selected_year, config.budget_plans)
    
    # KPI Metrics
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    with kpi1:
        st.metric(f"Months So Far ({selected_year})", f"{yearly_rollup['months_so_far']} months")
    with kpi2:
        m_def = yearly_rollup['monthly_deficit_vs_plan']
        # Deficit sign convention: negative = surplus (favorable), positive = deficit (unfavorable)
        st.metric(
            "Monthly Deficit vs Plan",
            f"${abs(m_def):,.2f}",
            delta=f"{'-' if m_def < 0 else '+'}${abs(m_def):,.2f} ({'Surplus' if m_def <= 0 else 'Deficit'})",
            delta_color="inverse"
        )
    with kpi3:
        ann_def = yearly_rollup['projected_annual_deficit']
        st.metric(
            "Projected Annual Deficit",
            f"${abs(ann_def):,.2f}",
            delta=f"{'-' if ann_def < 0 else '+'}${abs(ann_def):,.2f} ({'Surplus' if ann_def <= 0 else 'Deficit'})",
            delta_color="inverse"
        )
    with kpi4:
        cap_spent = abs(yearly_rollup['total_capital_spend'])
        st.metric("Capital Projects YTD", f"${cap_spent:,.2f}")

    st.markdown("---")

    # Section 1: Baseline Monthly Categories
    st.subheader(f"1. Baseline Monthly Averages (Year {selected_year})")
    st.caption("Averaged across the months elapsed so far this year. Deficit formula: `-(Actual + Plan)`.")
    base_df = yearly_rollup["baseline_df"].copy()
    st.dataframe(
        base_df.style.format({
            "Actual Total": "${:,.2f}",
            "Actual Avg / Mo": "${:,.2f}",
            "Plan / Mo": "${:,.2f}",
            "Deficit / Mo": "${:,.2f}",
        }),
        use_container_width=True
    )
    if yearly_rollup['monthly_deficit_vs_plan'] <= 0:
        st.success(f"**Monthly Deficit vs Plan**: **-${abs(yearly_rollup['monthly_deficit_vs_plan']):,.2f}** (Under budget / surplus across baseline categories)")
    else:
        st.error(f"**Monthly Deficit vs Plan**: **+${yearly_rollup['monthly_deficit_vs_plan']:,.2f}** (Over budget across baseline categories)")

    # Section 2: Periodic Annual Buckets
    st.subheader("2. Periodic Annual Buckets (Year-to-Date vs. Full Year Plan)")
    st.caption("Tracked as cumulative annual totals. Formula: `-(Actual Total + Annual Plan)`.")
    ann_df = yearly_rollup["annual_periodic_df"].copy()
    st.dataframe(
        ann_df.style.format({
            "Actual YTD Total": "${:,.2f}",
            "Annual Plan": "${:,.2f}",
            "Deficit vs Plan": "${:,.2f}",
        }),
        use_container_width=True
    )
    
    st.info(
        f"**Projected Annual Deficit vs Plan**: **{'-' if yearly_rollup['projected_annual_deficit'] < 0 else '+'}${abs(yearly_rollup['projected_annual_deficit']):,.2f}** "
        f"(= (Monthly Deficit ${yearly_rollup['monthly_deficit_vs_plan']:,.2f} × 12) + Periodic Deficits ${yearly_rollup['annual_periodic_deficit']:,.2f})"
    )

    # Section 3: Capital & Non-Operating Expenditures
    if not yearly_rollup["capital_df"].empty:
        st.subheader("3. Capital & Non-Operating Expenditures (Tracked Separately)")
        st.caption("Excluded from operating baseline so major home renovations and capital repairs do not skew your FIRE burn rate.")
        cap_df = yearly_rollup["capital_df"].copy()
        st.dataframe(
            cap_df.style.format({
                "Total Spent": "${:,.2f}",
            }),
            use_container_width=True
        )

    st.markdown("---")

    # Visual Trends
    st.subheader("Core Monthly Spending Over Time (Baseline Budget)")
    st.caption("Shows your 4 core living categories: Discretionary, Groceries, Data, and Utilities.")
    core_cats = ["Discretionary", "Groceries", "Data", "Utilities"]
    core_df = all_df[all_df["category"].isin(core_cats)]
    core_monthly = core_df.groupby(["year_month", "category"])["amount"].sum().abs().reset_index()
    pivot_core = core_monthly.pivot(index="year_month", columns="category", values="amount").fillna(0)
    pivot_core = pivot_core[[c for c in core_cats if c in pivot_core.columns]]
    st.bar_chart(pivot_core)

    st.subheader("Periodic & Capital Spending Over Time (Lump-Sum & Non-Operating)")
    st.caption("Shows Vacations, Repairs, Home Upgrades, Auto, Healthcare, and Off-Budget items.")
    other_df = all_df[~all_df["category"].isin(core_cats)]
    if not other_df.empty:
        other_monthly = other_df.groupby(["year_month", "category"])["amount"].sum().abs().reset_index()
        pivot_other = other_monthly.pivot(index="year_month", columns="category", values="amount").fillna(0)
        st.bar_chart(pivot_other)

    # Multi-Year Category Comparison Table
    st.subheader("Multi-Year Category Monthly Averages Comparison (2022 – 2026)")
    annual_summary = all_df.groupby(["year", "category"])["amount"].sum().abs().reset_index()
    months_per_year = all_df.groupby("year")["month"].nunique().to_dict()
    annual_summary["months_in_year"] = annual_summary["year"].map(months_per_year)
    annual_summary["monthly_avg"] = round(annual_summary["amount"] / annual_summary["months_in_year"], 2)

    pivot_annual = annual_summary.pivot(index="category", columns="year", values="monthly_avg").fillna(0)
    monthly_plan_dict = config.budget_plans.get("monthly", {})
    pivot_annual["Target Plan / Mo"] = pivot_annual.index.map(lambda c: monthly_plan_dict.get(c, "-"))
    st.dataframe(pivot_annual.style.format(precision=2), use_container_width=True)

# -------------------------------------------------------------
# TAB 2: Single-Stream Ledger
# -------------------------------------------------------------
with tab_ledger:
    col_h1, col_h2 = st.columns([4, 1])
    with col_h1:
        st.header("Searchable Unified Ledger")
        st.markdown("All 5,472+ historical transactions in a single continuous stream.")
    with col_h2:
        edit_mode = st.toggle("✏️ Edit Ledger Mode", value=False, help="Enable inline editing of categories, notes, or deleting rows.")

    with st.expander("➕ Log Cash / Manual Transaction", expanded=False):
        render_manual_transaction_form("ledger", repo, config, max_date)

    f_col1, f_col2, f_col3, f_col4 = st.columns([2, 2, 2, 2])
    with f_col1:
        date_range = st.date_input(
            "Date Range",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )
    with f_col2:
        all_categories = sorted(all_df["category"].unique())
        selected_cats = st.multiselect("Filter Categories", all_categories, default=[])
    with f_col3:
        search_query = st.text_input("Search Description / Merchant", "")
    with f_col4:
        sort_order = st.selectbox("Sort By", ["Date (Newest First)", "Date (Oldest First)", "Amount (Highest First)", "Amount (Lowest First)"])

    # Apply filters
    filtered_df = all_df.copy()
    if len(date_range) == 2:
        start_d, end_d = date_range
        filtered_df = filtered_df[
            (filtered_df["date_dt"].dt.date >= start_d) & 
            (filtered_df["date_dt"].dt.date <= end_d)
        ]
    if selected_cats:
        filtered_df = filtered_df[filtered_df["category"].isin(selected_cats)]
    if search_query:
        q = search_query.strip().lower()
        filtered_df = filtered_df[
            filtered_df["description"].str.lower().str.contains(q, na=False) |
            filtered_df["notes"].str.lower().str.contains(q, na=False)
        ]

    # Sorting
    if sort_order == "Date (Newest First)":
        filtered_df = filtered_df.sort_values(by="date", ascending=False)
    elif sort_order == "Date (Oldest First)":
        filtered_df = filtered_df.sort_values(by="date", ascending=True)
    elif sort_order == "Amount (Highest First)":
        filtered_df = filtered_df.sort_values(by="amount", ascending=False)
    elif sort_order == "Amount (Lowest First)":
        filtered_df = filtered_df.sort_values(by="amount", ascending=True)

    # Filtered Metrics
    m1, m2, m3 = st.columns(3)
    m1.metric("Matching Transactions", f"{len(filtered_df):,}")
    m2.metric("Total Spent", f"${abs(filtered_df['amount'].sum()):,.2f}")
    m3.metric("Average Transaction", f"${abs(filtered_df['amount'].mean()):,.2f}" if len(filtered_df) > 0 else "$0.00")

    all_cat_options = sorted(list(config.categories.keys()) + ["Discretionary", "Excluded", "Shared / Reimbursements"])

    if not edit_mode:
        # Table Display
        display_cols = ["date", "description", "amount", "category", "notes", "sheet_source"]
        st.dataframe(
            filtered_df[display_cols].style.format({"amount": "${:,.2f}"}),
            use_container_width=True,
            height=500
        )

        # Export Button
        csv_bytes = filtered_df[display_cols].to_csv(index=False).encode('utf-8')
        st.download_button(
            "📥 Download Filtered Slice as CSV",
            data=csv_bytes,
            file_name="filtered_expenses.csv",
            mime="text/csv"
        )
    else:
        st.info("💡 **Edit Mode Active**: You can change any category, adjust amounts or notes, and delete rows directly. Click **Save Changes** below to commit.")
        editable_cols = ["date", "description", "amount", "category", "notes", "sheet_source"]
        
        edited_ledger = st.data_editor(
            filtered_df[editable_cols],
            column_config={
                "category": st.column_config.SelectboxColumn("Category", options=all_cat_options, required=True),
                "amount": st.column_config.NumberColumn("Amount", format="$%.2f"),
                "notes": st.column_config.TextColumn("Notes"),
                "date": st.column_config.TextColumn("Date", disabled=True),
                "description": st.column_config.TextColumn("Description", disabled=True),
                "sheet_source": st.column_config.TextColumn("Source", disabled=True),
            },
            use_container_width=True,
            num_rows="dynamic",
            height=500,
            key="ledger_editor"
        )

        if st.button("💾 Save Changes to Ledger", type="primary"):
            updated_full = all_df.copy()
            for idx, row in edited_ledger.iterrows():
                if idx in updated_full.index:
                    updated_full.loc[idx, "category"] = row["category"]
                    updated_full.loc[idx, "notes"] = row["notes"]
                    updated_full.loc[idx, "amount"] = row["amount"]
            deleted_indices = set(filtered_df.index) - set(edited_ledger.index)
            if deleted_indices:
                updated_full = updated_full.drop(index=list(deleted_indices))
            
            repo.update_all_transactions(updated_full, backup_csv_path=BACKUP_CSV_PATH)
            st.success("Successfully updated ledger in DuckDB and CSV backup!")
            st.cache_resource.clear()
            st.rerun()

# -------------------------------------------------------------
# TAB 3: Ingest & Bookend Deduplicate Studio
# -------------------------------------------------------------
with tab_ingest:
    st.header("Monthly Bank CSV Ingestion & Deduplication")
    st.markdown("""
    Upload raw monthly transaction exports across any of your bank accounts.
    The engine will normalize column names, classify categories via regex, and 
    **flag monthly bookend duplicates** so you never double-count overlapping dates.
    """)

    with st.expander("💵 Need to log Cash Spending or a Manual Entry for this Month?", expanded=False):
        render_manual_transaction_form("ingest", repo, config, max_date)

    st.markdown("---")

    uploaded_files = st.file_uploader(
        "Upload one or more Bank CSV exports",
        type=["csv"],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.markdown("##### ⚙️ File Settings & Sign Normalization")
        file_invert_settings = {}
        for uf in uploaded_files:
            c1, c2 = st.columns([3, 2])
            with c1:
                st.write(f"📄 `{uf.name}`")
            with c2:
                is_sofi = "sofi" in uf.name.lower()
                file_invert_settings[uf.name] = st.checkbox(
                    "Invert signs (SoFi / Positive CC)",
                    value=is_sofi,
                    key=f"invert_{uf.name}",
                    help="Enable this if this export lists expenses as positive numbers."
                )

        parsed_batches = []
        for uf in uploaded_files:
            try:
                raw_df = pd.read_csv(uf)
                should_invert = file_invert_settings.get(uf.name, None)
                norm = normalizer.normalize_dataframe(raw_df, source_name=uf.name, invert_signs=should_invert)
                parsed_batches.append(norm)
            except Exception as e:
                st.error(f"Error parsing {uf.name}: {e}")

        if parsed_batches:
            combined_incoming = pd.concat(parsed_batches, ignore_index=True)
            st.success(f"Loaded {len(combined_incoming)} transactions across {len(uploaded_files)} file(s).")

            # Run Deduplication Engine
            new_txs, exact_dups, suspects = dedup.deduplicate(combined_incoming, all_df)

            col_a, col_b, col_c = st.columns(3)
            col_a.metric("New Transactions", len(new_txs))
            col_b.metric("Exact Duplicates (Skipped)", len(exact_dups))
            col_c.metric("Bookend Suspects Flagged", len(suspects))

            # Display Bookend Suspects if any
            if suspects:
                st.warning(f"⚠️ {len(suspects)} Bookend Duplicate Candidates Detected!")
                st.markdown("These transactions match existing records within $\pm 3$ days with identical amounts and merchants:")
                
                suspect_rows = []
                for s in suspects:
                    inc = s["incoming"]
                    ex = s["matched_existing"]
                    suspect_rows.append({
                        "Incoming Date": inc["date"],
                        "Incoming Desc": inc["description"],
                        "Amount": f"${inc['amount']:.2f}",
                        "Existing Match Date": ex["date_str"],
                        "Existing Match Desc": ex["description"],
                        "Flag Reason": s["reason"]
                    })
                st.dataframe(pd.DataFrame(suspect_rows), use_container_width=True)

            # Display Exact Duplicates
            if not exact_dups.empty:
                with st.expander(f"View {len(exact_dups)} Exact Duplicates (Will Be Skipped)"):
                    st.dataframe(exact_dups[["date", "description", "amount", "category"]], use_container_width=True)

            # Display New Transactions with Staging Data Editor
            if not new_txs.empty:
                st.subheader(f"Interactive Review: {len(new_txs)} Transactions Ready to Stage")
                st.markdown("""
                * **`Include?`**: Uncheck any row you want to skip or exclude (e.g. mortgages, tenant rent, transfers).
                * **`Category`**: Click the category dropdown on any row to reassign it (e.g. flip to *Home Upgrades* or *Repairs*).
                * **`Notes`**: Add details for reimbursements or tax tracking.
                """)

                all_cat_options = sorted(list(config.categories.keys()) + ["Discretionary", "Excluded", "Shared / Reimbursements"])
                if "include" not in new_txs.columns:
                    new_txs["include"] = True

                staging_cols = ["include", "date", "description", "amount", "category", "notes", "sheet_source"]
                
                edited_staging = st.data_editor(
                    new_txs[staging_cols],
                    column_config={
                        "include": st.column_config.CheckboxColumn("Include?", default=True, help="Uncheck to exclude"),
                        "category": st.column_config.SelectboxColumn("Category", options=all_cat_options, required=True),
                        "amount": st.column_config.NumberColumn("Amount", format="$%.2f", disabled=True),
                        "date": st.column_config.TextColumn("Date", disabled=True),
                        "description": st.column_config.TextColumn("Description", disabled=True),
                        "notes": st.column_config.TextColumn("Notes"),
                        "sheet_source": st.column_config.TextColumn("Source", disabled=True)
                    },
                    use_container_width=True,
                    num_rows="dynamic",
                    key="staging_editor"
                )

                to_commit = edited_staging[edited_staging["include"] == True].copy()
                to_commit = to_commit[to_commit["category"] != "Excluded"]
                excluded_count = len(edited_staging) - len(to_commit)

                st.markdown(f"**Ready to commit**: **{len(to_commit)}** approved transactions (*{excluded_count} excluded/skipped*).")

                if st.button("🚀 Commit Approved Transactions to Ledger", type="primary"):
                    save_batch = to_commit.drop(columns=["include"])
                    repo.append_transactions(save_batch, backup_csv_path=BACKUP_CSV_PATH)
                    st.success(f"Successfully committed {len(save_batch)} transactions to spending.duckdb and CSV backup!")
                    st.cache_resource.clear()
                    st.rerun()
            else:
                st.info("No new transactions found to commit (all were duplicates).")
