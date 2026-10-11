"""
Unit tests for the personal finance engine.
Tests column normalization, regex categorization, exact deduplication,
bookend drift deduplication, and DuckDB ledger repository operations.
"""

import os
import tempfile
import pandas as pd
import pytest
from engine import (
    RuleConfig,
    TransactionNormalizer,
    DeduplicationEngine,
    LedgerRepository,
    clean_merchant_string,
)

SAMPLE_CONFIG_YAML = """
column_aliases:
  date:
    - "Date"
    - "Posted Date"
    - "Transaction Date"
  description:
    - "Description"
    - "Payee"
    - "transaction"
  amount:
    - "Amount"
    - "amount"
  category:
    - "Category"
    - "Notes"

default_category: "Discretionary"

categories:
  Groceries:
    - "shoprite.*"
    - "whole foods.*"
    - "99 ranch market.*"
  Utilities:
    - ".*pseg.*"
    - ".*ridgewood water.*"
  Data:
    - "verizon.*"
    - "spotify.*"
  Vacations:
    - "airbnb.*"
    - "disney.*"

exclude_patterns:
  - ".*mortgage.*"
  - ".*tenant rent.*"
  - ".*autopay.*"

budget_plans:
  monthly:
    Groceries: 750.00
    Discretionary: 4000.00
    Data: 200.00
    Utilities: 250.00
"""


@pytest.fixture
def temp_config_file():
    with tempfile.NamedTemporaryFile("w", suffix=".yaml", delete=False) as f:
        f.write(SAMPLE_CONFIG_YAML)
        path = f.name
    yield path
    if os.path.exists(path):
        os.remove(path)


@pytest.fixture
def rule_config(temp_config_file):
    return RuleConfig(temp_config_file)


class TestCategorization:
    def test_regex_matching(self, rule_config):
        assert rule_config.categorize("Shoprite Fair Lawn") == "Groceries"
        assert rule_config.categorize("WHOLE FOODS MARKET #10") == "Groceries"
        assert rule_config.categorize("Public Service PSEG Electric") == "Utilities"
        assert rule_config.categorize("VERIZON WIRELESS BILL") == "Data"
        assert rule_config.categorize("SPOTIFY USA") == "Data"
        assert rule_config.categorize("Airbnb London stay") == "Vacations"

    def test_fallback_default_category(self, rule_config):
        assert rule_config.categorize("Random Unknown Coffee Shop") == "Discretionary"
        assert rule_config.categorize("") == "Discretionary"
        assert rule_config.categorize(None) == "Discretionary"

    def test_exclusion_patterns(self, rule_config):
        assert rule_config.is_excluded("WELLS FARGO MORTGAGE PAYMENT") is True
        assert rule_config.is_excluded("Deposit: Tenant Rent Unit 2") is True
        assert rule_config.is_excluded("CHASE CREDIT CARD AUTOPAY") is True
        assert rule_config.is_excluded("Shoprite Groceries") is False
        assert rule_config.is_excluded("") is False


class TestNormalizer:
    def test_column_alias_resolution(self, rule_config):
        normalizer = TransactionNormalizer(rule_config)
        cols1 = ["Posted Date", "Payee", "Amount"]
        assert normalizer.resolve_column(cols1, "date") == "Posted Date"
        assert normalizer.resolve_column(cols1, "description") == "Payee"
        assert normalizer.resolve_column(cols1, "amount") == "Amount"

        cols2 = ["Date", "transaction", "amount"]
        assert normalizer.resolve_column(cols2, "date") == "Date"
        assert normalizer.resolve_column(cols2, "description") == "transaction"
        assert normalizer.resolve_column(cols2, "amount") == "amount"

    def test_date_parsing(self, rule_config):
        normalizer = TransactionNormalizer(rule_config)
        assert normalizer.parse_date("2024-03-15") == "2024-03-15"
        assert normalizer.parse_date("03/15/2024") == "2024-03-15"
        # Excel serial dates: 45352 = 2024-03-01, 45353 = 2024-03-02
        assert normalizer.parse_date(45352) == "2024-03-01"
        assert normalizer.parse_date(45353) == "2024-03-02"

    def test_amount_parsing(self, rule_config):
        normalizer = TransactionNormalizer(rule_config)
        assert normalizer.parse_amount("$12.50") == 12.50
        assert normalizer.parse_amount("-$1,250.75") == -1250.75
        assert normalizer.parse_amount("(99.95)") == -99.95
        assert normalizer.parse_amount(45.2) == 45.2

    def test_dataframe_normalization_and_sign_inversion(self, rule_config):
        normalizer = TransactionNormalizer(rule_config)
        # Simulate credit card export where all charges are positive
        raw_df = pd.DataFrame({
            "Posted Date": ["2024-03-01", "2024-03-02"],
            "Payee": ["Shoprite Fair Lawn", "Random Store"],
            "Amount": ["45.50", "12.00"]
        })
        norm = normalizer.normalize_dataframe(raw_df, source_name="test.csv")
        assert len(norm) == 2
        # Should be inverted to negative amounts
        assert norm.iloc[0]["amount"] == -45.50
        assert norm.iloc[0]["category"] == "Groceries"
        assert norm.iloc[1]["amount"] == -12.00
        assert norm.iloc[1]["category"] == "Discretionary"

    def test_explicit_and_sofi_sign_inversion(self, rule_config):
        normalizer = TransactionNormalizer(rule_config)
        raw_df = pd.DataFrame({
            "Date": ["2024-04-01"],
            "Description": ["Target"],
            "Amount": ["100.00"]
        })
        # 1. SoFi auto-inversion by filename
        sofi_norm = normalizer.normalize_dataframe(raw_df, source_name="sofi_export_april.csv")
        assert sofi_norm.iloc[0]["amount"] == -100.00

        # 2. Explicit invert_signs=False keeps original positive
        no_invert = normalizer.normalize_dataframe(raw_df, source_name="custom.csv", invert_signs=False)
        assert no_invert.iloc[0]["amount"] == 100.00

        # 3. Explicit invert_signs=True inverts
        forced_invert = normalizer.normalize_dataframe(raw_df, source_name="custom.csv", invert_signs=True)
        assert forced_invert.iloc[0]["amount"] == -100.00

    def test_exclusion_normalization(self, rule_config):
        normalizer = TransactionNormalizer(rule_config)
        raw_df = pd.DataFrame({
            "Date": ["2024-04-01", "2024-04-02"],
            "Description": ["Wells Fargo Mortgage Payment", "Shoprite"],
            "Amount": ["-2500.00", "-50.00"]
        })
        norm = normalizer.normalize_dataframe(raw_df, source_name="test.csv")
        assert len(norm) == 2
        # First row is mortgage => auto-excluded
        assert bool(norm.iloc[0]["include"]) is False
        assert norm.iloc[0]["category"] == "Excluded"
        assert "Auto-excluded" in norm.iloc[0]["notes"]
        # Second row is normal groceries => included
        assert bool(norm.iloc[1]["include"]) is True
        assert norm.iloc[1]["category"] == "Groceries"


class TestDeduplicationEngine:
    def test_exact_duplicate_detection(self):
        engine = DeduplicationEngine(bookend_day_window=3)
        existing = pd.DataFrame([{
            "date": "2024-03-15",
            "amount": -50.00,
            "description": "Shoprite Fair Lawn",
            "category": "Groceries"
        }])

        incoming = pd.DataFrame([
            {"date": "2024-03-15", "amount": -50.00, "description": "Shoprite Fair Lawn", "category": "Groceries"},
            {"date": "2024-03-16", "amount": -25.00, "description": "Target", "category": "Discretionary"}
        ])

        new_txs, exact_dups, suspects = engine.deduplicate(incoming, existing)
        assert len(exact_dups) == 1
        assert exact_dups.iloc[0]["amount"] == -50.00
        assert len(new_txs) == 1
        assert new_txs.iloc[0]["description"] == "Target"
        assert len(suspects) == 0

    def test_bookend_monthly_drift_deduplication(self):
        """
        Tests transaction occurring on March 31 but posting on April 2
        across monthly statement boundaries.
        """
        engine = DeduplicationEngine(bookend_day_window=3)
        existing = pd.DataFrame([{
            "date": "2024-03-31",
            "amount": -89.45,
            "description": "WHOLE FOODS MARKET #10106",
            "category": "Groceries"
        }])

        # Incoming export contains the same charge posted on April 2
        incoming = pd.DataFrame([{
            "date": "2024-04-02",
            "amount": -89.45,
            "description": "WHOLE FOODS MARKET",
            "category": "Groceries"
        }])

        new_txs, exact_dups, suspects = engine.deduplicate(incoming, existing)
        assert len(exact_dups) == 0
        assert len(new_txs) == 0
        assert len(suspects) == 1
        assert "Matches existing transaction within 2 days" in suspects[0]["reason"]
        assert suspects[0]["incoming"]["date"] == "2024-04-02"
        assert suspects[0]["matched_existing"]["date_str"] == "2024-03-31"


class TestLedgerRepository:
    def test_repository_lifecycle(self, rule_config):
        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = os.path.join(tmpdir, "test.duckdb")
            repo = LedgerRepository(db_path)

            df = pd.DataFrame([
                {
                    "date": "2024-01-10",
                    "year_month": "2024-01",
                    "year": 2024,
                    "month": 1,
                    "description": "Shoprite",
                    "amount": -500.00,
                    "category": "Groceries",
                    "sheet_source": "test",
                    "row_number": 1,
                    "notes": ""
                },
                {
                    "date": "2024-01-15",
                    "year_month": "2024-01",
                    "year": 2024,
                    "month": 1,
                    "description": "Target Shopping",
                    "amount": -3000.00,
                    "category": "Discretionary",
                    "sheet_source": "test",
                    "row_number": 2,
                    "notes": ""
                }
            ])

            repo.append_transactions(df)
            loaded = repo.get_all_transactions()
            assert len(loaded) == 2

            summary = repo.get_fire_summary(rule_config.budget_plans)
            monthly_rep = summary["monthly_report"]
            assert len(monthly_rep) == 1
            # Groceries spent -500 with plan 750 => deficit = -1 * (-500 + 750) = -250 (i.e. +250 under budget)
            assert monthly_rep.iloc[0]["Groceries_spend"] == -500.00
            assert monthly_rep.iloc[0]["Discretionary_spend"] == -3000.00

            # Test update_all_transactions
            modified = loaded.copy()
            modified.loc[modified["description"] == "Shoprite", "category"] = "Home Upgrades"
            repo.update_all_transactions(modified)
            reloaded = repo.get_all_transactions()
            assert reloaded.loc[reloaded["description"] == "Shoprite", "category"].iloc[0] == "Home Upgrades"
