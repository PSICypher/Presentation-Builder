"""Tests for the data ingestion module."""

import math
import os
import tempfile
from pathlib import Path

import pandas as pd
import pytest

from src.processor.ingestion import (
    clean_columns,
    clean_numeric_columns,
    clean_percentage_columns,
    detect_encoding,
    ingest,
    ingest_offer_performance,
    ingest_product_sales,
    ingest_targets,
    ingest_historical,
    ingest_tracker,
    ingest_crm,
    ingest_affiliate,
    parse_numeric,
    parse_percentage,
    read_csv_auto,
    SOURCE_TYPES,
)


# ---------------------------------------------------------------------------
# parse_numeric
# ---------------------------------------------------------------------------

class TestParseNumeric:
    def test_plain_integer(self):
        assert parse_numeric(42) == 42.0

    def test_plain_float(self):
        assert parse_numeric(3.14) == 3.14

    def test_comma_formatted(self):
        assert parse_numeric("63,571") == 63571.0

    def test_large_comma_formatted(self):
        assert parse_numeric("1,138,771") == 1138771.0

    def test_float_artifact(self):
        assert parse_numeric("49,156.000000000") == 49156.0

    def test_nan_value(self):
        assert math.isnan(parse_numeric(float("nan")))

    def test_none_value(self):
        assert math.isnan(parse_numeric(None))

    def test_empty_string(self):
        assert math.isnan(parse_numeric(""))

    def test_whitespace_string(self):
        assert math.isnan(parse_numeric("   "))

    def test_non_numeric_string(self):
        assert math.isnan(parse_numeric("hello"))

    def test_negative(self):
        assert parse_numeric("-1,234") == -1234.0

    def test_decimal(self):
        assert parse_numeric("23.53") == 23.53


# ---------------------------------------------------------------------------
# parse_percentage
# ---------------------------------------------------------------------------

class TestParsePercentage:
    def test_simple_percentage(self):
        assert parse_percentage("170.6%") == pytest.approx(1.706)

    def test_negative_percentage(self):
        assert parse_percentage("-21.6%") == pytest.approx(-0.216)

    def test_positive_sign(self):
        assert parse_percentage("+5.2%") == pytest.approx(0.052)

    def test_ppts(self):
        assert parse_percentage("-3.52 ppts") == pytest.approx(-0.0352)

    def test_ppts_positive(self):
        assert parse_percentage("+1.5 ppts") == pytest.approx(0.015)

    def test_with_arrow_down(self):
        assert parse_percentage("-21.6% ↓") == pytest.approx(-0.216)

    def test_with_arrow_up(self):
        assert parse_percentage("+5.0% ↑") == pytest.approx(0.05)

    def test_absolute_rate(self):
        assert parse_percentage("96.90%") == pytest.approx(0.969)

    def test_nan(self):
        assert math.isnan(parse_percentage(float("nan")))

    def test_none(self):
        assert math.isnan(parse_percentage(None))

    def test_numeric_passthrough(self):
        assert parse_percentage(0.5) == 0.5

    def test_empty_string(self):
        assert math.isnan(parse_percentage(""))

    def test_unparseable(self):
        assert math.isnan(parse_percentage("N/A"))


# ---------------------------------------------------------------------------
# clean_columns
# ---------------------------------------------------------------------------

class TestCleanColumns:
    def test_strips_whitespace(self):
        df = pd.DataFrame({"  foo  ": [1], "bar ": [2], " baz": [3]})
        df = clean_columns(df)
        assert list(df.columns) == ["foo", "bar", "baz"]

    def test_no_change_needed(self):
        df = pd.DataFrame({"a": [1], "b": [2]})
        df = clean_columns(df)
        assert list(df.columns) == ["a", "b"]


# ---------------------------------------------------------------------------
# detect_encoding
# ---------------------------------------------------------------------------

class TestDetectEncoding:
    def test_utf16_le_bom(self, tmp_path):
        p = tmp_path / "test.csv"
        content = "col1\tcol2\n1\t2\n"
        p.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
        enc, sep = detect_encoding(p)
        assert enc == "utf-16-le"
        assert sep == "\t"

    def test_utf8(self, tmp_path):
        p = tmp_path / "test.csv"
        p.write_text("col1,col2\n1,2\n", encoding="utf-8")
        enc, sep = detect_encoding(p)
        assert enc == "utf-8"
        assert sep == ","


# ---------------------------------------------------------------------------
# read_csv_auto
# ---------------------------------------------------------------------------

class TestReadCsvAuto:
    def test_utf8_csv(self, tmp_path):
        p = tmp_path / "test.csv"
        p.write_text("col1 ,col2\n1,2\n3,4\n", encoding="utf-8")
        df = read_csv_auto(p)
        assert list(df.columns) == ["col1", "col2"]
        assert len(df) == 2

    def test_utf16_csv(self, tmp_path):
        p = tmp_path / "test.csv"
        content = "col1\tcol2\n10\t20\n"
        p.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
        df = read_csv_auto(p)
        assert "col1" in df.columns
        assert len(df) == 1


# ---------------------------------------------------------------------------
# clean_numeric_columns / clean_percentage_columns
# ---------------------------------------------------------------------------

class TestColumnCleaners:
    def test_clean_numeric(self):
        df = pd.DataFrame({"val": ["1,234", "5,678", "910"]})
        clean_numeric_columns(df, ["val"])
        assert df["val"].tolist() == [1234.0, 5678.0, 910.0]

    def test_clean_numeric_auto_detect(self):
        df = pd.DataFrame({"num": ["100", "200"], "txt": ["a", "b"]})
        clean_numeric_columns(df)
        # "a" and "b" become NaN, "100"/"200" become floats
        assert df["num"].tolist() == [100.0, 200.0]

    def test_clean_percentage(self):
        df = pd.DataFrame({"pct": ["5.0%", "-3.2 ppts", "+10.0% ↑"]})
        clean_percentage_columns(df, ["pct"])
        assert df["pct"].iloc[0] == pytest.approx(0.05)
        assert df["pct"].iloc[1] == pytest.approx(-0.032)
        assert df["pct"].iloc[2] == pytest.approx(0.10)

    def test_missing_column_ignored(self):
        df = pd.DataFrame({"a": [1]})
        clean_numeric_columns(df, ["nonexistent"])
        assert list(df.columns) == ["a"]


# ---------------------------------------------------------------------------
# ingest_targets (well-structured CSV, good for integration test)
# ---------------------------------------------------------------------------

class TestIngestTargets:
    def test_basic(self, tmp_path):
        p = tmp_path / "targets.csv"
        p.write_text(
            "Target_Type_Id,Date,Site_Id,Locale_Id,Channel_Id,Notes,"
            "Gross_Revenue_Target,Net_Revenue_Target,Marketing_Spend_Target,"
            "Session_Target,Order_Target,New_Customer_Target\n"
            "Daily,2025-09-01,No 7,en_US,AFFILIATE,,519.59,552.76,215.03,121,9,13\n",
            encoding="utf-8",
        )
        df = ingest_targets(p)
        assert len(df) == 1
        assert df["Channel_Id"].iloc[0] == "AFFILIATE"
        assert pd.notna(df["Date"].iloc[0])
        assert df["Gross_Revenue_Target"].iloc[0] == pytest.approx(519.59)


# ---------------------------------------------------------------------------
# ingest_offer_performance
# ---------------------------------------------------------------------------

class TestIngestOfferPerformance:
    def test_basic(self, tmp_path):
        p = tmp_path / "offers.csv"
        content = (
            "Dimension 1\tDimension 2\tRedemptions \t% Change Redemptions\t"
            "Revenue \t% Change Revenue\n"
            "Promo A\tAFFILIATE\t26,338\t-21.6% ↓\t2,173,984\t-21.9%\n"
        )
        p.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
        df = ingest_offer_performance(p)
        assert df["Redemptions"].iloc[0] == 26338.0
        assert df["% Change Redemptions"].iloc[0] == pytest.approx(-0.216)
        assert df["Revenue"].iloc[0] == 2173984.0
        assert df["% Change Revenue"].iloc[0] == pytest.approx(-0.219)


# ---------------------------------------------------------------------------
# ingest_product_sales
# ---------------------------------------------------------------------------

class TestIngestProductSales:
    def test_basic(self, tmp_path):
        p = tmp_path / "products.csv"
        content = (
            "Dimension 1 \tUnits (Analysis)\tUnits (Comparison)\tUnits (vs. Comp)\n"
            "Product A\t63,571\t23,500\t170.6%\n"
        )
        p.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
        df = ingest_product_sales(p)
        assert "Dimension 1" in df.columns  # trailing space stripped
        assert df["Units (Analysis)"].iloc[0] == 63571.0
        assert df["Units (vs. Comp)"].iloc[0] == pytest.approx(1.706)


# ---------------------------------------------------------------------------
# ingest_historical
# ---------------------------------------------------------------------------

class TestIngestHistorical:
    def test_basic(self, tmp_path):
        p = tmp_path / "hist.csv"
        content = (
            "Dimension 1\tDimension 2\tOrders\tNew Customers\tCoS\t"
            "Cost\tRevenue\tPhased\n"
            "2024\t10\t4,659\t1,725\t96.90%\t246,410\t254,388\t3%\n"
        )
        p.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
        df = ingest_historical(p)
        assert df["Orders"].iloc[0] == 4659.0
        assert df["Revenue"].iloc[0] == 254388.0
        assert df["CoS"].iloc[0] == pytest.approx(0.969)
        assert df["Phased"].iloc[0] == pytest.approx(0.03)


# ---------------------------------------------------------------------------
# ingest() dispatcher
# ---------------------------------------------------------------------------

class TestIngestDispatcher:
    def test_unknown_type(self):
        with pytest.raises(ValueError, match="Unknown source type"):
            ingest("/fake/path", "nonexistent_type")

    def test_all_types_registered(self):
        expected = {
            "tracker", "product_sales", "offer_performance",
            "crm", "affiliate", "targets", "trading", "historical",
        }
        assert set(SOURCE_TYPES.keys()) == expected

    def test_dispatch_targets(self, tmp_path):
        p = tmp_path / "targets.csv"
        p.write_text(
            "Target_Type_Id,Date,Channel_Id\nDaily,2025-09-01,PPC\n",
            encoding="utf-8",
        )
        df = ingest(p, "targets")
        assert len(df) == 1


# ---------------------------------------------------------------------------
# CRM typo handling
# ---------------------------------------------------------------------------

class TestCrmTypoHandling:
    def test_unsibscribe_renamed(self, tmp_path):
        p = tmp_path / "crm.xlsx"
        df = pd.DataFrame({
            "Emails Sent": [1000],
            "Unsibscribe Rate": [0.02],
        })
        df.to_excel(p, index=False, engine="openpyxl")
        result = ingest_crm(p, sheet_name="Sheet1")
        assert "Unsubscribe Rate" in result.columns
        assert "Unsibscribe Rate" not in result.columns


# ---------------------------------------------------------------------------
# Affiliate typo handling
# ---------------------------------------------------------------------------

class TestAffiliateTypoHandling:
    def test_sale_actvie_renamed(self, tmp_path):
        p = tmp_path / "affiliate.xlsx"
        df = pd.DataFrame({
            "Dimension 1": [1004849],
            "Sale-Actvie Publishers (CountD) (Analysis)": [5],
        })
        df.to_excel(p, index=False, engine="openpyxl")
        result = ingest_affiliate(p, sheet_name="Sheet1")
        assert "Sale-Active Publishers (CountD) (Analysis)" in result.columns
