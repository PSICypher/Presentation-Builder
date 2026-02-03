"""Tests for the data transformation module."""

import math

import pandas as pd
import pytest

from src.processor.transform import (
    CHANNELS,
    DataTransformer,
    ReportContext,
    _decimal_to_pct,
    _nan_to_none,
    _safe_div,
    _safe_pct_change,
    _safe_ppt_change,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def ctx():
    return ReportContext(year=2026, month=1)


@pytest.fixture
def transformer(ctx):
    return DataTransformer(ctx)


def _make_raw_data(rows):
    """Build a RAW DATA DataFrame from simplified row dicts."""
    columns = [
        "COS Year", "COS Month", "COS Day", "COS Channel",
        "COS Revenue", "COS Orders", "COS Sessions",
        "COS Cost", "COS New Customers",
    ]
    data = []
    for r in rows:
        data.append([
            r.get("year", 2026), r.get("month", 1), r.get("day", 1),
            r.get("channel", "DIRECT"),
            r.get("revenue", 100.0), r.get("orders", 10),
            r.get("sessions", 200), r.get("cost", 20.0),
            r.get("new_customers", 5),
        ])
    return pd.DataFrame(data, columns=columns)


def _make_targets(rows):
    """Build a targets DataFrame from simplified row dicts."""
    data = []
    for r in rows:
        data.append({
            "Date": pd.Timestamp(r.get("date", "2026-01-01")),
            "Channel_Id": r.get("channel", "DIRECT"),
            "Net_Revenue_Target": r.get("revenue", 100.0),
            "Marketing_Spend_Target": r.get("spend", 10.0),
            "Session_Target": r.get("sessions", 100),
            "Order_Target": r.get("orders", 10),
            "New_Customer_Target": r.get("new_customers", 5),
        })
    return pd.DataFrame(data)


# ---------------------------------------------------------------------------
# ReportContext
# ---------------------------------------------------------------------------

class TestReportContext:
    def test_month_name(self):
        assert ReportContext(2026, 1).month_name == "January"
        assert ReportContext(2026, 12).month_name == "December"

    def test_days_in_month(self):
        assert ReportContext(2026, 1).days_in_month == 31
        assert ReportContext(2026, 2).days_in_month == 28
        assert ReportContext(2024, 2).days_in_month == 29  # leap year

    def test_prior_year(self):
        assert ReportContext(2026, 1).prior_year == 2025


# ---------------------------------------------------------------------------
# Safe math helpers
# ---------------------------------------------------------------------------

class TestSafeDiv:
    def test_normal(self):
        assert _safe_div(10, 5) == 2.0

    def test_zero_denominator(self):
        assert math.isnan(_safe_div(10, 0))

    def test_none_denominator(self):
        assert math.isnan(_safe_div(10, None))

    def test_nan_denominator(self):
        assert math.isnan(_safe_div(10, float("nan")))

    def test_nan_numerator(self):
        assert math.isnan(_safe_div(float("nan"), 5))

    def test_custom_default(self):
        assert _safe_div(10, 0, default=0.0) == 0.0


class TestSafePctChange:
    def test_positive_change(self):
        assert _safe_pct_change(110, 100) == pytest.approx(10.0)

    def test_negative_change(self):
        assert _safe_pct_change(90, 100) == pytest.approx(-10.0)

    def test_zero_prior(self):
        assert math.isnan(_safe_pct_change(100, 0))

    def test_nan_prior(self):
        assert math.isnan(_safe_pct_change(100, float("nan")))

    def test_none_prior(self):
        assert math.isnan(_safe_pct_change(100, None))


class TestSafePptChange:
    def test_positive(self):
        # 0.10 - 0.08 = 0.02 -> 2.0 ppts
        assert _safe_ppt_change(0.10, 0.08) == pytest.approx(2.0)

    def test_negative(self):
        assert _safe_ppt_change(0.08, 0.10) == pytest.approx(-2.0)

    def test_nan(self):
        assert math.isnan(_safe_ppt_change(float("nan"), 0.10))
        assert math.isnan(_safe_ppt_change(0.10, float("nan")))

    def test_none(self):
        assert math.isnan(_safe_ppt_change(None, 0.10))


class TestDecimalToPct:
    def test_normal(self):
        assert _decimal_to_pct(0.095) == pytest.approx(9.5)

    def test_zero(self):
        assert _decimal_to_pct(0.0) == 0.0

    def test_nan(self):
        assert math.isnan(_decimal_to_pct(float("nan")))

    def test_none(self):
        assert math.isnan(_decimal_to_pct(None))


class TestNanToNone:
    def test_nan(self):
        assert _nan_to_none(float("nan")) is None

    def test_number(self):
        assert _nan_to_none(42.0) == 42.0

    def test_string(self):
        assert _nan_to_none("hello") == "hello"


# ---------------------------------------------------------------------------
# Cover KPIs
# ---------------------------------------------------------------------------

class TestTransformCover:
    def test_with_data(self, transformer):
        raw = _make_raw_data([
            {"channel": "DIRECT", "revenue": 1000, "orders": 50, "sessions": 500,
             "cost": 100, "new_customers": 20},
            {"channel": "PPC", "revenue": 500, "orders": 25, "sessions": 300,
             "cost": 80, "new_customers": 10},
        ])
        sources = {"tracker": {"RAW DATA": raw}}
        result = transformer._transform_cover(sources)

        assert result["cover.report_title"] == "No7 US x THGi Monthly eComm Report"
        assert result["cover.report_period"] == "January 2026 Overview"
        assert result["cover.total_revenue"] == 1500
        assert result["cover.total_orders"] == 75
        assert result["cover.aov"] == pytest.approx(20.0)
        assert result["cover.new_customers"] == 30
        assert result["cover.cvr"] == pytest.approx(9.375)  # 75/800 * 100
        assert result["cover.cos"] == pytest.approx(12.0)    # 180/1500 * 100

    def test_with_targets(self, transformer):
        raw = _make_raw_data([
            {"channel": "DIRECT", "revenue": 1100, "orders": 55, "sessions": 500,
             "cost": 100, "new_customers": 22},
        ])
        targets = _make_targets([
            {"channel": "DIRECT", "revenue": 1000, "orders": 50, "sessions": 500,
             "spend": 90, "new_customers": 20},
        ])
        sources = {"tracker": {"RAW DATA": raw}, "targets": targets}
        result = transformer._transform_cover(sources)

        assert result["cover.revenue_vs_target"] == pytest.approx(10.0)
        assert result["cover.orders_vs_target"] == pytest.approx(10.0)
        assert result["cover.nc_vs_target"] == pytest.approx(10.0)

    def test_no_data(self, transformer):
        result = transformer._transform_cover({})
        assert result["cover.total_revenue"] is None
        assert result["cover.revenue_vs_target"] is None

    def test_empty_raw(self, transformer):
        raw = _make_raw_data([])
        sources = {"tracker": {"RAW DATA": raw}}
        result = transformer._transform_cover(sources)
        assert result["cover.total_revenue"] is None


# ---------------------------------------------------------------------------
# Executive Summary
# ---------------------------------------------------------------------------

class TestTransformExecutiveSummary:
    def test_builds_channel_rows(self, transformer):
        raw = _make_raw_data([
            {"year": 2026, "month": 1, "channel": "DIRECT", "revenue": 1000, "orders": 50,
             "sessions": 500, "cost": 100, "new_customers": 20},
            {"year": 2026, "month": 1, "channel": "PPC", "revenue": 500, "orders": 25,
             "sessions": 300, "cost": 80, "new_customers": 10},
            {"year": 2025, "month": 1, "channel": "DIRECT", "revenue": 900, "orders": 45,
             "sessions": 450, "cost": 90, "new_customers": 18},
        ])
        sources = {"tracker": {"RAW DATA": raw}}
        result = transformer._transform_executive_summary(sources)

        rows = result["exec.performance_rows"]
        assert len(rows) == 10  # Total + 9 channels

        total_row = rows[0]
        assert total_row["channel"] == "Total"
        assert total_row["revenue"] == 1500
        assert total_row["orders"] == 75

        direct_row = next(r for r in rows if r["channel"] == "DIRECT")
        assert direct_row["revenue"] == 1000
        # YoY: (1000-900)/900 * 100 = 11.11%
        assert direct_row["revenue_vs_ly"] == pytest.approx(11.111, rel=0.01)

    def test_no_data(self, transformer):
        result = transformer._transform_executive_summary({})
        assert result["exec.performance_rows"] == []


# ---------------------------------------------------------------------------
# Daily Performance
# ---------------------------------------------------------------------------

class TestTransformDaily:
    def test_daily_revenue(self, transformer):
        raw = _make_raw_data([
            {"day": 1, "channel": "DIRECT", "revenue": 100},
            {"day": 1, "channel": "PPC", "revenue": 50},
            {"day": 2, "channel": "DIRECT", "revenue": 200},
        ])
        targets = _make_targets([
            {"date": "2026-01-01", "channel": "DIRECT", "revenue": 120},
            {"date": "2026-01-02", "channel": "DIRECT", "revenue": 180},
        ])
        sources = {"tracker": {"RAW DATA": raw}, "targets": targets}
        result = transformer._transform_daily(sources)

        assert result["daily.dates"][0] == "1/1"
        assert result["daily.dates"][1] == "1/2"
        assert len(result["daily.dates"]) == 31
        assert result["daily.revenue_actual"][0] == 150  # 100+50
        assert result["daily.revenue_actual"][1] == 200

    def test_achievement_gauge(self, transformer):
        raw = _make_raw_data([
            {"day": 1, "revenue": 500},
        ])
        targets = _make_targets([
            {"date": "2026-01-01", "revenue": 1000},
        ])
        sources = {"tracker": {"RAW DATA": raw}, "targets": targets}
        result = transformer._transform_daily(sources)

        assert result["daily.revenue_achieved_pct"] == pytest.approx(50.0)
        assert result["daily.revenue_remaining_pct"] == pytest.approx(50.0)

    def test_no_data(self, transformer):
        result = transformer._transform_daily({})
        assert result["daily.dates"] == []


# ---------------------------------------------------------------------------
# Promotion Performance
# ---------------------------------------------------------------------------

class TestTransformPromotions:
    def test_top_promotions(self, transformer):
        data = pd.DataFrame({
            "Dimension 1": ["Promo A", "Promo B", "Grand Total"],
            "Dimension 2": ["Total", "Total", "Total"],
            "Dimension 3": ["1", "1", "1"],
            "Dimension 4": ["Total", "Total", "Total"],
            "Redemptions": [100.0, 200.0, 300.0],
            "Revenue": [5000.0, 3000.0, 8000.0],
            "Discount Amount": [500.0, 300.0, 800.0],
            "% Change Revenue": [-0.10, 0.20, 0.05],
            "% Change Redemptions": [-0.05, 0.15, 0.05],
        })
        sources = {"offer_performance": data}
        result = transformer._transform_promotions(sources)

        rows = result["promo.rows"]
        assert len(rows) == 2  # Grand Total excluded
        assert rows[0]["promotion_name"] == "Promo A"  # higher revenue
        assert rows[0]["revenue"] == 5000.0
        assert rows[0]["revenue_vs_ly"] == pytest.approx(-10.0)

    def test_no_data(self, transformer):
        result = transformer._transform_promotions({})
        assert result["promo.rows"] == []


# ---------------------------------------------------------------------------
# Product Performance
# ---------------------------------------------------------------------------

class TestTransformProducts:
    def test_top_products(self, transformer):
        data = pd.DataFrame({
            "Dimension 1": ["Product A", "Product B", "Grand Total"],
            "Dimension 2": ["1", "1", "1"],
            "Dimension 3": ["Total", "Total", "Total"],
            "Units (Analysis)": [100.0, 50.0, 150.0],
            "Units (vs. Comp)": [0.50, -0.10, 0.30],
            "Product Revenue (Analysis)": [10000.0, 5000.0, 15000.0],
            "Product Revenue (vs. Comp)": [0.20, -0.05, 0.10],
            "AOV (Analysis)": [100.0, 100.0, 100.0],
            "Avg. Selling Price (Analysis)": [50.0, 45.0, 48.0],
            "Total Discount % (Analysis)": [0.15, 0.20, 0.17],
            "New Customers (Analysis)": [30.0, 15.0, 45.0],
        })
        sources = {"product_sales": data}
        result = transformer._transform_products(sources)

        rows = result["product.rows"]
        assert len(rows) == 2  # Grand Total excluded
        assert rows[0]["product_name"] == "Product A"
        assert rows[0]["revenue"] == 10000.0
        assert rows[0]["units_vs_ly"] == pytest.approx(50.0)

    def test_no_data(self, transformer):
        result = transformer._transform_products({})
        assert result["product.rows"] == []


# ---------------------------------------------------------------------------
# CRM Performance
# ---------------------------------------------------------------------------

class TestTransformCRM:
    def test_from_crm_excel(self, transformer):
        data = pd.DataFrame({
            "col_a": ["Grand Total", "Grand Total", "Grand Total"],
            "col_b": ["Total", "Total", "Total"],
            "Campaign Type": ["Manual", "Automated", "Total"],
            "Emails Sent": [50000, 20000, 70000],
            "Open Rate": [0.35, 0.45, 0.38],
            "Click-Through Rate": [0.05, 0.08, 0.06],
            "Revenue": [100000, 50000, 150000],
            "CVR": [0.03, 0.05, 0.035],
            "AOV": [45.0, 50.0, 47.0],
            "Sessions": [5000, 2000, 7000],
            "Orders": [150, 100, 250],
        })
        sources = {"crm": data}
        result = transformer._transform_crm(sources)

        assert result["crm.emails_sent"] == 70000
        assert result["crm.open_rate"] == pytest.approx(38.0)
        assert result["crm.revenue"] == 150000

        detail = result["crm.detail_rows"]
        assert len(detail) == 2
        assert detail[0]["campaign_type"] == "Manual"

    def test_fallback_to_tracker(self, transformer):
        raw = _make_raw_data([
            {"channel": "EMAIL", "revenue": 5000, "orders": 100,
             "sessions": 1000, "cost": 500, "new_customers": 40},
        ])
        sources = {"tracker": {"RAW DATA": raw}}
        result = transformer._transform_crm(sources)

        assert result["crm.revenue"] == 5000
        assert result["crm.cvr"] == pytest.approx(10.0)  # 100/1000 * 100

    def test_no_data(self, transformer):
        result = transformer._transform_crm({})
        assert result["crm.revenue"] is None


# ---------------------------------------------------------------------------
# Affiliate Performance
# ---------------------------------------------------------------------------

class TestTransformAffiliate:
    def test_kpis_from_tracker(self, transformer):
        raw = _make_raw_data([
            {"year": 2026, "month": 1, "channel": "AFFILIATE",
             "revenue": 3000, "orders": 60, "sessions": 1000,
             "cost": 300, "new_customers": 25},
            {"year": 2025, "month": 1, "channel": "AFFILIATE",
             "revenue": 2500, "orders": 50, "sessions": 900,
             "cost": 250, "new_customers": 20},
        ])
        sources = {"tracker": {"RAW DATA": raw}}
        result = transformer._transform_affiliate(sources)

        assert result["affiliate.revenue"] == 3000
        assert result["affiliate.orders"] == 60
        # YoY: (3000-2500)/2500 * 100 = 20%
        assert result["affiliate.revenue_vs_ly"] == pytest.approx(20.0)

    def test_publisher_table(self, transformer):
        aff_data = pd.DataFrame({
            "Dimension 1": ["1001", "1002", "Grand Total"],
            "Dimension 2": ["Publisher A", "Publisher B", "Grand Total"],
            "Dimension 3": ["Total", "Total", "Total"],
            "Influencer Filter": ["Affiliate", "Affiliate", "Total"],
            "Revenue (Analysis)": [5000.0, 3000.0, 8000.0],
            "Revenue (vs. Comp)": [0.10, -0.05, 0.04],
            "Total Commission (Analysis)": [500.0, 300.0, 800.0],
            "CoS (Analysis)": [0.10, 0.10, 0.10],
            "Orders (Analysis)": [100.0, 60.0, 160.0],
            "CVR (Analysis)": [0.05, 0.04, 0.045],
            "Sessions (Analysis)": [2000.0, 1500.0, 3500.0],
            "AOV (Analysis)": [50.0, 50.0, 50.0],
        })
        sources = {"tracker": {"RAW DATA": _make_raw_data([])}, "affiliate": aff_data}
        result = transformer._transform_affiliate(sources)

        rows = result["affiliate.publisher_rows"]
        assert len(rows) == 2  # Grand Total excluded
        assert rows[0]["publisher_name"] == "Publisher A"
        assert rows[0]["revenue"] == 5000.0

    def test_no_data(self, transformer):
        result = transformer._transform_affiliate({})
        assert result["affiliate.revenue"] is None


# ---------------------------------------------------------------------------
# SEO Performance
# ---------------------------------------------------------------------------

class TestTransformSEO:
    def test_seo_kpis(self, transformer):
        raw = _make_raw_data([
            {"year": 2026, "month": 1, "channel": "ORGANIC",
             "revenue": 8000, "orders": 200, "sessions": 5000,
             "cost": 0, "new_customers": 80},
            {"year": 2025, "month": 1, "channel": "ORGANIC",
             "revenue": 7000, "orders": 180, "sessions": 4500,
             "cost": 0, "new_customers": 70},
        ])
        sources = {"tracker": {"RAW DATA": raw}}
        result = transformer._transform_seo(sources)

        assert result["seo.revenue"] == 8000
        assert result["seo.sessions"] == 5000
        assert result["seo.orders"] == 200
        # CVR = 200/5000 * 100 = 4.0%
        assert result["seo.cvr"] == pytest.approx(4.0)
        # YoY revenue: (8000-7000)/7000*100 = 14.29%
        assert result["seo.revenue_vs_ly"] == pytest.approx(14.286, rel=0.01)

    def test_no_data(self, transformer):
        result = transformer._transform_seo({})
        assert result["seo.revenue"] is None


# ---------------------------------------------------------------------------
# Full transform
# ---------------------------------------------------------------------------

class TestFullTransform:
    def test_all_keys_present(self, transformer):
        """Full transform with minimal data still produces all required keys."""
        raw = _make_raw_data([
            {"channel": "DIRECT", "revenue": 1000, "orders": 50,
             "sessions": 500, "cost": 100, "new_customers": 20},
        ])
        sources = {"tracker": {"RAW DATA": raw}}
        result = transformer.transform(sources)

        # Check key prefixes exist
        prefixes = ["cover.", "exec.", "daily.", "promo.", "product.",
                     "crm.", "affiliate.", "seo."]
        for prefix in prefixes:
            matching = [k for k in result if k.startswith(prefix)]
            assert len(matching) > 0, f"No keys with prefix {prefix}"

    def test_empty_sources(self, transformer):
        """Transform with no sources returns payload with None/empty values."""
        result = transformer.transform({})
        assert result["cover.total_revenue"] is None
        assert result["exec.performance_rows"] == []
        assert result["daily.dates"] == []
        assert result["promo.rows"] == []
        assert result["product.rows"] == []
        assert result["crm.revenue"] is None
        assert result["affiliate.revenue"] is None
        assert result["seo.revenue"] is None
