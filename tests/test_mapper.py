"""Tests for the data-to-template mapper."""

import math

import pandas as pd
import pytest

from src.processor.mapper import (
    DataMapper,
    MappingResult,
    REPORT_CHANNELS,
    _clean,
    safe_divide,
    variance_pct,
)
from src.schema.monthly_report import build_monthly_report_schema


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def schema():
    return build_monthly_report_schema()


def _make_raw_data(year, month, days=None, channels=None):
    """Build a synthetic RAW DATA DataFrame.

    Each channel gets one row per day with deterministic values.
    """
    if channels is None:
        channels = REPORT_CHANNELS
    if days is None:
        days = list(range(1, 29))  # 28 days default

    rows = []
    for day in days:
        for ch in channels:
            rows.append({
                "COS Year": year,
                "COS Month": month,
                "COS Day": day,
                "COS Channel": ch,
                "COS Locale": "en_US",
                "COS Orders": 10,
                "COS New Customers": 5,
                "COS COS%": 0.10,
                "COS CAC": 20.0,
                "COS CPA": 10.0,
                "COS Cost": 100.0,
                "COS Revenue": 1000.0,
                "COS Sessions": 200,
                "COS AOV": 100.0,
                "COS Conversion": 0.05,
            })
    return pd.DataFrame(rows)


def _make_targets(year, month, channels=None):
    """Build a synthetic Targets DataFrame."""
    if channels is None:
        channels = REPORT_CHANNELS
    import calendar
    num_days = calendar.monthrange(year, month)[1]

    rows = []
    for day in range(1, num_days + 1):
        for ch in channels:
            rows.append({
                "Target_Type_Id": "Daily",
                "Date": pd.Timestamp(year, month, day),
                "Site_Id": "No 7",
                "Locale_Id": "en_US",
                "Channel_Id": ch,
                "Notes": float("nan"),
                "Gross_Revenue_Target": 900.0,
                "Net_Revenue_Target": 850.0,
                "Marketing_Spend_Target": 90.0,
                "Session_Target": 180,
                "Order_Target": 9,
                "New_Customer_Target": 4,
            })
    return pd.DataFrame(rows)


def _make_tracker(year, month, ly_year=None, ly_month=None):
    """Build a tracker dict with RAW DATA for current + prior year."""
    if ly_year is None:
        ly_year = year - 1
    if ly_month is None:
        ly_month = month
    raw_cur = _make_raw_data(year, month)
    raw_ly = _make_raw_data(ly_year, ly_month)
    raw = pd.concat([raw_cur, raw_ly], ignore_index=True)
    return {"RAW DATA": raw}


def _make_offer_performance(month):
    """Build a synthetic offer performance DataFrame."""
    rows = []
    promos = ["Promo A", "Promo B", "Promo C"]
    channels = ["AFFILIATE", "EMAIL", "PPC"]
    for promo in promos:
        for ch in channels:
            rows.append({
                "Dimension 1": promo,
                "Dimension 2": ch,
                "Dimension 3": str(month),
                "Dimension 4": "Total",
                "Redemptions": 1000.0,
                "% Change Redemptions": -0.10,
                "Revenue": 50000.0,
                "% Change Revenue": 0.15,
                "Discount Amount": 5000.0,
                "% Change Discount Amount": -0.05,
            })
    # Grand Total row
    rows.append({
        "Dimension 1": "Grand Total",
        "Dimension 2": "Total",
        "Dimension 3": "Total",
        "Dimension 4": "Total",
        "Redemptions": 9000.0,
        "% Change Redemptions": -0.10,
        "Revenue": 450000.0,
        "% Change Revenue": 0.15,
        "Discount Amount": 45000.0,
        "% Change Discount Amount": -0.05,
    })
    return pd.DataFrame(rows)


def _make_product_sales(month):
    """Build a synthetic product sales DataFrame."""
    rows = []
    products = ["Product X", "Product Y", "Product Z"]
    for prod in products:
        rows.append({
            "Dimension 1": prod,
            "Dimension 2": str(month),
            "Dimension 3": "Total",
            "Units (Analysis)": 500.0,
            "Units (Comparison)": 400.0,
            "Units (vs. Comp)": 0.25,
            "Total Revenue (Analysis)": 25000.0,
            "Total Revenue (Comparison)": 20000.0,
            "Total Revenue (vs. Comp)": 0.25,
            "AOV (Analysis)": 50.0,
            "Avg. Selling Price (Analysis)": 45.0,
            "Total Discount % (Analysis)": 0.10,
            "New Customers (Analysis)": 100.0,
        })
    # Grand Total
    rows.append({
        "Dimension 1": "Grand Total",
        "Dimension 2": "Total",
        "Dimension 3": "Total",
        "Units (Analysis)": 1500.0,
        "Units (Comparison)": 1200.0,
        "Units (vs. Comp)": 0.25,
        "Total Revenue (Analysis)": 75000.0,
        "Total Revenue (Comparison)": 60000.0,
        "Total Revenue (vs. Comp)": 0.25,
        "AOV (Analysis)": 50.0,
        "Avg. Selling Price (Analysis)": 45.0,
        "Total Discount % (Analysis)": 0.10,
        "New Customers (Analysis)": 300.0,
    })
    return pd.DataFrame(rows)


def _make_crm():
    """Build a synthetic CRM DataFrame."""
    return pd.DataFrame([
        {
            "Col A": "Grand Total",
            "Col B": "Total",
            "Col C": "Total",
            "Emails Sent": 50000,
            "Emails Sent vs Comp": -0.05,
            "Open Rate": 0.25,
            "Open Rate vs Comp": 0.02,
            "Click-Through Rate": 0.08,
            "Click-Through Rate vs Comp": -0.01,
            "Sessions": 10000,
            "Sessions vs Comp": 0.10,
            "Orders": 500,
            "Orders vs Comp": 0.20,
            "CVR": 0.05,
            "CVR vs Comp": 0.005,
            "Revenue": 75000.0,
            "Revenue vs Comp": 0.15,
            "AOV": 150.0,
            "AOV vs Comp": -0.03,
        },
        {
            "Col A": "Grand Total",
            "Col B": "Total",
            "Col C": "Manual",
            "Emails Sent": 30000,
            "Emails Sent vs Comp": -0.03,
            "Open Rate": 0.22,
            "Open Rate vs Comp": 0.01,
            "Click-Through Rate": 0.07,
            "Click-Through Rate vs Comp": -0.005,
            "Sessions": 6000,
            "Sessions vs Comp": 0.08,
            "Orders": 300,
            "Orders vs Comp": 0.18,
            "CVR": 0.05,
            "CVR vs Comp": 0.004,
            "Revenue": 45000.0,
            "Revenue vs Comp": 0.12,
            "AOV": 150.0,
            "AOV vs Comp": -0.02,
        },
        {
            "Col A": "Grand Total",
            "Col B": "Total",
            "Col C": "Automated",
            "Emails Sent": 20000,
            "Emails Sent vs Comp": -0.08,
            "Open Rate": 0.30,
            "Open Rate vs Comp": 0.03,
            "Click-Through Rate": 0.10,
            "Click-Through Rate vs Comp": -0.02,
            "Sessions": 4000,
            "Sessions vs Comp": 0.12,
            "Orders": 200,
            "Orders vs Comp": 0.22,
            "CVR": 0.05,
            "CVR vs Comp": 0.006,
            "Revenue": 30000.0,
            "Revenue vs Comp": 0.20,
            "AOV": 150.0,
            "AOV vs Comp": -0.05,
        },
    ])


def _make_affiliate():
    """Build a synthetic affiliate publisher DataFrame."""
    rows = [
        {
            "Dimension 1": "Grand Total",
            "Dimension 2": "All Publishers",
            "Dimension 3": "Total",
            "Influencer Filter": "Total",
            "Revenue (Analysis)": 100000.0,
            "Revenue (Comparison)": 80000.0,
            "Revenue (vs Comp)": 0.25,
            "Cost (Analysis)": 10000.0,
            "Cost (Comparison)": 9000.0,
            "Cost (vs Comp)": 0.111,
            "CoS (Analysis)": 0.10,
            "CoS (vs Comp)": -0.01,
            "Orders (Analysis)": 1000,
            "Orders (vs Comp)": 0.20,
            "CVR (Analysis)": 0.05,
            "CVR (vs Comp)": 0.005,
            "Sessions (Analysis)": 20000,
            "AOV (Analysis)": 100.0,
            "Total Commission (Analysis)": 8000.0,
        },
    ]
    # Add some publisher rows
    for i, name in enumerate(["Publisher A", "Publisher B", "Publisher C"]):
        rows.append({
            "Dimension 1": str(1000 + i),
            "Dimension 2": name,
            "Dimension 3": "Total",
            "Influencer Filter": "Affiliate",
            "Revenue (Analysis)": 30000.0 - i * 5000,
            "Revenue (Comparison)": 25000.0 - i * 5000,
            "Revenue (vs Comp)": 0.20,
            "Cost (Analysis)": 3000.0 - i * 500,
            "Cost (Comparison)": 2500.0 - i * 500,
            "Cost (vs Comp)": 0.20,
            "CoS (Analysis)": 0.10,
            "CoS (vs Comp)": -0.005,
            "Orders (Analysis)": 300 - i * 50,
            "Orders (vs Comp)": 0.15,
            "CVR (Analysis)": 0.05,
            "CVR (vs Comp)": 0.003,
            "Sessions (Analysis)": 6000 - i * 1000,
            "AOV (Analysis)": 100.0,
            "Total Commission (Analysis)": 2500.0 - i * 400,
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# safe_divide
# ---------------------------------------------------------------------------

class TestSafeDivide:
    def test_normal(self):
        assert safe_divide(10, 2) == 5.0

    def test_zero_denominator(self):
        assert math.isnan(safe_divide(10, 0))

    def test_none_denominator(self):
        assert math.isnan(safe_divide(10, None))

    def test_nan_denominator(self):
        assert math.isnan(safe_divide(10, float("nan")))

    def test_custom_default(self):
        assert safe_divide(10, 0, default=0.0) == 0.0

    def test_negative(self):
        assert safe_divide(-10, 5) == -2.0

    def test_both_float(self):
        assert safe_divide(7.5, 2.5) == 3.0


# ---------------------------------------------------------------------------
# variance_pct
# ---------------------------------------------------------------------------

class TestVariancePct:
    def test_positive_change(self):
        # 120 vs 100 = +20%
        assert variance_pct(120, 100) == pytest.approx(0.20)

    def test_negative_change(self):
        # 80 vs 100 = -20%
        assert variance_pct(80, 100) == pytest.approx(-0.20)

    def test_no_change(self):
        assert variance_pct(100, 100) == 0.0

    def test_both_zero(self):
        assert variance_pct(0, 0) == 0.0

    def test_prior_zero_current_nonzero(self):
        assert math.isnan(variance_pct(100, 0))

    def test_nan_input(self):
        assert math.isnan(variance_pct(float("nan"), 100))

    def test_none_input(self):
        assert math.isnan(variance_pct(None, 100))

    def test_large_growth(self):
        # 300 vs 100 = +200%
        assert variance_pct(300, 100) == pytest.approx(2.0)


# ---------------------------------------------------------------------------
# _clean
# ---------------------------------------------------------------------------

class TestClean:
    def test_none(self):
        assert _clean(None) is None

    def test_nan(self):
        assert _clean(float("nan")) is None

    def test_inf(self):
        assert _clean(float("inf")) is None

    def test_normal_float(self):
        assert _clean(3.14) == 3.14

    def test_normal_int(self):
        assert _clean(42) == 42

    def test_string(self):
        assert _clean("hello") == "hello"

    def test_list(self):
        assert _clean([1, 2, 3]) == [1, 2, 3]

    def test_custom_default(self):
        assert _clean(None, default=0) == 0


# ---------------------------------------------------------------------------
# DataMapper — Cover KPIs
# ---------------------------------------------------------------------------

class TestMapCover:
    def test_cover_with_tracker_and_targets(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({
            "tracker": _make_tracker(2026, 1),
            "targets": _make_targets(2026, 1),
        })
        p = result.payload

        assert p["cover.report_title"] == "No7 US x THGi Monthly eComm Report"
        assert p["cover.report_period"] == "January 2026 Overview"

        # 9 channels × 28 days × 1000 = 252,000
        assert p["cover.total_revenue"] == 9 * 28 * 1000.0
        # 9 channels × 28 days × 10 = 2,520
        assert p["cover.total_orders"] == 9 * 28 * 10
        # AOV = 252000 / 2520 = 100
        assert p["cover.aov"] == pytest.approx(100.0)
        # New customers = 9 × 28 × 5 = 1260
        assert p["cover.new_customers"] == 9 * 28 * 5

        # vs-target variances should be present
        assert p["cover.revenue_vs_target"] is not None

    def test_cover_no_targets(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({"tracker": _make_tracker(2026, 1)})
        p = result.payload

        assert p["cover.total_revenue"] == 9 * 28 * 1000.0
        # vs-target not calculated
        assert "cover.revenue_vs_target" not in p or p.get("cover.revenue_vs_target") is None

    def test_cover_no_tracker(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({})

        assert any("RAW DATA" in w for w in result.warnings)
        # Title should still be set
        assert result.payload["cover.report_title"] is not None


# ---------------------------------------------------------------------------
# DataMapper — Executive Summary
# ---------------------------------------------------------------------------

class TestMapExecutiveSummary:
    def test_exec_rows_count(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({
            "tracker": _make_tracker(2026, 1),
            "targets": _make_targets(2026, 1),
        })
        rows = result.payload["exec.performance_rows"]
        # TOTAL + 9 channels = 10 rows
        assert len(rows) == 10
        assert rows[0]["channel"] == "TOTAL"
        assert rows[1]["channel"] == "AFFILIATE"

    def test_exec_row_metrics(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({
            "tracker": _make_tracker(2026, 1),
        })
        total_row = result.payload["exec.performance_rows"][0]

        # Same actuals as cover: 9 channels × 28 days
        assert total_row["revenue"] == 9 * 28 * 1000.0
        assert total_row["orders"] == 9 * 28 * 10
        # YoY: same values → 0%
        assert total_row["revenue_vs_ly"] == 0.0

    def test_exec_channel_row(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({"tracker": _make_tracker(2026, 1)})

        aff_row = result.payload["exec.performance_rows"][1]
        assert aff_row["channel"] == "AFFILIATE"
        # 28 days × 1000 = 28,000
        assert aff_row["revenue"] == 28 * 1000.0
        # CVR = orders / sessions = (28*10) / (28*200) = 0.05
        assert aff_row["cvr"] == pytest.approx(0.05)

    def test_exec_no_tracker(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({})
        assert result.payload["exec.performance_rows"] == []


# ---------------------------------------------------------------------------
# DataMapper — Daily Performance
# ---------------------------------------------------------------------------

class TestMapDailyPerformance:
    def test_daily_series_length(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({
            "tracker": _make_tracker(2026, 1),
            "targets": _make_targets(2026, 1),
        })
        p = result.payload
        # January has 31 days
        assert len(p["daily.dates"]) == 31
        assert len(p["daily.revenue_actual"]) == 31
        assert len(p["daily.revenue_target"]) == 31
        assert len(p["daily.revenue_ly"]) == 31

    def test_daily_dates_format(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({"tracker": _make_tracker(2026, 1)})
        dates = result.payload["daily.dates"]
        assert dates[0] == "1/1"
        assert dates[-1] == "1/31"

    def test_daily_actual_values(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({"tracker": _make_tracker(2026, 1)})
        actuals = result.payload["daily.revenue_actual"]
        # Each day in our synthetic data has all 9 channels × $1000 = $9000
        # But we only have 28 days of data (days 1-28)
        assert actuals[0] == 9 * 1000.0  # day 1
        # Days 29-31 have no data → 0
        assert actuals[28] == 0.0

    def test_daily_gauge(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({
            "tracker": _make_tracker(2026, 1),
            "targets": _make_targets(2026, 1),
        })
        p = result.payload
        achieved = p["daily.revenue_achieved_pct"]
        remaining = p["daily.revenue_remaining_pct"]
        assert 0.0 <= achieved <= 1.0
        assert 0.0 <= remaining <= 1.0
        assert achieved + remaining == pytest.approx(1.0)


# ---------------------------------------------------------------------------
# DataMapper — Promotion Performance
# ---------------------------------------------------------------------------

class TestMapPromotionPerformance:
    def test_promo_rows(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({
            "tracker": _make_tracker(2026, 1),
            "offer_performance": _make_offer_performance(1),
        })
        rows = result.payload["promo.rows"]
        # 3 promos × 3 channels = 9, all ≤ 15 limit
        assert len(rows) == 9

    def test_promo_row_fields(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({
            "offer_performance": _make_offer_performance(1),
        })
        row = result.payload["promo.rows"][0]
        assert "promotion_name" in row
        assert "channel" in row
        assert "redemptions" in row
        assert "revenue" in row
        assert row["revenue"] == 50000.0

    def test_promo_no_data(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({})
        assert result.payload["promo.rows"] == []
        assert any("Offer performance" in w for w in result.warnings)

    def test_promo_wrong_month(self, schema):
        mapper = DataMapper(schema, month=2, year=2026)
        result = mapper.map({
            "offer_performance": _make_offer_performance(1),  # month 1 data
        })
        # No rows for month 2
        assert result.payload["promo.rows"] == []


# ---------------------------------------------------------------------------
# DataMapper — Product Performance
# ---------------------------------------------------------------------------

class TestMapProductPerformance:
    def test_product_rows(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({
            "product_sales": _make_product_sales(1),
        })
        rows = result.payload["product.rows"]
        # 3 products (Grand Total excluded)
        assert len(rows) == 3

    def test_product_row_fields(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({
            "product_sales": _make_product_sales(1),
        })
        row = result.payload["product.rows"][0]
        assert row["product_name"] in ("Product X", "Product Y", "Product Z")
        assert row["revenue"] == 25000.0
        assert row["units"] == 500.0
        assert row["aov"] == 50.0
        assert row["discount_pct"] == 0.10

    def test_product_no_data(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({})
        assert result.payload["product.rows"] == []


# ---------------------------------------------------------------------------
# DataMapper — CRM Performance
# ---------------------------------------------------------------------------

class TestMapCRM:
    def test_crm_kpis(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({"crm": _make_crm()})
        p = result.payload

        assert p["crm.emails_sent"] == 50000
        assert p["crm.emails_sent_vs_ly"] == -0.05
        assert p["crm.open_rate"] == 0.25
        assert p["crm.revenue"] == 75000.0
        assert p["crm.aov"] == 150.0

    def test_crm_detail_rows(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({"crm": _make_crm()})
        rows = result.payload["crm.detail_rows"]

        assert len(rows) == 2
        types = {r["campaign_type"] for r in rows}
        assert types == {"Manual", "Automated"}

    def test_crm_no_data(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({})
        assert any("CRM" in w for w in result.warnings)


# ---------------------------------------------------------------------------
# DataMapper — Affiliate Performance
# ---------------------------------------------------------------------------

class TestMapAffiliate:
    def test_affiliate_kpis(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({"affiliate": _make_affiliate()})
        p = result.payload

        assert p["affiliate.revenue"] == 100000.0
        assert p["affiliate.revenue_vs_ly"] == 0.25
        assert p["affiliate.orders"] == 1000
        assert p["affiliate.cos"] == 0.10

    def test_affiliate_roas(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({"affiliate": _make_affiliate()})
        p = result.payload
        # ROAS = 100000 / 10000 = 10.0
        assert p["affiliate.roas"] == pytest.approx(10.0)

    def test_affiliate_publisher_rows(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({"affiliate": _make_affiliate()})
        rows = result.payload["affiliate.publisher_rows"]
        assert len(rows) == 3
        # Sorted by revenue descending
        assert rows[0]["publisher_name"] == "Publisher A"
        assert rows[0]["revenue"] == 30000.0

    def test_affiliate_tracker_fallback(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({"tracker": _make_tracker(2026, 1)})
        p = result.payload

        # Should use tracker fallback
        assert any("fallback" in w for w in result.warnings)
        assert p["affiliate.revenue"] == 28 * 1000.0
        assert p["affiliate.publisher_rows"] == []


# ---------------------------------------------------------------------------
# DataMapper — SEO Performance
# ---------------------------------------------------------------------------

class TestMapSEO:
    def test_seo_kpis(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({"tracker": _make_tracker(2026, 1)})
        p = result.payload

        # ORGANIC channel: 28 days × 1000 = 28000 revenue
        assert p["seo.revenue"] == 28 * 1000.0
        assert p["seo.sessions"] == 28 * 200
        assert p["seo.orders"] == 28 * 10
        # YoY: same values → 0%
        assert p["seo.revenue_vs_ly"] == 0.0

    def test_seo_no_tracker(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({})
        assert any("seo" in w for w in result.warnings)


# ---------------------------------------------------------------------------
# DataMapper — Static slides
# ---------------------------------------------------------------------------

class TestMapStaticSlides:
    def test_toc(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({})
        items = result.payload["toc.items"]
        assert isinstance(items, list)
        assert len(items) == 9
        assert "eComm Performance" in items

    def test_dividers(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({})
        p = result.payload
        assert p["divider.ecomm_title"] == "eComm Performance"
        assert p["divider.channels_title"] == "Channel Deep Dives"
        assert p["divider.outlook_title"] == "Outlook"

    def test_manual_slides(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({})
        p = result.payload
        assert p["upcoming.title"] == "Upcoming Promotions"
        assert p["upcoming.rows"] == []
        assert p["next_steps.title"] == "Next Steps"


# ---------------------------------------------------------------------------
# DataMapper — Coverage and validation
# ---------------------------------------------------------------------------

class TestCoverage:
    def test_full_sources_coverage(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({
            "tracker": _make_tracker(2026, 1),
            "targets": _make_targets(2026, 1),
            "offer_performance": _make_offer_performance(1),
            "product_sales": _make_product_sales(1),
            "crm": _make_crm(),
            "affiliate": _make_affiliate(),
        })
        # With all sources, coverage should be high
        assert result.coverage > 0.7

    def test_empty_sources_coverage(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({})
        # Static slides still provide some coverage
        assert 0.0 < result.coverage < 1.0

    def test_result_type(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({})
        assert isinstance(result, MappingResult)
        assert isinstance(result.payload, dict)
        assert isinstance(result.warnings, list)
        assert isinstance(result.coverage, float)


# ---------------------------------------------------------------------------
# DataMapper — Edge cases
# ---------------------------------------------------------------------------

class TestEdgeCases:
    def test_empty_raw_data(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({
            "tracker": {"RAW DATA": pd.DataFrame()},
        })
        assert any("RAW DATA" in w for w in result.warnings)

    def test_different_month(self, schema):
        mapper = DataMapper(schema, month=6, year=2026)
        result = mapper.map({
            "tracker": _make_tracker(2026, 6),
        })
        assert result.payload["cover.report_period"] == "June 2026 Overview"

    def test_february_days(self, schema):
        mapper = DataMapper(schema, month=2, year=2026)
        result = mapper.map({
            "tracker": _make_tracker(2026, 2),
        })
        # 2026 is not a leap year → February has 28 days
        assert len(result.payload["daily.dates"]) == 28

    def test_leap_year_february(self, schema):
        mapper = DataMapper(schema, month=2, year=2024)
        result = mapper.map({
            "tracker": _make_tracker(2024, 2),
        })
        # 2024 is a leap year → February has 29 days
        assert len(result.payload["daily.dates"]) == 29

    def test_map_is_idempotent(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        sources = {"tracker": _make_tracker(2026, 1)}
        r1 = mapper.map(sources)
        r2 = mapper.map(sources)
        assert r1.payload == r2.payload
        assert r1.coverage == r2.coverage

    def test_payload_has_no_nan(self, schema):
        mapper = DataMapper(schema, month=1, year=2026)
        result = mapper.map({
            "tracker": _make_tracker(2026, 1),
            "targets": _make_targets(2026, 1),
            "offer_performance": _make_offer_performance(1),
            "product_sales": _make_product_sales(1),
            "crm": _make_crm(),
            "affiliate": _make_affiliate(),
        })
        for key, val in result.payload.items():
            if isinstance(val, float):
                assert not math.isnan(val), f"{key} is NaN"
                assert not math.isinf(val), f"{key} is inf"
            elif isinstance(val, list):
                for item in val:
                    if isinstance(item, float):
                        assert not math.isnan(item), f"{key} list has NaN"
                    elif isinstance(item, dict):
                        for k, v in item.items():
                            if isinstance(v, float):
                                assert not math.isnan(v), (
                                    f"{key}[].{k} is NaN")
