"""Data-to-template mapper for Presentation Builder.

Transforms ingested data sources into a flat payload dict keyed by the
data_key identifiers defined in the template schema.  Sits between the
ingestion layer (src.processor.ingestion) and the generator layer.

Each slide in the schema declares the data_keys it needs.  The mapper
reads ingested DataFrames, aggregates/filters/calculates derived metrics,
and writes every key into a single dict the generator can consume.

Usage::

    from src.processor.mapper import DataMapper
    from src.schema.monthly_report import build_monthly_report_schema

    schema = build_monthly_report_schema()
    mapper = DataMapper(schema, month=1, year=2026)
    result = mapper.map({
        "tracker": ingest_tracker("tracker.xlsx"),
        "targets": ingest_targets("targets.csv"),
        "offer_performance": ingest_offer_performance("offers.csv"),
        "product_sales": ingest_product_sales("products.csv"),
        "crm": ingest_crm("crm.xlsm"),
        "affiliate": ingest_affiliate("affiliate.xlsm"),
    })
    print(result.coverage)       # 0.0-1.0
    print(result.warnings)       # list of warning strings
    payload = result.payload     # dict[str, Any]
"""

import calendar
import math
from dataclasses import dataclass, field
from typing import Any

import pandas as pd

from src.schema.models import TemplateSchema


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

MONTH_NAMES = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December",
}

REPORT_CHANNELS = [
    "AFFILIATE", "DIRECT", "DISPLAY", "EMAIL",
    "INFLUENCER", "ORGANIC", "OTHER", "PPC", "SOCIAL",
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def safe_divide(numerator, denominator, default=float("nan")):
    """Divide *numerator* by *denominator*, returning *default* on failure."""
    try:
        if denominator is None:
            return default
        denom = float(denominator)
        if denom == 0 or math.isnan(denom):
            return default
        result = float(numerator) / denom
        if math.isinf(result):
            return default
        return result
    except (TypeError, ValueError, ZeroDivisionError):
        return default


def variance_pct(current, prior):
    """Percentage change: ``(current - prior) / prior``.

    Returns ``NaN`` when *prior* is zero or either value is missing.
    Returns ``0.0`` when both values are equal (including both zero).
    """
    try:
        c, p = float(current), float(prior)
    except (TypeError, ValueError):
        return float("nan")
    if math.isnan(c) or math.isnan(p):
        return float("nan")
    if c == p:
        return 0.0
    if p == 0:
        return float("nan")
    return (c - p) / p


def _clean(val, default=None):
    """Convert NaN / None / numpy scalars to clean Python types."""
    if val is None:
        return default
    try:
        if pd.isna(val):
            return default
    except (ValueError, TypeError):
        pass  # non-scalar (list, dict, etc.) — pass through
    if hasattr(val, "item"):  # numpy scalar → native Python
        val = val.item()
    if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
        return default
    return val


# ---------------------------------------------------------------------------
# Result container
# ---------------------------------------------------------------------------

@dataclass
class MappingResult:
    """Output of :meth:`DataMapper.map`."""
    payload: dict[str, Any]
    warnings: list[str]
    coverage: float  # fraction of required data_keys populated (0.0-1.0)


# ---------------------------------------------------------------------------
# DataMapper
# ---------------------------------------------------------------------------

class DataMapper:
    """Map ingested data sources to template schema data keys.

    Parameters
    ----------
    schema : TemplateSchema
        The template schema whose ``data_keys`` define the payload shape.
    month : int
        Report month (1-12).
    year : int
        Report year (e.g. 2026).
    """

    def __init__(self, schema: TemplateSchema, month: int, year: int):
        self.schema = schema
        self.month = month
        self.year = year
        self.month_name = MONTH_NAMES[month]
        self.days_in_month = calendar.monthrange(year, month)[1]
        self._payload: dict[str, Any] = {}
        self._warnings: list[str] = []

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def map(self, sources: dict) -> MappingResult:
        """Map all *sources* to template data keys.

        Parameters
        ----------
        sources : dict
            Keyed by source type.  Expected keys (all optional):

            - ``"tracker"``  — ``dict[str, DataFrame]`` from
              :func:`ingest_tracker`
            - ``"targets"``  — ``DataFrame`` from :func:`ingest_targets`
            - ``"offer_performance"``  — ``DataFrame``
            - ``"product_sales"``  — ``DataFrame``
            - ``"crm"``  — ``DataFrame``
            - ``"affiliate"``  — ``DataFrame``
            - ``"historical"``  — ``DataFrame``

        Returns
        -------
        MappingResult
        """
        self._payload = {}
        self._warnings = []

        tracker = sources.get("tracker", {})
        targets = sources.get("targets")
        offer_perf = sources.get("offer_performance")
        product_sales = sources.get("product_sales")
        crm_data = sources.get("crm")
        affiliate_data = sources.get("affiliate")

        self._map_cover(tracker, targets)
        self._map_executive_summary(tracker, targets)
        self._map_daily_performance(tracker, targets)
        self._map_promotion_performance(offer_perf)
        self._map_product_performance(product_sales)
        self._map_crm(crm_data)
        self._map_affiliate(affiliate_data, tracker)
        self._map_seo(tracker)
        self._map_static_slides()

        return MappingResult(
            payload=dict(self._payload),
            warnings=list(self._warnings),
            coverage=self._compute_coverage(),
        )

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _put(self, key: str, value):
        """Set a payload key, converting numpy/NaN to clean Python."""
        if value is None:
            self._payload[key] = None
            return
        try:
            if pd.isna(value):
                self._payload[key] = None
                return
        except (ValueError, TypeError):
            pass
        if hasattr(value, "item"):
            value = value.item()
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            self._payload[key] = None
            return
        self._payload[key] = value

    def _warn(self, msg: str):
        self._warnings.append(msg)

    def _compute_coverage(self) -> float:
        required = self.schema.all_data_keys()
        if not required:
            return 1.0
        covered = sum(
            1 for k in required
            if k in self._payload and self._payload[k] is not None
        )
        return covered / len(required)

    # ---- RAW DATA helpers ----

    def _get_raw(self, tracker: dict) -> pd.DataFrame | None:
        """Return the RAW DATA sheet or ``None``."""
        raw = tracker.get("RAW DATA")
        if raw is None or (hasattr(raw, "empty") and raw.empty):
            return None
        return raw

    def _filter_raw(self, raw: pd.DataFrame, month: int, year: int,
                    channel: str | None = None) -> pd.DataFrame:
        """Filter RAW DATA by period and optionally channel."""
        mask = (raw["COS Month"] == month) & (raw["COS Year"] == year)
        if channel:
            mask = mask & (raw["COS Channel"] == channel)
        return raw[mask]

    def _agg_raw(self, filtered: pd.DataFrame) -> dict:
        """Sum core metrics from filtered RAW DATA rows."""
        if filtered.empty:
            return {
                "revenue": 0.0, "orders": 0.0, "sessions": 0.0,
                "new_customers": 0.0, "cost": 0.0,
            }
        return {
            "revenue": filtered["COS Revenue"].sum(),
            "orders": filtered["COS Orders"].sum(),
            "sessions": filtered["COS Sessions"].sum(),
            "new_customers": filtered["COS New Customers"].sum(),
            "cost": filtered["COS Cost"].sum(),
        }

    def _enrich(self, agg: dict) -> dict:
        """Add derived metrics (AOV, CVR, COS, ROAS) to aggregated dict."""
        agg["aov"] = safe_divide(agg["revenue"], agg["orders"])
        agg["cvr"] = safe_divide(agg["orders"], agg["sessions"])
        agg["cos"] = safe_divide(agg["cost"], agg["revenue"])
        agg["roas"] = safe_divide(agg["revenue"], agg["cost"])
        return agg

    # ---- Target helpers ----

    def _filter_targets(self, targets: pd.DataFrame,
                        channel: str | None = None) -> pd.DataFrame:
        """Filter targets to current report month/year."""
        mask = (
            (targets["Date"].dt.month == self.month)
            & (targets["Date"].dt.year == self.year)
        )
        if channel:
            mask = mask & (targets["Channel_Id"] == channel)
        return targets[mask]

    def _agg_targets(self, filtered: pd.DataFrame) -> dict:
        """Sum target values from filtered rows."""
        if filtered.empty:
            return {
                "revenue": 0.0, "orders": 0.0, "sessions": 0.0,
                "new_customers": 0.0, "spend": 0.0,
            }
        return {
            "revenue": filtered["Gross_Revenue_Target"].sum(),
            "orders": filtered["Order_Target"].sum(),
            "sessions": filtered["Session_Target"].sum(),
            "new_customers": filtered["New_Customer_Target"].sum(),
            "spend": filtered["Marketing_Spend_Target"].sum(),
        }

    # ---- Column lookup ----

    @staticmethod
    def _find_col(df: pd.DataFrame, *candidates: str) -> str | None:
        """Return first column name from *candidates* that exists in *df*."""
        for c in candidates:
            if c in df.columns:
                return c
        return None

    # ======================================================================
    # SLIDE MAPPERS
    # ======================================================================

    # ------------------------------------------------------------------
    # Slide 0: Cover + KPIs
    # ------------------------------------------------------------------

    def _map_cover(self, tracker: dict, targets):
        self._put("cover.report_title",
                   "No7 US x THGi Monthly eComm Report")
        self._put("cover.report_period",
                   f"{self.month_name} {self.year} Overview")

        raw = self._get_raw(tracker)
        if raw is None:
            self._warn("cover: RAW DATA sheet not available — "
                       "KPI values cannot be calculated")
            return

        # Current-month totals
        cur = self._enrich(self._agg_raw(
            self._filter_raw(raw, self.month, self.year)))

        self._put("cover.total_revenue", cur["revenue"])
        self._put("cover.total_orders", cur["orders"])
        self._put("cover.aov", cur["aov"])
        self._put("cover.new_customers", cur["new_customers"])
        self._put("cover.cvr", cur["cvr"])
        self._put("cover.cos", cur["cos"])

        # vs-Target variances
        if targets is not None and not targets.empty:
            tgt = self._agg_targets(self._filter_targets(targets))
            tgt["aov"] = safe_divide(tgt["revenue"], tgt["orders"])
            tgt["cvr"] = safe_divide(tgt["orders"], tgt["sessions"])
            tgt["cos"] = safe_divide(tgt["spend"], tgt["revenue"])

            self._put("cover.revenue_vs_target",
                       variance_pct(cur["revenue"], tgt["revenue"]))
            self._put("cover.orders_vs_target",
                       variance_pct(cur["orders"], tgt["orders"]))
            self._put("cover.aov_vs_target",
                       variance_pct(cur["aov"], tgt["aov"]))
            self._put("cover.nc_vs_target",
                       variance_pct(cur["new_customers"], tgt["new_customers"]))
            self._put("cover.cvr_vs_target",
                       variance_pct(cur["cvr"], tgt["cvr"]))
            self._put("cover.cos_vs_target",
                       variance_pct(cur["cos"], tgt["cos"]))
        else:
            self._warn("cover: Target data not available — "
                       "vs-target variances not calculated")

    # ------------------------------------------------------------------
    # Slide 3: Executive Summary
    # ------------------------------------------------------------------

    def _map_executive_summary(self, tracker: dict, targets):
        self._put("exec.title",
                   f"eComm Performance — {self.month_name} {self.year}")
        self._put("exec.performance_table", "Performance Summary")

        raw = self._get_raw(tracker)
        if raw is None:
            self._warn("exec: RAW DATA sheet not available")
            self._put("exec.performance_rows", [])
            self._put("exec.narrative", None)
            return

        rows = []
        for channel in ["TOTAL"] + REPORT_CHANNELS:
            if channel == "TOTAL":
                cur = self._enrich(self._agg_raw(
                    self._filter_raw(raw, self.month, self.year)))
                ly = self._enrich(self._agg_raw(
                    self._filter_raw(raw, self.month, self.year - 1)))
            else:
                cur = self._enrich(self._agg_raw(
                    self._filter_raw(raw, self.month, self.year, channel)))
                ly = self._enrich(self._agg_raw(
                    self._filter_raw(raw, self.month, self.year - 1, channel)))

            rev_vs_target = None
            if targets is not None and not targets.empty:
                if channel == "TOTAL":
                    tgt = self._agg_targets(self._filter_targets(targets))
                else:
                    tgt = self._agg_targets(
                        self._filter_targets(targets, channel))
                rev_vs_target = _clean(variance_pct(
                    cur["revenue"], tgt["revenue"]))

            rows.append({
                "channel": channel,
                "revenue": _clean(cur["revenue"]),
                "revenue_vs_target": rev_vs_target,
                "revenue_vs_ly": _clean(
                    variance_pct(cur["revenue"], ly["revenue"])),
                "orders": _clean(cur["orders"]),
                "sessions": _clean(cur["sessions"]),
                "cvr": _clean(cur["cvr"]),
                "aov": _clean(cur["aov"]),
                "cos": _clean(cur["cos"]),
                "new_customers": _clean(cur["new_customers"]),
            })

        self._put("exec.performance_rows", rows)
        # Narrative requires LLM/manual authoring
        self._put("exec.narrative", None)

    # ------------------------------------------------------------------
    # Slide 4: Daily Performance
    # ------------------------------------------------------------------

    def _map_daily_performance(self, tracker: dict, targets):
        self._put("daily.title",
                   f"Daily Performance — {self.month_name} {self.year}")
        self._put("daily.chart", "Daily Revenue vs Target")
        self._put("daily.campaign_table", "Campaign Activity")
        self._put("daily.revenue_gauge", "Revenue vs Target")

        raw = self._get_raw(tracker)
        if raw is None:
            self._warn("daily: RAW DATA sheet not available")
            return

        # Daily revenue from RAW DATA (current month)
        cur_month = self._filter_raw(raw, self.month, self.year)
        daily_actual = cur_month.groupby("COS Day")["COS Revenue"].sum()

        # LY daily
        ly_month = self._filter_raw(raw, self.month, self.year - 1)
        daily_ly = ly_month.groupby("COS Day")["COS Revenue"].sum()

        days = list(range(1, self.days_in_month + 1))
        dates = [f"{self.month}/{d}" for d in days]
        rev_actual = [_clean(daily_actual.get(d, 0.0), 0.0) for d in days]
        rev_ly = [_clean(daily_ly.get(d, 0.0), 0.0) for d in days]

        # Daily targets
        rev_target = [0.0] * len(days)
        if targets is not None and not targets.empty:
            tgt_month = self._filter_targets(targets)
            if not tgt_month.empty:
                daily_tgt = tgt_month.groupby(
                    tgt_month["Date"].dt.day
                )["Gross_Revenue_Target"].sum()
                rev_target = [
                    _clean(daily_tgt.get(d, 0.0), 0.0) for d in days
                ]

        self._put("daily.dates", dates)
        self._put("daily.revenue_actual", rev_actual)
        self._put("daily.revenue_target", rev_target)
        self._put("daily.revenue_ly", rev_ly)

        # Revenue gauge
        total_actual = sum(rev_actual)
        total_target = sum(rev_target)
        achieved = safe_divide(total_actual, total_target, 0.0)
        achieved = min(1.0, max(0.0, achieved))
        self._put("daily.revenue_achieved_pct", achieved)
        self._put("daily.revenue_remaining_pct", max(0.0, 1.0 - achieved))

        # Campaign activity — requires manual input
        self._put("daily.campaign_rows", [])

    # ------------------------------------------------------------------
    # Slide 5: Promotion Performance
    # ------------------------------------------------------------------

    def _map_promotion_performance(self, offer_df):
        self._put("promo.title",
                   f"Promotion Performance — {self.month_name} {self.year}")
        self._put("promo.table", "Promotion Summary")

        if offer_df is None or offer_df.empty:
            self._warn("promo: Offer performance data not available")
            self._put("promo.rows", [])
            return

        dim1 = self._find_col(offer_df, "Dimension 1")
        dim2 = self._find_col(offer_df, "Dimension 2")
        dim3 = self._find_col(offer_df, "Dimension 3")
        dim4 = self._find_col(offer_df, "Dimension 4")

        if not dim1:
            self._warn("promo: Dimension 1 column not found")
            self._put("promo.rows", [])
            return

        # Filter: current month totals, exclude Grand Total
        mask = offer_df[dim1].astype(str).str.strip() != "Grand Total"
        if dim3:
            mask = mask & (
                offer_df[dim3].astype(str).str.strip() == str(self.month))
        if dim4:
            mask = mask & (
                offer_df[dim4].astype(str).str.strip() == "Total")

        monthly = offer_df[mask].copy()

        if monthly.empty:
            self._warn(f"promo: No offer data for month {self.month}")
            self._put("promo.rows", [])
            return

        # Sort by revenue descending
        rev_col = self._find_col(monthly, "Revenue")
        if rev_col:
            monthly = monthly.sort_values(rev_col, ascending=False)

        rows = []
        for _, r in monthly.head(15).iterrows():
            rows.append({
                "promotion_name": _clean(r.get(dim1)),
                "channel": _clean(r.get(dim2)) if dim2 else None,
                "redemptions": _clean(r.get("Redemptions")),
                "redemptions_vs_ly": _clean(r.get("% Change Redemptions")),
                "revenue": _clean(r.get("Revenue")),
                "revenue_vs_ly": _clean(r.get("% Change Revenue")),
                "discount_amount": _clean(r.get("Discount Amount")),
            })

        self._put("promo.rows", rows)

    # ------------------------------------------------------------------
    # Slide 6: Product Performance
    # ------------------------------------------------------------------

    def _map_product_performance(self, product_df):
        self._put("product.title",
                   f"Product Performance — {self.month_name} {self.year}")
        self._put("product.table", "Product Summary")

        if product_df is None or product_df.empty:
            self._warn("product: Product sales data not available")
            self._put("product.rows", [])
            return

        dim1 = self._find_col(product_df, "Dimension 1")
        dim2 = self._find_col(product_df, "Dimension 2")
        dim3 = self._find_col(product_df, "Dimension 3")

        if not dim1:
            self._warn("product: Dimension 1 column not found")
            self._put("product.rows", [])
            return

        # Filter: current month, sub-period total, exclude Grand Total
        mask = product_df[dim1].astype(str).str.strip() != "Grand Total"
        if dim2:
            mask = mask & (
                product_df[dim2].astype(str).str.strip() == str(self.month))
        if dim3:
            mask = mask & (
                product_df[dim3].astype(str).str.strip() == "Total")

        monthly = product_df[mask].copy()

        if monthly.empty:
            self._warn(f"product: No product data for month {self.month}")
            self._put("product.rows", [])
            return

        # Sort by revenue descending
        rev_col = self._find_col(
            monthly,
            "Total Revenue (Analysis)",
            "Product Revenue (Analysis)",
        )
        if rev_col:
            monthly = monthly.sort_values(rev_col, ascending=False)

        rows = []
        for _, r in monthly.head(15).iterrows():
            rev = _clean(r.get("Total Revenue (Analysis)"))
            if rev is None:
                rev = _clean(r.get("Product Revenue (Analysis)"))
            rev_vs = _clean(r.get("Total Revenue (vs. Comp)"))
            if rev_vs is None:
                rev_vs = _clean(r.get("Product Revenue (vs. Comp)"))

            rows.append({
                "product_name": _clean(r.get(dim1)),
                "units": _clean(r.get("Units (Analysis)")),
                "units_vs_ly": _clean(r.get("Units (vs. Comp)")),
                "revenue": rev,
                "revenue_vs_ly": rev_vs,
                "aov": _clean(r.get("AOV (Analysis)")),
                "avg_selling_price": _clean(
                    r.get("Avg. Selling Price (Analysis)")),
                "discount_pct": _clean(
                    r.get("Total Discount % (Analysis)")),
                "new_customers": _clean(
                    r.get("New Customers (Analysis)")),
            })

        self._put("product.rows", rows)

    # ------------------------------------------------------------------
    # Slide 8: CRM Performance
    # ------------------------------------------------------------------

    def _map_crm(self, crm_df):
        self._put("crm.title",
                   f"CRM Performance — {self.month_name} {self.year}")
        self._put("crm.detail_table", "CRM Detail")

        if crm_df is None or crm_df.empty:
            self._warn("crm: CRM data not available")
            return

        # Separate dimension columns from metric columns
        metric_keywords = {
            "Emails Sent", "Emails Delivered", "Emails Opened",
            "Open Rate", "Unsubscribe", "Link Clicks",
            "Click-Through Rate", "Sessions", "Landing Page Bounce Rate",
            "Orders", "CVR", "Revenue", "AOV",
        }
        dim_cols = [
            c for c in crm_df.columns
            if not any(k in c for k in metric_keywords)
               and "vs Comp" not in c
        ]

        # Locate the total row and campaign-type detail rows.
        # Campaign-type rows (Manual, Automated) are detail rows even if
        # they also contain "Grand Total" in another dimension column.
        total_row = None
        detail_rows = []

        for _, row in crm_df.iterrows():
            dim_vals = [
                str(row.get(c, "")).strip().lower() for c in dim_cols
            ]
            has_campaign_type = any(
                t in dim_vals for t in ("manual", "automated")
            )
            if has_campaign_type:
                detail_rows.append(row)
            elif total_row is None and (
                "grand total" in dim_vals or "total" in dim_vals
            ):
                total_row = row

        if total_row is None:
            # Fall back to a row labelled "Total" or first row
            for _, row in crm_df.iterrows():
                dim_vals = [
                    str(row.get(c, "")).strip().lower() for c in dim_cols
                ]
                if "total" in dim_vals:
                    total_row = row
                    break
            if total_row is None:
                total_row = crm_df.iloc[0]

        # Helpers for metric + vs-comp lookup
        def _val(row, metric):
            return _clean(row.get(metric))

        def _vs(row, metric):
            return _clean(row.get(f"{metric} vs Comp"))

        # Headline KPIs
        self._put("crm.emails_sent", _val(total_row, "Emails Sent"))
        self._put("crm.emails_sent_vs_ly", _vs(total_row, "Emails Sent"))
        self._put("crm.open_rate", _val(total_row, "Open Rate"))
        self._put("crm.open_rate_vs_ly", _vs(total_row, "Open Rate"))
        self._put("crm.ctr", _val(total_row, "Click-Through Rate"))
        self._put("crm.ctr_vs_ly", _vs(total_row, "Click-Through Rate"))
        self._put("crm.revenue", _val(total_row, "Revenue"))
        self._put("crm.revenue_vs_ly", _vs(total_row, "Revenue"))
        self._put("crm.cvr", _val(total_row, "CVR"))
        self._put("crm.cvr_vs_ly", _vs(total_row, "CVR"))
        self._put("crm.aov", _val(total_row, "AOV"))
        self._put("crm.aov_vs_ly", _vs(total_row, "AOV"))

        # Detail table rows
        table_rows = []
        for row in detail_rows:
            campaign_type = None
            for c in dim_cols:
                v = str(row.get(c, "")).strip()
                if v.lower() in ("manual", "automated"):
                    campaign_type = v
                    break

            table_rows.append({
                "campaign_type": campaign_type or "Unknown",
                "emails_sent": _val(row, "Emails Sent"),
                "open_rate": _val(row, "Open Rate"),
                "ctr": _val(row, "Click-Through Rate"),
                "sessions": _val(row, "Sessions"),
                "orders": _val(row, "Orders"),
                "cvr": _val(row, "CVR"),
                "revenue": _val(row, "Revenue"),
                "aov": _val(row, "AOV"),
                "revenue_vs_ly": _vs(row, "Revenue"),
            })

        self._put("crm.detail_rows", table_rows)

    # ------------------------------------------------------------------
    # Slide 9: Affiliate Performance
    # ------------------------------------------------------------------

    def _map_affiliate(self, affiliate_df, tracker: dict):
        self._put("affiliate.title",
                   f"Affiliate Performance — {self.month_name} {self.year}")
        self._put("affiliate.publisher_table", "Top Publishers")

        if affiliate_df is None or affiliate_df.empty:
            self._warn("affiliate: Affiliate Excel not available — "
                       "using tracker fallback")
            self._map_affiliate_from_tracker(tracker)
            return

        dim1 = self._find_col(affiliate_df, "Dimension 1")
        dim2 = self._find_col(affiliate_df, "Dimension 2")
        dim3 = self._find_col(affiliate_df, "Dimension 3")

        # Helper: get (Analysis) and (vs Comp) values
        def _a(row, metric):
            return _clean(row.get(f"{metric} (Analysis)"))

        def _vs(row, metric):
            return _clean(row.get(f"{metric} (vs Comp)"))

        # Grand Total row for headline KPIs
        grand_total = None
        if dim1:
            gt_mask = (
                affiliate_df[dim1].astype(str).str.strip() == "Grand Total"
            )
            if dim3:
                gt_mask = gt_mask & (
                    affiliate_df[dim3].astype(str).str.strip() == "Total"
                )
            if gt_mask.any():
                grand_total = affiliate_df[gt_mask].iloc[0]

        if grand_total is None:
            grand_total = affiliate_df.iloc[0]
            self._warn("affiliate: Could not identify Grand Total row")

        rev = _a(grand_total, "Revenue")
        cost = _a(grand_total, "Cost")
        cos_val = _a(grand_total, "CoS")

        self._put("affiliate.revenue", rev)
        self._put("affiliate.revenue_vs_ly", _vs(grand_total, "Revenue"))
        self._put("affiliate.cos",
                   cos_val if cos_val is not None
                   else _clean(safe_divide(cost, rev)))
        self._put("affiliate.cos_vs_ly", _vs(grand_total, "CoS"))
        self._put("affiliate.roas", _clean(safe_divide(rev, cost)))
        self._put("affiliate.orders", _a(grand_total, "Orders"))
        self._put("affiliate.orders_vs_ly", _vs(grand_total, "Orders"))
        self._put("affiliate.cvr", _a(grand_total, "CVR"))
        self._put("affiliate.cvr_vs_ly", _vs(grand_total, "CVR"))

        # ROAS vs LY — derive from comparison values
        rev_comp = _clean(grand_total.get("Revenue (Comparison)"))
        cost_comp = _clean(grand_total.get("Cost (Comparison)"))
        if rev_comp is not None and cost_comp is not None:
            roas_ly = safe_divide(rev_comp, cost_comp)
            roas_cur = safe_divide(rev, cost)
            self._put("affiliate.roas_vs_ly",
                       _clean(variance_pct(roas_cur, roas_ly)))
        else:
            self._put("affiliate.roas_vs_ly", None)

        # Publisher table — top publishers by revenue
        if dim1 and dim2:
            pub_mask = (
                affiliate_df[dim1].astype(str).str.strip() != "Grand Total"
            )
            if dim3:
                pub_mask = pub_mask & (
                    affiliate_df[dim3].astype(str).str.strip() == "Total"
                )
            inf_col = self._find_col(affiliate_df, "Influencer Filter")
            if inf_col:
                pub_mask = pub_mask & (
                    affiliate_df[inf_col].astype(str).str.strip().isin(
                        ["Affiliate", "Total"]
                    )
                )

            publishers = affiliate_df[pub_mask].copy()
            rev_col = "Revenue (Analysis)"
            if rev_col in publishers.columns:
                publishers = publishers.sort_values(
                    rev_col, ascending=False)

            rows = []
            for _, r in publishers.head(10).iterrows():
                commission = _a(r, "Total Commission")
                if commission is None:
                    commission = _a(r, "Cost")
                rows.append({
                    "publisher_name": _clean(r.get(dim2)),
                    "revenue": _a(r, "Revenue"),
                    "revenue_vs_ly": _vs(r, "Revenue"),
                    "commission": commission,
                    "cos": _a(r, "CoS"),
                    "orders": _a(r, "Orders"),
                    "cvr": _a(r, "CVR"),
                    "sessions": _a(r, "Sessions"),
                    "aov": _a(r, "AOV"),
                })
            self._put("affiliate.publisher_rows", rows)
        else:
            self._put("affiliate.publisher_rows", [])

    def _map_affiliate_from_tracker(self, tracker: dict):
        """Fallback: derive affiliate KPIs from tracker RAW DATA."""
        raw = self._get_raw(tracker)
        if raw is None:
            self._warn("affiliate: No fallback data available")
            self._put("affiliate.publisher_rows", [])
            return

        cur = self._enrich(self._agg_raw(
            self._filter_raw(raw, self.month, self.year, "AFFILIATE")))
        ly = self._enrich(self._agg_raw(
            self._filter_raw(raw, self.month, self.year - 1, "AFFILIATE")))

        self._put("affiliate.revenue", cur["revenue"])
        self._put("affiliate.revenue_vs_ly",
                   _clean(variance_pct(cur["revenue"], ly["revenue"])))
        self._put("affiliate.cos", cur["cos"])
        self._put("affiliate.cos_vs_ly",
                   _clean(variance_pct(cur["cos"], ly["cos"])))
        self._put("affiliate.roas", cur["roas"])
        self._put("affiliate.roas_vs_ly",
                   _clean(variance_pct(cur["roas"], ly["roas"])))
        self._put("affiliate.orders", cur["orders"])
        self._put("affiliate.orders_vs_ly",
                   _clean(variance_pct(cur["orders"], ly["orders"])))
        self._put("affiliate.cvr", cur["cvr"])
        self._put("affiliate.cvr_vs_ly",
                   _clean(variance_pct(cur["cvr"], ly["cvr"])))
        self._put("affiliate.publisher_rows", [])

    # ------------------------------------------------------------------
    # Slide 10: SEO Performance
    # ------------------------------------------------------------------

    def _map_seo(self, tracker: dict):
        self._put("seo.title",
                   f"SEO Performance — {self.month_name} {self.year}")

        raw = self._get_raw(tracker)
        if raw is None:
            self._warn("seo: RAW DATA sheet not available")
            return

        cur = self._enrich(self._agg_raw(
            self._filter_raw(raw, self.month, self.year, "ORGANIC")))
        ly = self._enrich(self._agg_raw(
            self._filter_raw(raw, self.month, self.year - 1, "ORGANIC")))

        self._put("seo.revenue", cur["revenue"])
        self._put("seo.revenue_vs_ly",
                   _clean(variance_pct(cur["revenue"], ly["revenue"])))
        self._put("seo.sessions", cur["sessions"])
        self._put("seo.sessions_vs_ly",
                   _clean(variance_pct(cur["sessions"], ly["sessions"])))
        self._put("seo.cvr", cur["cvr"])
        self._put("seo.cvr_vs_ly",
                   _clean(variance_pct(cur["cvr"], ly["cvr"])))
        self._put("seo.orders", cur["orders"])
        self._put("seo.orders_vs_ly",
                   _clean(variance_pct(cur["orders"], ly["orders"])))
        self._put("seo.aov", cur["aov"])
        self._put("seo.aov_vs_ly",
                   _clean(variance_pct(cur["aov"], ly["aov"])))
        # Narrative requires manual/LLM authoring
        self._put("seo.narrative", None)

    # ------------------------------------------------------------------
    # Static / Manual slides
    # ------------------------------------------------------------------

    def _map_static_slides(self):
        """Populate static content for TOC, dividers, and manual slides."""
        # Slide 1: TOC
        self._put("toc.items", [
            "eComm Performance",
            "Daily Performance",
            "Promotion Performance",
            "Product Performance",
            "Channel Deep Dives",
            "CRM Performance",
            "Affiliate Performance",
            "SEO Performance",
            "Outlook",
        ])

        # Section dividers
        self._put("divider.ecomm_title", "eComm Performance")
        self._put("divider.channels_title", "Channel Deep Dives")
        self._put("divider.outlook_title", "Outlook")

        # Slide 12: Upcoming Promotions (manual)
        self._put("upcoming.title", "Upcoming Promotions")
        self._put("upcoming.table", "Upcoming Promotions Calendar")
        self._put("upcoming.rows", [])

        # Slide 13: Next Steps (manual)
        self._put("next_steps.title", "Next Steps")
        self._put("next_steps.items", None)
