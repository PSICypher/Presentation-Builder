"""Data transformation module for Presentation Builder.

Takes raw ingested DataFrames (from the ingestion module) and produces a
structured data payload — a flat dict keyed by the schema's data keys.
This is the bridge between raw data ingestion and the presentation generator.

The ``DataTransformer`` class accepts a ``ReportContext`` (month, year) and
a dict of ingested data sources, then computes all derived metrics,
aggregations, rankings, and variances needed for the 14-slide monthly report.

Data key convention (matches monthly_report.py schema):
    <slide_name>.<field>         — single value
    <slide_name>.<table>_rows    — list of row dicts for a table
    <slide_name>.<chart>_series  — list of series data for a chart
    <slide_name>.<chart>_cats    — category labels for a chart
"""

import calendar
import math
from dataclasses import dataclass, field

import pandas as pd

from .ingestion import parse_numeric, parse_percentage


# ---------------------------------------------------------------------------
# Report context
# ---------------------------------------------------------------------------

@dataclass
class ReportContext:
    """Context for the report being generated."""
    year: int
    month: int

    @property
    def month_name(self) -> str:
        return calendar.month_name[self.month]

    @property
    def days_in_month(self) -> int:
        return calendar.monthrange(self.year, self.month)[1]

    @property
    def prior_year(self) -> int:
        return self.year - 1


# ---------------------------------------------------------------------------
# Safe math helpers
# ---------------------------------------------------------------------------

def _safe_div(numerator, denominator, default=float("nan")):
    """Divide safely, returning default on zero/NaN denominator."""
    if denominator is None or denominator == 0:
        return default
    if isinstance(denominator, float) and math.isnan(denominator):
        return default
    if isinstance(numerator, float) and math.isnan(numerator):
        return default
    return numerator / denominator


def _safe_pct_change(current, prior, default=float("nan")):
    """Calculate percentage change: (current - prior) / prior * 100.

    Returns a value in percentage points (e.g. 5.2 for +5.2%).
    """
    if prior is None or prior == 0:
        return default
    if isinstance(prior, float) and math.isnan(prior):
        return default
    if isinstance(current, float) and math.isnan(current):
        return default
    return (current - prior) / prior * 100


def _safe_ppt_change(current, prior, default=float("nan")):
    """Calculate percentage-point change: (current - prior) * 100.

    Both current and prior are expected as decimals (e.g. 0.09 = 9%).
    Returns value in percentage points (e.g. 2.5 for +2.5 ppts).
    """
    if current is None or prior is None:
        return default
    if isinstance(current, float) and math.isnan(current):
        return default
    if isinstance(prior, float) and math.isnan(prior):
        return default
    return (current - prior) * 100


def _decimal_to_pct(value, default=float("nan")):
    """Convert a decimal rate (0.09) to percentage (9.0)."""
    if value is None:
        return default
    if isinstance(value, float) and math.isnan(value):
        return default
    return value * 100


def _nan_to_none(value):
    """Convert NaN to None for JSON-safe output."""
    if isinstance(value, float) and math.isnan(value):
        return None
    return value


# ---------------------------------------------------------------------------
# Channel constants
# ---------------------------------------------------------------------------

CHANNELS = [
    "AFFILIATE", "DIRECT", "DISPLAY", "EMAIL",
    "INFLUENCER", "ORGANIC", "OTHER", "PPC", "SOCIAL",
]


# ---------------------------------------------------------------------------
# DataTransformer
# ---------------------------------------------------------------------------

class DataTransformer:
    """Transforms raw ingested data into a presentation data payload.

    Args:
        context: ReportContext with year and month for the report.

    Usage::

        ctx = ReportContext(year=2026, month=1)
        transformer = DataTransformer(ctx)
        payload = transformer.transform({
            "tracker": ingest_tracker("path/to/tracker.xlsx"),
            "targets": ingest_targets("path/to/targets.csv"),
            "offer_performance": ingest_offer_performance("path/to/offers.csv"),
            "product_sales": ingest_product_sales("path/to/products.csv"),
            "crm": ingest_crm("path/to/crm.xlsm"),
            "affiliate": ingest_affiliate("path/to/affiliate.xlsm"),
            "historical": ingest_historical("path/to/historical.csv"),
        })
    """

    def __init__(self, context: ReportContext):
        self.ctx = context

    def transform(self, sources: dict) -> dict:
        """Transform all sources into a complete data payload.

        Args:
            sources: Dict mapping source name to ingested data.
                Expected keys: 'tracker', 'targets', 'offer_performance',
                'product_sales', 'crm', 'affiliate', 'historical'.
                Any key may be missing — the transformer will produce
                None/N/A values for unavailable data.

        Returns:
            Flat dict mapping schema data keys to values.
        """
        payload = {}
        payload.update(self._transform_cover(sources))
        payload.update(self._transform_executive_summary(sources))
        payload.update(self._transform_daily(sources))
        payload.update(self._transform_promotions(sources))
        payload.update(self._transform_products(sources))
        payload.update(self._transform_crm(sources))
        payload.update(self._transform_affiliate(sources))
        payload.update(self._transform_seo(sources))
        return payload

    # -------------------------------------------------------------------
    # Internal: extract tracker RAW DATA for a specific period
    # -------------------------------------------------------------------

    def _get_raw_data(self, sources: dict) -> pd.DataFrame | None:
        """Get the RAW DATA sheet from the tracker."""
        tracker = sources.get("tracker")
        if tracker is None:
            return None
        return tracker.get("RAW DATA")

    def _filter_month(self, df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
        """Filter RAW DATA to a specific year and month."""
        mask = (df["COS Year"] == year) & (df["COS Month"] == month)
        return df[mask]

    def _aggregate_channel(self, df: pd.DataFrame) -> dict:
        """Aggregate RAW DATA rows into channel-level totals.

        Returns dict mapping channel name -> metrics dict.
        """
        result = {}
        for channel in CHANNELS:
            cdf = df[df["COS Channel"] == channel]
            if cdf.empty:
                result[channel] = self._empty_channel_metrics()
                continue
            result[channel] = self._compute_channel_metrics(cdf)

        # Total across all channels
        result["Total"] = self._compute_channel_metrics(df)
        return result

    def _compute_channel_metrics(self, df: pd.DataFrame) -> dict:
        """Compute aggregated metrics from a filtered DataFrame."""
        revenue = df["COS Revenue"].sum()
        orders = df["COS Orders"].sum()
        sessions = df["COS Sessions"].sum()
        cost = df["COS Cost"].sum()
        new_customers = df["COS New Customers"].sum()

        return {
            "revenue": revenue,
            "orders": int(orders),
            "sessions": int(sessions),
            "cost": cost,
            "new_customers": int(new_customers),
            "aov": _safe_div(revenue, orders),
            "cvr": _safe_div(orders, sessions),  # decimal
            "cos": _safe_div(cost, revenue),      # decimal
            "cac": _safe_div(cost, new_customers),
            "cpa": _safe_div(cost, orders),
            "roas": _safe_div(revenue, cost),
        }

    def _empty_channel_metrics(self) -> dict:
        """Return zeroed metrics for a channel with no data."""
        return {
            "revenue": 0.0,
            "orders": 0,
            "sessions": 0,
            "cost": 0.0,
            "new_customers": 0,
            "aov": float("nan"),
            "cvr": float("nan"),
            "cos": float("nan"),
            "cac": float("nan"),
            "cpa": float("nan"),
            "roas": float("nan"),
        }

    # -------------------------------------------------------------------
    # Internal: target aggregation
    # -------------------------------------------------------------------

    def _get_monthly_targets(self, sources: dict) -> dict:
        """Aggregate target data for the report month by channel.

        Returns dict mapping channel -> {revenue, sessions, orders, new_customers}.
        """
        targets_df = sources.get("targets")
        if targets_df is None:
            return {}

        # Filter to report month
        if "Date" in targets_df.columns:
            mask = (
                (targets_df["Date"].dt.year == self.ctx.year)
                & (targets_df["Date"].dt.month == self.ctx.month)
            )
        else:
            return {}

        mdf = targets_df[mask]
        if mdf.empty:
            return {}

        result = {}
        for channel in CHANNELS:
            cdf = mdf[mdf["Channel_Id"] == channel] if "Channel_Id" in mdf.columns else pd.DataFrame()
            if cdf.empty:
                continue
            result[channel] = {
                "revenue": cdf["Net_Revenue_Target"].sum() if "Net_Revenue_Target" in cdf.columns else 0.0,
                "spend": cdf["Marketing_Spend_Target"].sum() if "Marketing_Spend_Target" in cdf.columns else 0.0,
                "sessions": int(cdf["Session_Target"].sum()) if "Session_Target" in cdf.columns else 0,
                "orders": int(cdf["Order_Target"].sum()) if "Order_Target" in cdf.columns else 0,
                "new_customers": int(cdf["New_Customer_Target"].sum()) if "New_Customer_Target" in cdf.columns else 0,
            }

        # Total
        total = {
            "revenue": mdf["Net_Revenue_Target"].sum() if "Net_Revenue_Target" in mdf.columns else 0.0,
            "spend": mdf["Marketing_Spend_Target"].sum() if "Marketing_Spend_Target" in mdf.columns else 0.0,
            "sessions": int(mdf["Session_Target"].sum()) if "Session_Target" in mdf.columns else 0,
            "orders": int(mdf["Order_Target"].sum()) if "Order_Target" in mdf.columns else 0,
            "new_customers": int(mdf["New_Customer_Target"].sum()) if "New_Customer_Target" in mdf.columns else 0,
        }
        result["Total"] = total
        return result

    # -------------------------------------------------------------------
    # Cover KPIs (Slide 0)
    # -------------------------------------------------------------------

    def _transform_cover(self, sources: dict) -> dict:
        """Generate cover slide KPIs."""
        raw = self._get_raw_data(sources)
        targets = self._get_monthly_targets(sources)

        payload = {
            "cover.report_title": "No7 US x THGi Monthly eComm Report",
            "cover.report_period": f"{self.ctx.month_name} {self.ctx.year} Overview",
        }

        if raw is None or raw.empty:
            for key in ["total_revenue", "total_orders", "aov", "new_customers", "cvr", "cos"]:
                payload[f"cover.{key}"] = None
            for key in ["revenue_vs_target", "orders_vs_target", "aov_vs_target",
                        "nc_vs_target", "cvr_vs_target", "cos_vs_target"]:
                payload[f"cover.{key}"] = None
            return payload

        current = self._filter_month(raw, self.ctx.year, self.ctx.month)
        metrics = self._compute_channel_metrics(current)

        payload["cover.total_revenue"] = metrics["revenue"]
        payload["cover.total_orders"] = metrics["orders"]
        payload["cover.aov"] = _nan_to_none(metrics["aov"])
        payload["cover.new_customers"] = metrics["new_customers"]
        payload["cover.cvr"] = _nan_to_none(_decimal_to_pct(metrics["cvr"]))
        payload["cover.cos"] = _nan_to_none(_decimal_to_pct(metrics["cos"]))

        # vs Target variances
        total_target = targets.get("Total", {})
        if total_target:
            target_revenue = total_target.get("revenue", 0)
            target_orders = total_target.get("orders", 0)
            target_nc = total_target.get("new_customers", 0)
            target_sessions = total_target.get("sessions", 0)

            payload["cover.revenue_vs_target"] = _nan_to_none(
                _safe_pct_change(metrics["revenue"], target_revenue)
            )
            payload["cover.orders_vs_target"] = _nan_to_none(
                _safe_pct_change(metrics["orders"], target_orders)
            )

            target_aov = _safe_div(target_revenue, target_orders)
            payload["cover.aov_vs_target"] = _nan_to_none(
                _safe_pct_change(metrics["aov"], target_aov)
            )
            payload["cover.nc_vs_target"] = _nan_to_none(
                _safe_pct_change(metrics["new_customers"], target_nc)
            )

            target_cvr = _safe_div(target_orders, target_sessions)
            payload["cover.cvr_vs_target"] = _nan_to_none(
                _safe_ppt_change(metrics["cvr"], target_cvr)
            )

            target_cos = _safe_div(total_target.get("spend", 0), target_revenue)
            payload["cover.cos_vs_target"] = _nan_to_none(
                _safe_ppt_change(metrics["cos"], target_cos)
            )
        else:
            for key in ["revenue_vs_target", "orders_vs_target", "aov_vs_target",
                        "nc_vs_target", "cvr_vs_target", "cos_vs_target"]:
                payload[f"cover.{key}"] = None

        return payload

    # -------------------------------------------------------------------
    # Executive Summary (Slide 3)
    # -------------------------------------------------------------------

    def _transform_executive_summary(self, sources: dict) -> dict:
        """Generate executive summary performance table."""
        raw = self._get_raw_data(sources)
        targets = self._get_monthly_targets(sources)

        payload = {
            "exec.title": "Executive Summary",
            "exec.narrative": "",
        }

        if raw is None or raw.empty:
            payload["exec.performance_rows"] = []
            return payload

        current = self._filter_month(raw, self.ctx.year, self.ctx.month)
        prior = self._filter_month(raw, self.ctx.prior_year, self.ctx.month)

        current_by_channel = self._aggregate_channel(current)
        prior_by_channel = self._aggregate_channel(prior)

        rows = []
        for ch in ["Total"] + CHANNELS:
            cur = current_by_channel.get(ch, self._empty_channel_metrics())
            pri = prior_by_channel.get(ch, self._empty_channel_metrics())
            tgt = targets.get(ch, {})

            target_revenue = tgt.get("revenue", 0)

            row = {
                "channel": ch,
                "revenue": cur["revenue"],
                "revenue_vs_target": _nan_to_none(
                    _safe_pct_change(cur["revenue"], target_revenue)
                ) if target_revenue else None,
                "revenue_vs_ly": _nan_to_none(
                    _safe_pct_change(cur["revenue"], pri["revenue"])
                ),
                "orders": cur["orders"],
                "sessions": cur["sessions"],
                "cvr": _nan_to_none(_decimal_to_pct(cur["cvr"])),
                "aov": _nan_to_none(cur["aov"]),
                "cos": _nan_to_none(_decimal_to_pct(cur["cos"])),
                "new_customers": cur["new_customers"],
            }
            rows.append(row)

        payload["exec.performance_rows"] = rows
        return payload

    # -------------------------------------------------------------------
    # Daily Performance (Slide 4)
    # -------------------------------------------------------------------

    def _transform_daily(self, sources: dict) -> dict:
        """Generate daily performance chart data."""
        payload = {
            "daily.title": "Daily Performance",
            "daily.dates": [],
            "daily.revenue_actual": [],
            "daily.revenue_target": [],
            "daily.revenue_ly": [],
            "daily.campaign_rows": [],
            "daily.revenue_achieved_pct": None,
            "daily.revenue_remaining_pct": None,
        }

        tracker = sources.get("tracker")
        if tracker is None:
            return payload

        raw = tracker.get("RAW DATA")
        if raw is not None and not raw.empty:
            current = self._filter_month(raw, self.ctx.year, self.ctx.month)
            prior = self._filter_month(raw, self.ctx.prior_year, self.ctx.month)

            # Group by day
            if not current.empty:
                daily_actual = current.groupby("COS Day")["COS Revenue"].sum()
                dates = []
                actual_values = []
                ly_values = []

                daily_prior = (
                    prior.groupby("COS Day")["COS Revenue"].sum()
                    if not prior.empty else pd.Series(dtype=float)
                )

                for day in range(1, self.ctx.days_in_month + 1):
                    dates.append(f"{self.ctx.month}/{day}")
                    actual_values.append(
                        daily_actual.get(day, 0.0)
                    )
                    ly_values.append(
                        daily_prior.get(day, 0.0)
                    )

                payload["daily.dates"] = dates
                payload["daily.revenue_actual"] = actual_values
                payload["daily.revenue_ly"] = ly_values

        # Daily targets from target phasing
        targets_df = sources.get("targets")
        if targets_df is not None and "Date" in targets_df.columns:
            month_targets = targets_df[
                (targets_df["Date"].dt.year == self.ctx.year)
                & (targets_df["Date"].dt.month == self.ctx.month)
            ]
            if not month_targets.empty and "Net_Revenue_Target" in month_targets.columns:
                daily_targets = month_targets.groupby(
                    month_targets["Date"].dt.day
                )["Net_Revenue_Target"].sum()

                target_values = []
                for day in range(1, self.ctx.days_in_month + 1):
                    target_values.append(daily_targets.get(day, 0.0))
                payload["daily.revenue_target"] = target_values

                # Achievement gauge
                total_actual = sum(payload["daily.revenue_actual"])
                total_target = sum(target_values)
                if total_target > 0:
                    achieved = total_actual / total_target * 100
                    payload["daily.revenue_achieved_pct"] = min(achieved, 100.0)
                    payload["daily.revenue_remaining_pct"] = max(100.0 - achieved, 0.0)

        return payload

    # -------------------------------------------------------------------
    # Promotion Performance (Slide 5)
    # -------------------------------------------------------------------

    def _transform_promotions(self, sources: dict) -> dict:
        """Generate promotion performance table (top offers by revenue)."""
        payload = {
            "promo.title": "Promotion Performance",
            "promo.rows": [],
        }

        offer_df = sources.get("offer_performance")
        if offer_df is None or offer_df.empty:
            return payload

        # Filter to report month total rows
        dim_cols = [c for c in offer_df.columns if c.startswith("Dimension")]
        if not dim_cols:
            return payload

        # Get "Total" channel aggregation per offer for the report month
        # Dimension 1 = Offer, Dimension 2 = Channel, Dimension 3 = Month, Dimension 4 = Week
        month_str = str(self.ctx.month)

        has_dim3 = "Dimension 3" in offer_df.columns
        has_dim4 = "Dimension 4" in offer_df.columns

        # Filter to target month totals
        mask = pd.Series(True, index=offer_df.index)
        if has_dim3:
            mask = mask & (offer_df["Dimension 3"].astype(str) == month_str)
        if has_dim4:
            mask = mask & (offer_df["Dimension 4"].astype(str) == "Total")

        month_data = offer_df[mask].copy()

        if month_data.empty:
            return payload

        # Exclude Grand Total rows
        if "Dimension 1" in month_data.columns:
            month_data = month_data[
                month_data["Dimension 1"].astype(str) != "Grand Total"
            ]

        # Get per-offer totals (Dimension 2 = "Total" for all-channel aggregate)
        if "Dimension 2" in month_data.columns:
            total_rows = month_data[
                month_data["Dimension 2"].astype(str) == "Total"
            ]
        else:
            total_rows = month_data

        if total_rows.empty:
            # If no "Total" channel rows, aggregate across channels
            if "Dimension 1" in month_data.columns:
                total_rows = month_data.groupby("Dimension 1", as_index=False).agg({
                    c: "sum" for c in ["Redemptions", "Revenue", "Discount Amount"]
                    if c in month_data.columns
                })

        # Sort by revenue descending, take top 15
        if "Revenue" in total_rows.columns:
            sorted_rows = total_rows.sort_values("Revenue", ascending=False).head(15)
        else:
            sorted_rows = total_rows.head(15)

        rows = []
        for _, r in sorted_rows.iterrows():
            row = {
                "promotion_name": r.get("Dimension 1", ""),
                "channel": r.get("Dimension 2", "Total")
                    if "Dimension 2" in r.index else "Total",
                "redemptions": _nan_to_none(r.get("Redemptions")),
                "redemptions_vs_ly": _nan_to_none(
                    _decimal_to_pct(r.get("% Change Redemptions"))
                ) if "% Change Redemptions" in r.index else None,
                "revenue": _nan_to_none(r.get("Revenue")),
                "revenue_vs_ly": _nan_to_none(
                    _decimal_to_pct(r.get("% Change Revenue"))
                ) if "% Change Revenue" in r.index else None,
                "discount_amount": _nan_to_none(r.get("Discount Amount")),
            }
            rows.append(row)

        payload["promo.rows"] = rows
        return payload

    # -------------------------------------------------------------------
    # Product Performance (Slide 6)
    # -------------------------------------------------------------------

    def _transform_products(self, sources: dict) -> dict:
        """Generate product performance table (top 15 by revenue)."""
        payload = {
            "product.title": "Product Performance",
            "product.rows": [],
        }

        product_df = sources.get("product_sales")
        if product_df is None or product_df.empty:
            return payload

        # Filter to month total rows
        # Dimension 1 = Product, Dimension 2 = Month, Dimension 3 = Week
        month_str = str(self.ctx.month)

        has_dim2 = "Dimension 2" in product_df.columns
        has_dim3 = "Dimension 3" in product_df.columns

        mask = pd.Series(True, index=product_df.index)
        if has_dim2:
            mask = mask & (product_df["Dimension 2"].astype(str) == month_str)
        if has_dim3:
            mask = mask & (product_df["Dimension 3"].astype(str) == "Total")

        month_data = product_df[mask].copy()

        if month_data.empty:
            return payload

        # Exclude Grand Total
        if "Dimension 1" in month_data.columns:
            month_data = month_data[
                month_data["Dimension 1"].astype(str) != "Grand Total"
            ]

        # Sort by Product Revenue (Analysis) or Total Revenue (Analysis)
        revenue_col = None
        for col_name in ["Product Revenue (Analysis)", "Total Revenue (Analysis)"]:
            if col_name in month_data.columns:
                revenue_col = col_name
                break

        if revenue_col:
            month_data = month_data.sort_values(revenue_col, ascending=False)

        top15 = month_data.head(15)

        rows = []
        for _, r in top15.iterrows():
            def _get_analysis(metric):
                col = f"{metric} (Analysis)"
                return _nan_to_none(r.get(col)) if col in r.index else None

            def _get_vs_comp(metric):
                col = f"{metric} (vs. Comp)"
                val = r.get(col) if col in r.index else None
                if val is None or (isinstance(val, float) and math.isnan(val)):
                    return None
                # vs. Comp values are already parsed as decimals by ingestion
                return _nan_to_none(_decimal_to_pct(val))

            revenue = _get_analysis("Product Revenue") or _get_analysis("Total Revenue")
            units = _get_analysis("Units")
            orders = _get_analysis("Orders")

            row = {
                "product_name": r.get("Dimension 1", ""),
                "units": _nan_to_none(units),
                "units_vs_ly": _get_vs_comp("Units"),
                "revenue": _nan_to_none(revenue),
                "revenue_vs_ly": (
                    _get_vs_comp("Product Revenue") or _get_vs_comp("Total Revenue")
                ),
                "aov": _nan_to_none(_get_analysis("AOV")),
                "avg_selling_price": _nan_to_none(_get_analysis("Avg. Selling Price")),
                "discount_pct": _nan_to_none(
                    _decimal_to_pct(_get_analysis("Total Discount %"))
                ) if _get_analysis("Total Discount %") is not None else None,
                "new_customers": _nan_to_none(_get_analysis("New Customers")),
            }
            rows.append(row)

        payload["product.rows"] = rows
        return payload

    # -------------------------------------------------------------------
    # CRM Performance (Slide 8)
    # -------------------------------------------------------------------

    def _transform_crm(self, sources: dict) -> dict:
        """Generate CRM performance KPIs and detail table."""
        payload = {
            "crm.title": "CRM Performance",
            "crm.emails_sent": None,
            "crm.open_rate": None,
            "crm.ctr": None,
            "crm.revenue": None,
            "crm.cvr": None,
            "crm.aov": None,
            "crm.emails_sent_vs_ly": None,
            "crm.open_rate_vs_ly": None,
            "crm.ctr_vs_ly": None,
            "crm.revenue_vs_ly": None,
            "crm.cvr_vs_ly": None,
            "crm.aov_vs_ly": None,
            "crm.detail_rows": [],
        }

        crm_df = sources.get("crm")
        if crm_df is None or crm_df.empty:
            # Fallback to tracker EMAIL channel
            return self._transform_crm_from_tracker(sources, payload)

        # Find the campaign-type column — the one that contains both "Total"
        # and at least one non-Total value (e.g. "Manual", "Automated").
        type_col = None
        for col in crm_df.columns:
            vals = crm_df[col].astype(str).str.strip()
            has_total = (vals == "Total").any()
            has_non_total = (~vals.isin(["Total", "Grand Total", ""])).any()
            if has_total and has_non_total:
                type_col = col
                break

        if type_col is None:
            # Fallback: use the last dimension column (third column in CRM schema)
            type_col = crm_df.columns[min(2, len(crm_df.columns) - 1)]

        total_mask = crm_df[type_col].astype(str).str.strip() == "Total"
        total_row = crm_df[total_mask]

        if total_row.empty:
            return payload

        tr = total_row.iloc[0]

        # Extract metrics and their vs-comp values
        metric_map = {
            "emails_sent": ("Emails Sent", False),
            "open_rate": ("Open Rate", True),
            "ctr": ("Click-Through Rate", True),
            "revenue": ("Revenue", False),
            "cvr": ("CVR", True),
            "aov": ("AOV", False),
        }

        for key, (metric_name, is_rate) in metric_map.items():
            val = tr.get(metric_name)
            if val is not None and not (isinstance(val, float) and math.isnan(val)):
                if is_rate:
                    payload[f"crm.{key}"] = _decimal_to_pct(val)
                else:
                    payload[f"crm.{key}"] = val

            # vs LY
            vs_col = f"{metric_name} vs Comp"
            # Try alternate column name patterns
            if vs_col not in tr.index:
                for col in crm_df.columns:
                    if metric_name in col and "vs" in col.lower():
                        vs_col = col
                        break

            vs_val = tr.get(vs_col) if vs_col in tr.index else None
            if vs_val is not None and not (isinstance(vs_val, float) and math.isnan(vs_val)):
                payload[f"crm.{key}_vs_ly"] = _decimal_to_pct(vs_val)

        # Detail rows by campaign type
        detail_rows = []
        non_total = crm_df[~total_mask]
        for _, row in non_total.iterrows():
            campaign_type = str(row.get(type_col, "")).strip()
            if not campaign_type or campaign_type.lower() in ("grand total", "total"):
                continue
            detail = {
                "campaign_type": campaign_type,
                "emails_sent": _nan_to_none(row.get("Emails Sent")),
                "open_rate": _nan_to_none(
                    _decimal_to_pct(row.get("Open Rate"))
                ) if "Open Rate" in row.index else None,
                "ctr": _nan_to_none(
                    _decimal_to_pct(row.get("Click-Through Rate"))
                ) if "Click-Through Rate" in row.index else None,
                "sessions": _nan_to_none(row.get("Sessions")),
                "orders": _nan_to_none(row.get("Orders")),
                "cvr": _nan_to_none(
                    _decimal_to_pct(row.get("CVR"))
                ) if "CVR" in row.index else None,
                "revenue": _nan_to_none(row.get("Revenue")),
                "aov": _nan_to_none(row.get("AOV")),
                "revenue_vs_ly": _nan_to_none(
                    _decimal_to_pct(row.get("Revenue vs Comp"))
                ) if "Revenue vs Comp" in row.index else None,
            }
            detail_rows.append(detail)

        payload["crm.detail_rows"] = detail_rows
        return payload

    def _transform_crm_from_tracker(self, sources: dict, payload: dict) -> dict:
        """Fallback: derive CRM KPIs from tracker EMAIL channel."""
        raw = self._get_raw_data(sources)
        if raw is None or raw.empty:
            return payload

        current = self._filter_month(raw, self.ctx.year, self.ctx.month)
        prior = self._filter_month(raw, self.ctx.prior_year, self.ctx.month)

        email_current = current[current["COS Channel"] == "EMAIL"]
        email_prior = prior[prior["COS Channel"] == "EMAIL"]

        if email_current.empty:
            return payload

        cur = self._compute_channel_metrics(email_current)
        pri = self._compute_channel_metrics(email_prior) if not email_prior.empty else self._empty_channel_metrics()

        payload["crm.revenue"] = cur["revenue"]
        payload["crm.cvr"] = _nan_to_none(_decimal_to_pct(cur["cvr"]))
        payload["crm.aov"] = _nan_to_none(cur["aov"])
        payload["crm.revenue_vs_ly"] = _nan_to_none(
            _safe_pct_change(cur["revenue"], pri["revenue"])
        )
        payload["crm.cvr_vs_ly"] = _nan_to_none(
            _safe_ppt_change(cur["cvr"], pri["cvr"])
        )
        payload["crm.aov_vs_ly"] = _nan_to_none(
            _safe_pct_change(cur["aov"], pri["aov"])
        )

        return payload

    # -------------------------------------------------------------------
    # Affiliate Performance (Slide 9)
    # -------------------------------------------------------------------

    def _transform_affiliate(self, sources: dict) -> dict:
        """Generate affiliate performance KPIs and publisher table."""
        payload = {
            "affiliate.title": "Affiliate Performance",
            "affiliate.revenue": None,
            "affiliate.cos": None,
            "affiliate.roas": None,
            "affiliate.orders": None,
            "affiliate.cvr": None,
            "affiliate.revenue_vs_ly": None,
            "affiliate.cos_vs_ly": None,
            "affiliate.roas_vs_ly": None,
            "affiliate.orders_vs_ly": None,
            "affiliate.cvr_vs_ly": None,
            "affiliate.publisher_rows": [],
        }

        # KPIs from tracker AFFILIATE channel
        raw = self._get_raw_data(sources)
        if raw is not None and not raw.empty:
            current = self._filter_month(raw, self.ctx.year, self.ctx.month)
            prior = self._filter_month(raw, self.ctx.prior_year, self.ctx.month)

            aff_current = current[current["COS Channel"] == "AFFILIATE"]
            aff_prior = prior[prior["COS Channel"] == "AFFILIATE"]

            if not aff_current.empty:
                cur = self._compute_channel_metrics(aff_current)
                pri = (
                    self._compute_channel_metrics(aff_prior)
                    if not aff_prior.empty
                    else self._empty_channel_metrics()
                )

                payload["affiliate.revenue"] = cur["revenue"]
                payload["affiliate.cos"] = _nan_to_none(_decimal_to_pct(cur["cos"]))
                payload["affiliate.roas"] = _nan_to_none(cur["roas"])
                payload["affiliate.orders"] = cur["orders"]
                payload["affiliate.cvr"] = _nan_to_none(_decimal_to_pct(cur["cvr"]))

                payload["affiliate.revenue_vs_ly"] = _nan_to_none(
                    _safe_pct_change(cur["revenue"], pri["revenue"])
                )
                payload["affiliate.cos_vs_ly"] = _nan_to_none(
                    _safe_ppt_change(cur["cos"], pri["cos"])
                )
                payload["affiliate.roas_vs_ly"] = _nan_to_none(
                    _safe_pct_change(cur["roas"], pri["roas"])
                )
                payload["affiliate.orders_vs_ly"] = _nan_to_none(
                    _safe_pct_change(cur["orders"], pri["orders"])
                )
                payload["affiliate.cvr_vs_ly"] = _nan_to_none(
                    _safe_ppt_change(cur["cvr"], pri["cvr"])
                )

        # Publisher table from affiliate Excel
        aff_df = sources.get("affiliate")
        if aff_df is not None and not aff_df.empty:
            payload["affiliate.publisher_rows"] = self._build_publisher_rows(aff_df)

        return payload

    def _build_publisher_rows(self, aff_df: pd.DataFrame) -> list[dict]:
        """Build top publisher rows from affiliate data."""
        # Filter to Total sub-period, exclude Grand Total
        if "Dimension 3" in aff_df.columns:
            total_rows = aff_df[aff_df["Dimension 3"].astype(str) == "Total"]
        else:
            total_rows = aff_df

        if "Dimension 1" in total_rows.columns:
            total_rows = total_rows[
                total_rows["Dimension 1"].astype(str) != "Grand Total"
            ]

        # Filter to Affiliate type (exclude Influencer)
        if "Influencer Filter" in total_rows.columns:
            affiliate_rows = total_rows[
                total_rows["Influencer Filter"].astype(str) == "Affiliate"
            ]
            if affiliate_rows.empty:
                affiliate_rows = total_rows
        else:
            affiliate_rows = total_rows

        # Sort by revenue
        rev_col = None
        for col_name in ["Revenue (Analysis)", "Revenue"]:
            if col_name in affiliate_rows.columns:
                rev_col = col_name
                break

        if rev_col:
            affiliate_rows = affiliate_rows.sort_values(rev_col, ascending=False)

        top_publishers = affiliate_rows.head(10)

        rows = []
        for _, r in top_publishers.iterrows():
            def _get_val(metric, analysis=True):
                suffix = " (Analysis)" if analysis else ""
                col = f"{metric}{suffix}"
                if col in r.index:
                    return _nan_to_none(r[col])
                if metric in r.index:
                    return _nan_to_none(r[metric])
                return None

            def _get_vs_comp(metric):
                col = f"{metric} (vs. Comp)"
                if col in r.index:
                    val = r[col]
                    if val is not None and not (isinstance(val, float) and math.isnan(val)):
                        return _decimal_to_pct(val)
                return None

            revenue = _get_val("Revenue")
            orders = _get_val("Orders")
            sessions = _get_val("Sessions")
            cos_val = _get_val("CoS")
            commission = _get_val("Total Commission")

            row = {
                "publisher_name": r.get("Dimension 2", r.get("Dimension 1", "")),
                "revenue": revenue,
                "revenue_vs_ly": _get_vs_comp("Revenue"),
                "commission": commission,
                "cos": _nan_to_none(_decimal_to_pct(cos_val)) if cos_val is not None else None,
                "orders": _nan_to_none(orders),
                "cvr": _nan_to_none(
                    _decimal_to_pct(_get_val("CVR"))
                ) if _get_val("CVR") is not None else None,
                "sessions": _nan_to_none(sessions),
                "aov": _get_val("AOV"),
            }
            rows.append(row)

        return rows

    # -------------------------------------------------------------------
    # SEO Performance (Slide 10)
    # -------------------------------------------------------------------

    def _transform_seo(self, sources: dict) -> dict:
        """Generate SEO performance KPIs from tracker ORGANIC channel."""
        payload = {
            "seo.title": "SEO Performance",
            "seo.revenue": None,
            "seo.sessions": None,
            "seo.cvr": None,
            "seo.orders": None,
            "seo.aov": None,
            "seo.revenue_vs_ly": None,
            "seo.sessions_vs_ly": None,
            "seo.cvr_vs_ly": None,
            "seo.orders_vs_ly": None,
            "seo.aov_vs_ly": None,
            "seo.narrative": "",
        }

        raw = self._get_raw_data(sources)
        if raw is None or raw.empty:
            return payload

        current = self._filter_month(raw, self.ctx.year, self.ctx.month)
        prior = self._filter_month(raw, self.ctx.prior_year, self.ctx.month)

        org_current = current[current["COS Channel"] == "ORGANIC"]
        org_prior = prior[prior["COS Channel"] == "ORGANIC"]

        if org_current.empty:
            return payload

        cur = self._compute_channel_metrics(org_current)
        pri = (
            self._compute_channel_metrics(org_prior)
            if not org_prior.empty
            else self._empty_channel_metrics()
        )

        payload["seo.revenue"] = cur["revenue"]
        payload["seo.sessions"] = cur["sessions"]
        payload["seo.cvr"] = _nan_to_none(_decimal_to_pct(cur["cvr"]))
        payload["seo.orders"] = cur["orders"]
        payload["seo.aov"] = _nan_to_none(cur["aov"])

        payload["seo.revenue_vs_ly"] = _nan_to_none(
            _safe_pct_change(cur["revenue"], pri["revenue"])
        )
        payload["seo.sessions_vs_ly"] = _nan_to_none(
            _safe_pct_change(cur["sessions"], pri["sessions"])
        )
        payload["seo.cvr_vs_ly"] = _nan_to_none(
            _safe_ppt_change(cur["cvr"], pri["cvr"])
        )
        payload["seo.orders_vs_ly"] = _nan_to_none(
            _safe_pct_change(cur["orders"], pri["orders"])
        )
        payload["seo.aov_vs_ly"] = _nan_to_none(
            _safe_pct_change(cur["aov"], pri["aov"])
        )

        return payload
