"""Monthly eComm Report — canonical 14-slide schema definition.

Defines the complete structure for the No7 US x THGi Monthly eComm Report.
Slide dimensions: 13.333" x 7.5" (standard 16:9).

Data key convention:
    <slide_name>.<field>         — single value
    <slide_name>.<table>_rows    — list of row dicts for a table
    <slide_name>.<chart>_series  — list of series data for a chart
    <slide_name>.<chart>_cats    — category labels for a chart

Positions are approximate reference coordinates from the December 2025
template analysis. The generator should use shape_name matching where
possible and fall back to coordinate matching.
"""

from .models import (
    ChartSeries,
    ChartType,
    DataSlot,
    DesignSystem,
    FontSpec,
    FormatRule,
    FormatType,
    Position,
    SlideSchema,
    SlideType,
    SlotType,
    TableColumn,
    TemplateSchema,
)

# Shared formatting rules
_CURRENCY = FormatRule(FormatType.CURRENCY)
_PERCENTAGE = FormatRule(FormatType.PERCENTAGE)
_VARIANCE = FormatRule(FormatType.VARIANCE_PERCENTAGE)
_POINTS = FormatRule(FormatType.POINTS_CHANGE)
_NUMBER = FormatRule(FormatType.NUMBER)
_INTEGER = FormatRule(FormatType.INTEGER)

# Shared font specs
_KPI_NUMBER = FontSpec(name="DM Sans", size_pt=48.0, bold=True, color="#000000")
_KPI_LABEL = FontSpec(name="DM Sans", size_pt=12.0, bold=False, color="#1C2B33")
_TITLE = FontSpec(name="DM Sans", size_pt=36.0, bold=True, color="#000000")
_HEADER = FontSpec(name="DM Sans", size_pt=24.0, bold=True, color="#000000")
_BODY = FontSpec(name="DM Sans", size_pt=14.0, bold=False, color="#000000")
_TABLE_HEADER = FontSpec(name="DM Sans", size_pt=11.0, bold=True, color="#FFFFFF")
_TABLE_CELL = FontSpec(name="DM Sans", size_pt=11.0, bold=False, color="#000000")
_DIVIDER_TITLE = FontSpec(name="DM Sans", size_pt=36.0, bold=True, color="#FFFFFF")


# ---------------------------------------------------------------------------
# Slide 0: Cover + KPIs
# ---------------------------------------------------------------------------
def _slide_cover() -> SlideSchema:
    return SlideSchema(
        index=0,
        name="cover_kpis",
        title="Cover + KPIs",
        slide_type=SlideType.COVER,
        data_source="tracker:mtd_reporting",
        layout="Title Only",
        slots=[
            DataSlot(
                name="report_title",
                slot_type=SlotType.TEXT,
                data_key="cover.report_title",
                position=Position(left=0.5, top=0.4, width=12.0, height=0.8),
                font=_TITLE,
            ),
            DataSlot(
                name="report_period",
                slot_type=SlotType.TEXT,
                data_key="cover.report_period",
                position=Position(left=0.5, top=1.2, width=8.0, height=0.5),
                font=FontSpec(name="DM Sans", size_pt=20.0, bold=False, color="#1C2B33"),
            ),
            # KPI row — 6 headline metrics across the cover
            DataSlot(
                name="kpi_revenue",
                slot_type=SlotType.KPI_VALUE,
                data_key="cover.total_revenue",
                position=Position(left=0.5, top=3.0, width=2.0, height=1.5),
                font=_KPI_NUMBER,
                format_rule=_CURRENCY,
                label="Revenue",
                variance_key="cover.revenue_vs_target",
            ),
            DataSlot(
                name="kpi_orders",
                slot_type=SlotType.KPI_VALUE,
                data_key="cover.total_orders",
                position=Position(left=2.7, top=3.0, width=2.0, height=1.5),
                font=_KPI_NUMBER,
                format_rule=_NUMBER,
                label="Orders",
                variance_key="cover.orders_vs_target",
            ),
            DataSlot(
                name="kpi_aov",
                slot_type=SlotType.KPI_VALUE,
                data_key="cover.aov",
                position=Position(left=4.9, top=3.0, width=2.0, height=1.5),
                font=_KPI_NUMBER,
                format_rule=_CURRENCY,
                label="AOV",
                variance_key="cover.aov_vs_target",
            ),
            DataSlot(
                name="kpi_new_customers",
                slot_type=SlotType.KPI_VALUE,
                data_key="cover.new_customers",
                position=Position(left=7.1, top=3.0, width=2.0, height=1.5),
                font=_KPI_NUMBER,
                format_rule=_NUMBER,
                label="New Customers",
                variance_key="cover.nc_vs_target",
            ),
            DataSlot(
                name="kpi_cvr",
                slot_type=SlotType.KPI_VALUE,
                data_key="cover.cvr",
                position=Position(left=9.3, top=3.0, width=2.0, height=1.5),
                font=_KPI_NUMBER,
                format_rule=_PERCENTAGE,
                label="CVR",
                variance_key="cover.cvr_vs_target",
            ),
            DataSlot(
                name="kpi_cos",
                slot_type=SlotType.KPI_VALUE,
                data_key="cover.cos",
                position=Position(left=11.5, top=3.0, width=1.5, height=1.5),
                font=_KPI_NUMBER,
                format_rule=_PERCENTAGE,
                label="COS",
                variance_key="cover.cos_vs_target",
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 1: Table of Contents
# ---------------------------------------------------------------------------
def _slide_toc() -> SlideSchema:
    return SlideSchema(
        index=1,
        name="toc",
        title="Table of Contents",
        slide_type=SlideType.TABLE_OF_CONTENTS,
        data_source="static",
        layout="Title Only",
        is_static=True,
        slots=[
            DataSlot(
                name="toc_items",
                slot_type=SlotType.STATIC,
                data_key="toc.items",
                position=Position(left=1.0, top=1.5, width=11.0, height=5.0),
                font=_BODY,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 2: Section Divider — eComm Performance
# ---------------------------------------------------------------------------
def _slide_divider_ecomm() -> SlideSchema:
    return SlideSchema(
        index=2,
        name="divider_ecomm",
        title="eComm Performance",
        slide_type=SlideType.SECTION_DIVIDER,
        data_source="static",
        layout="Title Only",
        is_static=True,
        slots=[
            DataSlot(
                name="section_title",
                slot_type=SlotType.SECTION_DIVIDER,
                data_key="divider.ecomm_title",
                position=Position(left=0.0, top=0.0, width=13.333, height=7.5),
                font=_DIVIDER_TITLE,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 3: Executive Summary / eComm Performance Overview
# ---------------------------------------------------------------------------
def _slide_executive_summary() -> SlideSchema:
    return SlideSchema(
        index=3,
        name="executive_summary",
        title="Executive Summary",
        slide_type=SlideType.DATA,
        data_source="tracker:raw_data,tracker:channel_deep_dive",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="exec.title",
                position=Position(left=0.3, top=0.2, width=12.0, height=0.5),
                font=_HEADER,
            ),
            # Performance summary table (9 rows: Total + 8 channels)
            DataSlot(
                name="performance_table",
                slot_type=SlotType.TABLE,
                data_key="exec.performance_table",
                position=Position(left=0.3, top=0.9, width=12.7, height=4.5),
                row_data_key="exec.performance_rows",
                columns=[
                    TableColumn(header="Channel", data_key="channel", width_inches=1.8,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Revenue", data_key="revenue", width_inches=1.3,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs Target", data_key="revenue_vs_target", width_inches=1.0,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="revenue_vs_ly", width_inches=1.0,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Orders", data_key="orders", width_inches=1.0,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Sessions", data_key="sessions", width_inches=1.1,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="CVR", data_key="cvr", width_inches=0.8,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="AOV", data_key="aov", width_inches=0.9,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="COS", data_key="cos", width_inches=0.8,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="New Customers", data_key="new_customers", width_inches=1.3,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                ],
            ),
            # Narrative summary text
            DataSlot(
                name="narrative",
                slot_type=SlotType.TEXT,
                data_key="exec.narrative",
                position=Position(left=0.3, top=5.6, width=12.7, height=1.5),
                font=_BODY,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 4: Daily Performance
# ---------------------------------------------------------------------------
def _slide_daily_performance() -> SlideSchema:
    return SlideSchema(
        index=4,
        name="daily_performance",
        title="Daily Performance",
        slide_type=SlideType.DATA,
        data_source="tracker:daily",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="daily.title",
                position=Position(left=0.3, top=0.2, width=12.0, height=0.5),
                font=_HEADER,
            ),
            # Daily revenue chart — column clustered with target overlay
            DataSlot(
                name="daily_chart",
                slot_type=SlotType.CHART,
                data_key="daily.chart",
                position=Position(left=0.3, top=0.9, width=8.5, height=4.5),
                chart_type=ChartType.COLUMN_CLUSTERED,
                categories_key="daily.dates",
                series=[
                    ChartSeries(name="Revenue", data_key="daily.revenue_actual",
                                color="#0065E0"),
                    ChartSeries(name="Target", data_key="daily.revenue_target",
                                color="#D1D5DB"),
                    ChartSeries(name="LY", data_key="daily.revenue_ly",
                                color="#1C2B33"),
                ],
            ),
            # Campaign/activity table alongside the chart
            DataSlot(
                name="campaign_table",
                slot_type=SlotType.TABLE,
                data_key="daily.campaign_table",
                position=Position(left=9.0, top=0.9, width=4.0, height=4.5),
                row_data_key="daily.campaign_rows",
                columns=[
                    TableColumn(header="Date", data_key="date", width_inches=1.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Campaign/Activity", data_key="activity", width_inches=3.0,
                                font=_TABLE_HEADER, alignment="left"),
                ],
            ),
            # KPI donuts — revenue vs target gauge
            DataSlot(
                name="revenue_gauge",
                slot_type=SlotType.CHART,
                data_key="daily.revenue_gauge",
                position=Position(left=0.5, top=5.5, width=2.0, height=1.5),
                chart_type=ChartType.DOUGHNUT,
                series=[
                    ChartSeries(name="Achieved", data_key="daily.revenue_achieved_pct",
                                color="#0065E0"),
                    ChartSeries(name="Remaining", data_key="daily.revenue_remaining_pct",
                                color="#D1D5DB"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 5: Promotion Performance
# ---------------------------------------------------------------------------
def _slide_promotion_performance() -> SlideSchema:
    return SlideSchema(
        index=5,
        name="promotion_performance",
        title="Promotion Performance",
        slide_type=SlideType.DATA,
        data_source="offer_performance_csv",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="promo.title",
                position=Position(left=0.3, top=0.2, width=12.0, height=0.5),
                font=_HEADER,
            ),
            DataSlot(
                name="promotion_table",
                slot_type=SlotType.TABLE,
                data_key="promo.table",
                position=Position(left=0.3, top=0.9, width=12.7, height=6.0),
                row_data_key="promo.rows",
                columns=[
                    TableColumn(header="Promotion", data_key="promotion_name", width_inches=4.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Channel", data_key="channel", width_inches=1.5,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Redemptions", data_key="redemptions", width_inches=1.5,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="redemptions_vs_ly", width_inches=1.0,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Revenue", data_key="revenue", width_inches=1.5,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="revenue_vs_ly", width_inches=1.0,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Discount", data_key="discount_amount", width_inches=1.2,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 6: Product Performance
# ---------------------------------------------------------------------------
def _slide_product_performance() -> SlideSchema:
    return SlideSchema(
        index=6,
        name="product_performance",
        title="Product Performance",
        slide_type=SlideType.DATA,
        data_source="product_sales_csv",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="product.title",
                position=Position(left=0.3, top=0.2, width=12.0, height=0.5),
                font=_HEADER,
            ),
            DataSlot(
                name="product_table",
                slot_type=SlotType.TABLE,
                data_key="product.table",
                position=Position(left=0.3, top=0.9, width=12.7, height=6.0),
                row_data_key="product.rows",
                columns=[
                    TableColumn(header="Product", data_key="product_name", width_inches=3.5,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Units", data_key="units", width_inches=1.0,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="units_vs_ly", width_inches=0.9,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Revenue", data_key="revenue", width_inches=1.3,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="revenue_vs_ly", width_inches=0.9,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="AOV", data_key="aov", width_inches=0.9,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="ASP", data_key="avg_selling_price", width_inches=0.9,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Discount %", data_key="discount_pct", width_inches=1.0,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="New Cust", data_key="new_customers", width_inches=1.0,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 7: Section Divider — Channel Deep Dives
# ---------------------------------------------------------------------------
def _slide_divider_channels() -> SlideSchema:
    return SlideSchema(
        index=7,
        name="divider_channels",
        title="Channel Deep Dives",
        slide_type=SlideType.SECTION_DIVIDER,
        data_source="static",
        layout="Title Only",
        is_static=True,
        slots=[
            DataSlot(
                name="section_title",
                slot_type=SlotType.SECTION_DIVIDER,
                data_key="divider.channels_title",
                position=Position(left=0.0, top=0.0, width=13.333, height=7.5),
                font=_DIVIDER_TITLE,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 8: CRM Performance
# ---------------------------------------------------------------------------
def _slide_crm() -> SlideSchema:
    return SlideSchema(
        index=8,
        name="crm_performance",
        title="CRM Performance",
        slide_type=SlideType.DATA,
        data_source="crm_excel,tracker:email",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="crm.title",
                position=Position(left=0.3, top=0.2, width=12.0, height=0.5),
                font=_HEADER,
            ),
            # CRM KPI row
            DataSlot(
                name="kpi_emails_sent",
                slot_type=SlotType.KPI_VALUE,
                data_key="crm.emails_sent",
                position=Position(left=0.5, top=1.0, width=2.0, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_INTEGER,
                label="Emails Sent",
                variance_key="crm.emails_sent_vs_ly",
            ),
            DataSlot(
                name="kpi_open_rate",
                slot_type=SlotType.KPI_VALUE,
                data_key="crm.open_rate",
                position=Position(left=2.7, top=1.0, width=2.0, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_PERCENTAGE,
                label="Open Rate",
                variance_key="crm.open_rate_vs_ly",
            ),
            DataSlot(
                name="kpi_ctr",
                slot_type=SlotType.KPI_VALUE,
                data_key="crm.ctr",
                position=Position(left=4.9, top=1.0, width=2.0, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_PERCENTAGE,
                label="CTR",
                variance_key="crm.ctr_vs_ly",
            ),
            DataSlot(
                name="kpi_crm_revenue",
                slot_type=SlotType.KPI_VALUE,
                data_key="crm.revenue",
                position=Position(left=7.1, top=1.0, width=2.0, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_CURRENCY,
                label="Revenue",
                variance_key="crm.revenue_vs_ly",
            ),
            DataSlot(
                name="kpi_crm_cvr",
                slot_type=SlotType.KPI_VALUE,
                data_key="crm.cvr",
                position=Position(left=9.3, top=1.0, width=2.0, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_PERCENTAGE,
                label="CVR",
                variance_key="crm.cvr_vs_ly",
            ),
            DataSlot(
                name="kpi_crm_aov",
                slot_type=SlotType.KPI_VALUE,
                data_key="crm.aov",
                position=Position(left=11.5, top=1.0, width=1.5, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_CURRENCY,
                label="AOV",
                variance_key="crm.aov_vs_ly",
            ),
            # CRM detail table
            DataSlot(
                name="crm_detail_table",
                slot_type=SlotType.TABLE,
                data_key="crm.detail_table",
                position=Position(left=0.3, top=2.8, width=12.7, height=4.2),
                row_data_key="crm.detail_rows",
                columns=[
                    TableColumn(header="Campaign Type", data_key="campaign_type", width_inches=2.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Emails Sent", data_key="emails_sent", width_inches=1.3,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Open Rate", data_key="open_rate", width_inches=1.1,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="CTR", data_key="ctr", width_inches=1.0,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Sessions", data_key="sessions", width_inches=1.1,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Orders", data_key="orders", width_inches=1.0,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="CVR", data_key="cvr", width_inches=0.8,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Revenue", data_key="revenue", width_inches=1.3,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="AOV", data_key="aov", width_inches=1.0,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="revenue_vs_ly", width_inches=1.1,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 9: Affiliate Performance
# ---------------------------------------------------------------------------
def _slide_affiliate() -> SlideSchema:
    return SlideSchema(
        index=9,
        name="affiliate_performance",
        title="Affiliate Performance",
        slide_type=SlideType.DATA,
        data_source="affiliate_excel,tracker:affiliate",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="affiliate.title",
                position=Position(left=0.3, top=0.2, width=12.0, height=0.5),
                font=_HEADER,
            ),
            # Affiliate KPI row
            DataSlot(
                name="kpi_aff_revenue",
                slot_type=SlotType.KPI_VALUE,
                data_key="affiliate.revenue",
                position=Position(left=0.5, top=1.0, width=2.5, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_CURRENCY,
                label="Revenue",
                variance_key="affiliate.revenue_vs_ly",
            ),
            DataSlot(
                name="kpi_aff_cos",
                slot_type=SlotType.KPI_VALUE,
                data_key="affiliate.cos",
                position=Position(left=3.2, top=1.0, width=2.0, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_PERCENTAGE,
                label="COS",
                variance_key="affiliate.cos_vs_ly",
            ),
            DataSlot(
                name="kpi_aff_roas",
                slot_type=SlotType.KPI_VALUE,
                data_key="affiliate.roas",
                position=Position(left=5.4, top=1.0, width=2.0, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_NUMBER,
                label="ROAS",
                variance_key="affiliate.roas_vs_ly",
            ),
            DataSlot(
                name="kpi_aff_orders",
                slot_type=SlotType.KPI_VALUE,
                data_key="affiliate.orders",
                position=Position(left=7.6, top=1.0, width=2.0, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_INTEGER,
                label="Orders",
                variance_key="affiliate.orders_vs_ly",
            ),
            DataSlot(
                name="kpi_aff_cvr",
                slot_type=SlotType.KPI_VALUE,
                data_key="affiliate.cvr",
                position=Position(left=9.8, top=1.0, width=2.0, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_PERCENTAGE,
                label="CVR",
                variance_key="affiliate.cvr_vs_ly",
            ),
            # Top publishers table
            DataSlot(
                name="publisher_table",
                slot_type=SlotType.TABLE,
                data_key="affiliate.publisher_table",
                position=Position(left=0.3, top=2.8, width=12.7, height=4.2),
                row_data_key="affiliate.publisher_rows",
                columns=[
                    TableColumn(header="Publisher", data_key="publisher_name", width_inches=3.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Revenue", data_key="revenue", width_inches=1.5,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="revenue_vs_ly", width_inches=1.0,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Commission", data_key="commission", width_inches=1.3,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="COS", data_key="cos", width_inches=0.8,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Orders", data_key="orders", width_inches=1.0,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="CVR", data_key="cvr", width_inches=0.8,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Sessions", data_key="sessions", width_inches=1.1,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="AOV", data_key="aov", width_inches=1.0,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 10: SEO Performance
# ---------------------------------------------------------------------------
def _slide_seo() -> SlideSchema:
    return SlideSchema(
        index=10,
        name="seo_performance",
        title="SEO Performance",
        slide_type=SlideType.DATA,
        data_source="tracker:organic",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="seo.title",
                position=Position(left=0.3, top=0.2, width=12.0, height=0.5),
                font=_HEADER,
            ),
            # SEO KPI row
            DataSlot(
                name="kpi_seo_revenue",
                slot_type=SlotType.KPI_VALUE,
                data_key="seo.revenue",
                position=Position(left=0.5, top=1.0, width=2.5, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_CURRENCY,
                label="Revenue",
                variance_key="seo.revenue_vs_ly",
            ),
            DataSlot(
                name="kpi_seo_sessions",
                slot_type=SlotType.KPI_VALUE,
                data_key="seo.sessions",
                position=Position(left=3.2, top=1.0, width=2.0, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_INTEGER,
                label="Sessions",
                variance_key="seo.sessions_vs_ly",
            ),
            DataSlot(
                name="kpi_seo_cvr",
                slot_type=SlotType.KPI_VALUE,
                data_key="seo.cvr",
                position=Position(left=5.4, top=1.0, width=2.0, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_PERCENTAGE,
                label="CVR",
                variance_key="seo.cvr_vs_ly",
            ),
            DataSlot(
                name="kpi_seo_orders",
                slot_type=SlotType.KPI_VALUE,
                data_key="seo.orders",
                position=Position(left=7.6, top=1.0, width=2.0, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_INTEGER,
                label="Orders",
                variance_key="seo.orders_vs_ly",
            ),
            DataSlot(
                name="kpi_seo_aov",
                slot_type=SlotType.KPI_VALUE,
                data_key="seo.aov",
                position=Position(left=9.8, top=1.0, width=2.0, height=1.2),
                font=FontSpec(name="DM Sans", size_pt=30.0, bold=True),
                format_rule=_CURRENCY,
                label="AOV",
                variance_key="seo.aov_vs_ly",
            ),
            # Narrative / performance summary
            DataSlot(
                name="seo_narrative",
                slot_type=SlotType.TEXT,
                data_key="seo.narrative",
                position=Position(left=0.3, top=2.8, width=12.7, height=4.2),
                font=_BODY,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 11: Section Divider — Outlook
# ---------------------------------------------------------------------------
def _slide_divider_outlook() -> SlideSchema:
    return SlideSchema(
        index=11,
        name="divider_outlook",
        title="Outlook",
        slide_type=SlideType.SECTION_DIVIDER,
        data_source="static",
        layout="Title Only",
        is_static=True,
        slots=[
            DataSlot(
                name="section_title",
                slot_type=SlotType.SECTION_DIVIDER,
                data_key="divider.outlook_title",
                position=Position(left=0.0, top=0.0, width=13.333, height=7.5),
                font=_DIVIDER_TITLE,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 12: Upcoming Promotions
# ---------------------------------------------------------------------------
def _slide_upcoming_promos() -> SlideSchema:
    return SlideSchema(
        index=12,
        name="upcoming_promotions",
        title="Upcoming Promotions",
        slide_type=SlideType.MANUAL,
        data_source="manual",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="upcoming.title",
                position=Position(left=0.3, top=0.2, width=12.0, height=0.5),
                font=_HEADER,
            ),
            DataSlot(
                name="promo_calendar_table",
                slot_type=SlotType.TABLE,
                data_key="upcoming.table",
                position=Position(left=0.3, top=0.9, width=12.7, height=6.0),
                row_data_key="upcoming.rows",
                columns=[
                    TableColumn(header="Date", data_key="date", width_inches=2.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Promotion", data_key="promotion", width_inches=4.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Discount", data_key="discount", width_inches=2.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Channels", data_key="channels", width_inches=3.0,
                                font=_TABLE_HEADER, alignment="left"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 13: Next Steps
# ---------------------------------------------------------------------------
def _slide_next_steps() -> SlideSchema:
    return SlideSchema(
        index=13,
        name="next_steps",
        title="Next Steps",
        slide_type=SlideType.MANUAL,
        data_source="manual",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="next_steps.title",
                position=Position(left=0.3, top=0.2, width=12.0, height=0.5),
                font=_HEADER,
            ),
            DataSlot(
                name="action_items",
                slot_type=SlotType.TEXT,
                data_key="next_steps.items",
                position=Position(left=0.3, top=0.9, width=12.7, height=6.0),
                font=_BODY,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Assembled schema
# ---------------------------------------------------------------------------

def build_monthly_report_schema() -> TemplateSchema:
    """Build and return the complete 14-slide monthly eComm report schema."""
    return TemplateSchema(
        name="No7 US Monthly eComm Report",
        report_type="monthly",
        width_inches=13.333,
        height_inches=7.5,
        naming_convention="No7 US x THGi Monthly eComm Report - {month} {year} Overview.pptx",
        design=DesignSystem(),
        slides=[
            _slide_cover(),          # 0
            _slide_toc(),            # 1
            _slide_divider_ecomm(),  # 2
            _slide_executive_summary(),  # 3
            _slide_daily_performance(),  # 4
            _slide_promotion_performance(),  # 5
            _slide_product_performance(),    # 6
            _slide_divider_channels(),       # 7
            _slide_crm(),                    # 8
            _slide_affiliate(),              # 9
            _slide_seo(),                    # 10
            _slide_divider_outlook(),        # 11
            _slide_upcoming_promos(),        # 12
            _slide_next_steps(),             # 13
        ],
    )
