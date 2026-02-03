"""Quarterly Business Review (QBR) — 29-slide schema definition.

Defines the complete structure for the No7 US x THGi Quarterly Business
Review presentation.  Slide dimensions: 21.986" x 12.368" (oversized 16:9).

The QBR aggregates three months of data and adds strategy, operational,
and forward-looking sections that the monthly report does not include.

Data key convention (same as monthly):
    <slide_name>.<field>         — single value
    <slide_name>.<table>_rows    — list of row dicts for a table
    <slide_name>.<chart>_series  — list of series data for a chart
    <slide_name>.<chart>_cats    — category labels for a chart

Quarterly keys use the prefix of the slide name (e.g. ``qcover``,
``qexec``, ``qrevenue``).  Channel deep-dives use the ``q`` prefix
to disambiguate from the monthly namespace (``qcrm``, ``qaff``, etc.).

Positions are approximate reference coordinates derived from the
qbr-template.pptx analysis.  The generator should use shape_name
matching where possible and fall back to coordinate matching.
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

# -- slide dimensions -------------------------------------------------------
_W = 21.986
_H = 12.368

# -- shared formatting rules -------------------------------------------------
_CURRENCY = FormatRule(FormatType.CURRENCY)
_PERCENTAGE = FormatRule(FormatType.PERCENTAGE)
_VARIANCE = FormatRule(FormatType.VARIANCE_PERCENTAGE)
_POINTS = FormatRule(FormatType.POINTS_CHANGE)
_NUMBER = FormatRule(FormatType.NUMBER)
_INTEGER = FormatRule(FormatType.INTEGER)

# -- shared font specs -------------------------------------------------------
_KPI_NUMBER = FontSpec(name="DM Sans", size_pt=60.0, bold=True, color="#000000")
_KPI_LABEL = FontSpec(name="DM Sans", size_pt=14.0, bold=False, color="#1C2B33")
_TITLE = FontSpec(name="DM Sans", size_pt=44.0, bold=True, color="#000000")
_HEADER = FontSpec(name="DM Sans", size_pt=30.0, bold=True, color="#000000")
_SUBHEADER = FontSpec(name="DM Sans", size_pt=24.0, bold=True, color="#000000")
_BODY = FontSpec(name="DM Sans", size_pt=16.0, bold=False, color="#000000")
_TABLE_HEADER = FontSpec(name="DM Sans", size_pt=14.0, bold=True, color="#FFFFFF")
_TABLE_CELL = FontSpec(name="DM Sans", size_pt=14.0, bold=False, color="#000000")
_DIVIDER_TITLE = FontSpec(name="DM Sans", size_pt=60.0, bold=True, color="#FFFFFF")
_CHART_LABEL = FontSpec(name="DM Sans", size_pt=11.0, bold=False, color="#1C2B33")


# ---------------------------------------------------------------------------
# Slide 0: Cover + Quarter KPIs
# ---------------------------------------------------------------------------
def _slide_cover() -> SlideSchema:
    return SlideSchema(
        index=0,
        name="qbr_cover",
        title="Cover + Quarter KPIs",
        slide_type=SlideType.COVER,
        data_source="tracker:raw_data",
        layout="Title Only",
        slots=[
            DataSlot(
                name="report_title",
                slot_type=SlotType.TEXT,
                data_key="qcover.report_title",
                position=Position(left=1.0, top=0.8, width=20.0, height=1.2),
                font=_TITLE,
            ),
            DataSlot(
                name="report_period",
                slot_type=SlotType.TEXT,
                data_key="qcover.report_period",
                position=Position(left=1.0, top=2.2, width=14.0, height=0.8),
                font=FontSpec(name="DM Sans", size_pt=24.0, bold=False, color="#1C2B33"),
            ),
            # Quarter KPI row — 6 headline metrics
            DataSlot(
                name="kpi_revenue",
                slot_type=SlotType.KPI_VALUE,
                data_key="qcover.total_revenue",
                position=Position(left=1.0, top=5.0, width=3.2, height=2.5),
                font=_KPI_NUMBER,
                format_rule=_CURRENCY,
                label="Revenue",
                variance_key="qcover.revenue_vs_target",
            ),
            DataSlot(
                name="kpi_orders",
                slot_type=SlotType.KPI_VALUE,
                data_key="qcover.total_orders",
                position=Position(left=4.5, top=5.0, width=3.2, height=2.5),
                font=_KPI_NUMBER,
                format_rule=_NUMBER,
                label="Orders",
                variance_key="qcover.orders_vs_target",
            ),
            DataSlot(
                name="kpi_aov",
                slot_type=SlotType.KPI_VALUE,
                data_key="qcover.aov",
                position=Position(left=8.0, top=5.0, width=3.2, height=2.5),
                font=_KPI_NUMBER,
                format_rule=_CURRENCY,
                label="AOV",
                variance_key="qcover.aov_vs_target",
            ),
            DataSlot(
                name="kpi_new_customers",
                slot_type=SlotType.KPI_VALUE,
                data_key="qcover.new_customers",
                position=Position(left=11.5, top=5.0, width=3.2, height=2.5),
                font=_KPI_NUMBER,
                format_rule=_NUMBER,
                label="New Customers",
                variance_key="qcover.nc_vs_target",
            ),
            DataSlot(
                name="kpi_cvr",
                slot_type=SlotType.KPI_VALUE,
                data_key="qcover.cvr",
                position=Position(left=15.0, top=5.0, width=3.2, height=2.5),
                font=_KPI_NUMBER,
                format_rule=_PERCENTAGE,
                label="CVR",
                variance_key="qcover.cvr_vs_target",
            ),
            DataSlot(
                name="kpi_cos",
                slot_type=SlotType.KPI_VALUE,
                data_key="qcover.cos",
                position=Position(left=18.5, top=5.0, width=2.5, height=2.5),
                font=_KPI_NUMBER,
                format_rule=_PERCENTAGE,
                label="COS",
                variance_key="qcover.cos_vs_target",
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 1: Agenda / Table of Contents
# ---------------------------------------------------------------------------
def _slide_agenda() -> SlideSchema:
    return SlideSchema(
        index=1,
        name="qbr_agenda",
        title="Agenda",
        slide_type=SlideType.TABLE_OF_CONTENTS,
        data_source="static",
        layout="Title Only",
        is_static=True,
        slots=[
            DataSlot(
                name="agenda_items",
                slot_type=SlotType.TABLE,
                data_key="qagenda.table",
                position=Position(left=1.5, top=1.5, width=19.0, height=9.5),
                row_data_key="qagenda.rows",
                columns=[
                    TableColumn(header="#", data_key="number", width_inches=1.0,
                                font=_TABLE_HEADER, alignment="center"),
                    TableColumn(header="Section", data_key="section", width_inches=8.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Slide", data_key="slide_ref", width_inches=2.0,
                                font=_TABLE_HEADER, alignment="center"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 2: Executive Summary — Quarter Performance Table
# ---------------------------------------------------------------------------
def _slide_executive_summary() -> SlideSchema:
    return SlideSchema(
        index=2,
        name="qbr_executive_summary",
        title="Executive Summary",
        slide_type=SlideType.DATA,
        data_source="tracker:raw_data",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qexec.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            # Quarter performance summary table (channels x metrics)
            DataSlot(
                name="performance_table",
                slot_type=SlotType.TABLE,
                data_key="qexec.performance_table",
                position=Position(left=0.5, top=1.5, width=21.0, height=6.5),
                row_data_key="qexec.performance_rows",
                columns=[
                    TableColumn(header="Channel", data_key="channel", width_inches=2.5,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Revenue", data_key="revenue", width_inches=1.8,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs Target", data_key="revenue_vs_target", width_inches=1.3,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="revenue_vs_ly", width_inches=1.3,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Orders", data_key="orders", width_inches=1.3,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Sessions", data_key="sessions", width_inches=1.5,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="CVR", data_key="cvr", width_inches=1.0,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="AOV", data_key="aov", width_inches=1.2,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="COS", data_key="cos", width_inches=1.0,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="New Customers", data_key="new_customers", width_inches=1.8,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Contribution", data_key="contribution_pct", width_inches=1.3,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                ],
            ),
            # Three-box narrative summary
            DataSlot(
                name="theme_1",
                slot_type=SlotType.TEXT,
                data_key="qexec.theme_1",
                position=Position(left=0.5, top=8.5, width=6.5, height=3.0),
                font=_BODY,
            ),
            DataSlot(
                name="theme_2",
                slot_type=SlotType.TEXT,
                data_key="qexec.theme_2",
                position=Position(left=7.5, top=8.5, width=6.5, height=3.0),
                font=_BODY,
            ),
            DataSlot(
                name="theme_3",
                slot_type=SlotType.TEXT,
                data_key="qexec.theme_3",
                position=Position(left=14.5, top=8.5, width=6.5, height=3.0),
                font=_BODY,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 3: Section Divider — Strategy Review
# ---------------------------------------------------------------------------
def _slide_divider_strategy() -> SlideSchema:
    return SlideSchema(
        index=3,
        name="divider_strategy",
        title="Strategy Review",
        slide_type=SlideType.SECTION_DIVIDER,
        data_source="static",
        layout="Title Only",
        is_static=True,
        slots=[
            DataSlot(
                name="section_title",
                slot_type=SlotType.SECTION_DIVIDER,
                data_key="qdivider.strategy_title",
                position=Position(left=0.0, top=0.0, width=_W, height=_H),
                font=_DIVIDER_TITLE,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 4: Strategy Review — 4-Pillar Grid
# ---------------------------------------------------------------------------
def _slide_strategy_review() -> SlideSchema:
    return SlideSchema(
        index=4,
        name="qbr_strategy_review",
        title="Strategy Review",
        slide_type=SlideType.MANUAL,
        data_source="manual",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qstrategy.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            # Four strategic pillars arranged in 2x2 grid
            DataSlot(
                name="pillar_1",
                slot_type=SlotType.TEXT,
                data_key="qstrategy.pillar_1",
                position=Position(left=0.5, top=1.5, width=10.0, height=4.5),
                font=_BODY,
            ),
            DataSlot(
                name="pillar_2",
                slot_type=SlotType.TEXT,
                data_key="qstrategy.pillar_2",
                position=Position(left=11.0, top=1.5, width=10.0, height=4.5),
                font=_BODY,
            ),
            DataSlot(
                name="pillar_3",
                slot_type=SlotType.TEXT,
                data_key="qstrategy.pillar_3",
                position=Position(left=0.5, top=6.5, width=10.0, height=4.5),
                font=_BODY,
            ),
            DataSlot(
                name="pillar_4",
                slot_type=SlotType.TEXT,
                data_key="qstrategy.pillar_4",
                position=Position(left=11.0, top=6.5, width=10.0, height=4.5),
                font=_BODY,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 5: Quarter Successes
# ---------------------------------------------------------------------------
def _slide_successes() -> SlideSchema:
    return SlideSchema(
        index=5,
        name="qbr_successes",
        title="Quarter Successes",
        slide_type=SlideType.MANUAL,
        data_source="manual",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qsuccesses.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            DataSlot(
                name="successes_narrative",
                slot_type=SlotType.TEXT,
                data_key="qsuccesses.narrative",
                position=Position(left=0.5, top=1.5, width=20.0, height=9.5),
                font=_BODY,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 6: Quarter Challenges
# ---------------------------------------------------------------------------
def _slide_challenges() -> SlideSchema:
    return SlideSchema(
        index=6,
        name="qbr_challenges",
        title="Quarter Challenges",
        slide_type=SlideType.MANUAL,
        data_source="manual",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qchallenges.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            DataSlot(
                name="challenges_narrative",
                slot_type=SlotType.TEXT,
                data_key="qchallenges.narrative",
                position=Position(left=0.5, top=1.5, width=20.0, height=9.5),
                font=_BODY,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 7: Revenue Performance — Monthly Chart (3 months)
# ---------------------------------------------------------------------------
def _slide_revenue_chart() -> SlideSchema:
    return SlideSchema(
        index=7,
        name="qbr_revenue_chart",
        title="Revenue Performance",
        slide_type=SlideType.DATA,
        data_source="tracker:raw_data",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qrevenue.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            # Monthly revenue bars — 3 months of the quarter
            DataSlot(
                name="revenue_chart",
                slot_type=SlotType.CHART,
                data_key="qrevenue.chart",
                position=Position(left=0.5, top=1.5, width=14.0, height=7.0),
                chart_type=ChartType.COLUMN_CLUSTERED,
                categories_key="qrevenue.months",
                series=[
                    ChartSeries(name="Revenue TY", data_key="qrevenue.revenue_ty",
                                color="#0065E0"),
                    ChartSeries(name="Revenue LY", data_key="qrevenue.revenue_ly",
                                color="#1C2B33"),
                    ChartSeries(name="Target", data_key="qrevenue.revenue_target",
                                color="#D1D5DB"),
                ],
            ),
            # Quarter-level KPI gauges alongside the chart
            DataSlot(
                name="revenue_gauge",
                slot_type=SlotType.CHART,
                data_key="qrevenue.revenue_gauge",
                position=Position(left=15.5, top=1.5, width=3.0, height=3.0),
                chart_type=ChartType.DOUGHNUT,
                series=[
                    ChartSeries(name="Achieved", data_key="qrevenue.achieved_pct",
                                color="#0065E0"),
                    ChartSeries(name="Remaining", data_key="qrevenue.remaining_pct",
                                color="#D1D5DB"),
                ],
            ),
            DataSlot(
                name="cos_gauge",
                slot_type=SlotType.CHART,
                data_key="qrevenue.cos_gauge",
                position=Position(left=15.5, top=5.0, width=3.0, height=3.0),
                chart_type=ChartType.DOUGHNUT,
                series=[
                    ChartSeries(name="COS", data_key="qrevenue.cos_actual_pct",
                                color="#0065E0"),
                    ChartSeries(name="Remaining", data_key="qrevenue.cos_remaining_pct",
                                color="#D1D5DB"),
                ],
            ),
            # Monthly breakdown table below chart
            DataSlot(
                name="monthly_breakdown",
                slot_type=SlotType.TABLE,
                data_key="qrevenue.monthly_table",
                position=Position(left=0.5, top=9.0, width=20.0, height=2.8),
                row_data_key="qrevenue.monthly_rows",
                columns=[
                    TableColumn(header="Month", data_key="month", width_inches=3.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Revenue", data_key="revenue", width_inches=2.5,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs Target", data_key="revenue_vs_target", width_inches=2.0,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="revenue_vs_ly", width_inches=2.0,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Orders", data_key="orders", width_inches=2.0,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="AOV", data_key="aov", width_inches=1.5,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="CVR", data_key="cvr", width_inches=1.5,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="COS", data_key="cos", width_inches=1.5,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 8: KPI Overview — Doughnut Gauges
# ---------------------------------------------------------------------------
def _slide_kpi_overview() -> SlideSchema:
    return SlideSchema(
        index=8,
        name="qbr_kpi_overview",
        title="KPI Overview",
        slide_type=SlideType.DATA,
        data_source="tracker:raw_data",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qkpi.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            # Top row: Revenue, AOV, CVR gauges
            DataSlot(
                name="revenue_gauge",
                slot_type=SlotType.CHART,
                data_key="qkpi.revenue_gauge",
                position=Position(left=1.0, top=1.5, width=6.0, height=4.5),
                chart_type=ChartType.DOUGHNUT,
                series=[
                    ChartSeries(name="Achieved", data_key="qkpi.revenue_achieved",
                                color="#0065E0"),
                    ChartSeries(name="Remaining", data_key="qkpi.revenue_remaining",
                                color="#D1D5DB"),
                ],
            ),
            DataSlot(
                name="aov_gauge",
                slot_type=SlotType.CHART,
                data_key="qkpi.aov_gauge",
                position=Position(left=8.0, top=1.5, width=6.0, height=4.5),
                chart_type=ChartType.DOUGHNUT,
                series=[
                    ChartSeries(name="Achieved", data_key="qkpi.aov_achieved",
                                color="#0065E0"),
                    ChartSeries(name="Remaining", data_key="qkpi.aov_remaining",
                                color="#D1D5DB"),
                ],
            ),
            DataSlot(
                name="cvr_gauge",
                slot_type=SlotType.CHART,
                data_key="qkpi.cvr_gauge",
                position=Position(left=15.0, top=1.5, width=6.0, height=4.5),
                chart_type=ChartType.DOUGHNUT,
                series=[
                    ChartSeries(name="Achieved", data_key="qkpi.cvr_achieved",
                                color="#0065E0"),
                    ChartSeries(name="Remaining", data_key="qkpi.cvr_remaining",
                                color="#D1D5DB"),
                ],
            ),
            # Bottom row: COS, NC, Orders gauges
            DataSlot(
                name="cos_gauge",
                slot_type=SlotType.CHART,
                data_key="qkpi.cos_gauge",
                position=Position(left=1.0, top=7.0, width=6.0, height=4.5),
                chart_type=ChartType.DOUGHNUT,
                series=[
                    ChartSeries(name="Achieved", data_key="qkpi.cos_achieved",
                                color="#0065E0"),
                    ChartSeries(name="Remaining", data_key="qkpi.cos_remaining",
                                color="#D1D5DB"),
                ],
            ),
            DataSlot(
                name="nc_gauge",
                slot_type=SlotType.CHART,
                data_key="qkpi.nc_gauge",
                position=Position(left=8.0, top=7.0, width=6.0, height=4.5),
                chart_type=ChartType.DOUGHNUT,
                series=[
                    ChartSeries(name="Achieved", data_key="qkpi.nc_achieved",
                                color="#0065E0"),
                    ChartSeries(name="Remaining", data_key="qkpi.nc_remaining",
                                color="#D1D5DB"),
                ],
            ),
            DataSlot(
                name="orders_gauge",
                slot_type=SlotType.CHART,
                data_key="qkpi.orders_gauge",
                position=Position(left=15.0, top=7.0, width=6.0, height=4.5),
                chart_type=ChartType.DOUGHNUT,
                series=[
                    ChartSeries(name="Achieved", data_key="qkpi.orders_achieved",
                                color="#0065E0"),
                    ChartSeries(name="Remaining", data_key="qkpi.orders_remaining",
                                color="#D1D5DB"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 9: Section Divider — Channel Performance
# ---------------------------------------------------------------------------
def _slide_divider_channels() -> SlideSchema:
    return SlideSchema(
        index=9,
        name="divider_channels",
        title="Channel Performance",
        slide_type=SlideType.SECTION_DIVIDER,
        data_source="static",
        layout="Title Only",
        is_static=True,
        slots=[
            DataSlot(
                name="section_title",
                slot_type=SlotType.SECTION_DIVIDER,
                data_key="qdivider.channels_title",
                position=Position(left=0.0, top=0.0, width=_W, height=_H),
                font=_DIVIDER_TITLE,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 10: Channel Mix Overview
# ---------------------------------------------------------------------------
def _slide_channel_mix() -> SlideSchema:
    return SlideSchema(
        index=10,
        name="qbr_channel_mix",
        title="Channel Mix",
        slide_type=SlideType.DATA,
        data_source="tracker:raw_data",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qchannel_mix.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            # Revenue by channel — stacked/clustered chart
            DataSlot(
                name="channel_chart",
                slot_type=SlotType.CHART,
                data_key="qchannel_mix.chart",
                position=Position(left=0.5, top=1.5, width=12.0, height=9.0),
                chart_type=ChartType.COLUMN_CLUSTERED,
                categories_key="qchannel_mix.channels",
                series=[
                    ChartSeries(name="Revenue TY", data_key="qchannel_mix.revenue_ty",
                                color="#0065E0"),
                    ChartSeries(name="Revenue LY", data_key="qchannel_mix.revenue_ly",
                                color="#1C2B33"),
                ],
            ),
            # Contribution table
            DataSlot(
                name="contribution_table",
                slot_type=SlotType.TABLE,
                data_key="qchannel_mix.contribution_table",
                position=Position(left=13.5, top=1.5, width=8.0, height=9.0),
                row_data_key="qchannel_mix.contribution_rows",
                columns=[
                    TableColumn(header="Channel", data_key="channel", width_inches=2.5,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Revenue", data_key="revenue", width_inches=2.0,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Mix %", data_key="contribution_pct", width_inches=1.5,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="revenue_vs_ly", width_inches=1.5,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 11: CRM Deep Dive (Two-column: Q Recap | Q+1 Strategy)
# ---------------------------------------------------------------------------
def _slide_crm() -> SlideSchema:
    return SlideSchema(
        index=11,
        name="qbr_crm",
        title="CRM Performance",
        slide_type=SlideType.DATA,
        data_source="crm_excel,tracker:email",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qcrm.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            # CRM headline KPIs
            DataSlot(
                name="kpi_emails_sent",
                slot_type=SlotType.KPI_VALUE,
                data_key="qcrm.emails_sent",
                position=Position(left=0.5, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_INTEGER,
                label="Emails Sent",
                variance_key="qcrm.emails_sent_vs_ly",
            ),
            DataSlot(
                name="kpi_open_rate",
                slot_type=SlotType.KPI_VALUE,
                data_key="qcrm.open_rate",
                position=Position(left=4.5, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_PERCENTAGE,
                label="Open Rate",
                variance_key="qcrm.open_rate_vs_ly",
            ),
            DataSlot(
                name="kpi_ctr",
                slot_type=SlotType.KPI_VALUE,
                data_key="qcrm.ctr",
                position=Position(left=8.5, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_PERCENTAGE,
                label="CTR",
                variance_key="qcrm.ctr_vs_ly",
            ),
            DataSlot(
                name="kpi_revenue",
                slot_type=SlotType.KPI_VALUE,
                data_key="qcrm.revenue",
                position=Position(left=12.5, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_CURRENCY,
                label="Revenue",
                variance_key="qcrm.revenue_vs_ly",
            ),
            DataSlot(
                name="kpi_cvr",
                slot_type=SlotType.KPI_VALUE,
                data_key="qcrm.cvr",
                position=Position(left=16.5, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_PERCENTAGE,
                label="CVR",
                variance_key="qcrm.cvr_vs_ly",
            ),
            # CRM monthly trend chart
            DataSlot(
                name="crm_chart",
                slot_type=SlotType.CHART,
                data_key="qcrm.chart",
                position=Position(left=0.5, top=3.5, width=10.0, height=5.0),
                chart_type=ChartType.COLUMN_CLUSTERED,
                categories_key="qcrm.months",
                series=[
                    ChartSeries(name="Revenue TY", data_key="qcrm.revenue_monthly_ty",
                                color="#0065E0"),
                    ChartSeries(name="Revenue LY", data_key="qcrm.revenue_monthly_ly",
                                color="#1C2B33"),
                ],
            ),
            # CRM detail table — by campaign type
            DataSlot(
                name="detail_table",
                slot_type=SlotType.TABLE,
                data_key="qcrm.detail_table",
                position=Position(left=11.0, top=3.5, width=10.0, height=5.0),
                row_data_key="qcrm.detail_rows",
                columns=[
                    TableColumn(header="Campaign Type", data_key="campaign_type", width_inches=2.5,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Emails Sent", data_key="emails_sent", width_inches=1.5,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Open Rate", data_key="open_rate", width_inches=1.3,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="CTR", data_key="ctr", width_inches=1.0,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Revenue", data_key="revenue", width_inches=1.5,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="revenue_vs_ly", width_inches=1.2,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                ],
            ),
            # Q+1 strategy sidebar
            DataSlot(
                name="next_quarter_strategy",
                slot_type=SlotType.TEXT,
                data_key="qcrm.next_quarter_strategy",
                position=Position(left=0.5, top=9.0, width=20.5, height=2.8),
                font=_BODY,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 12: Affiliate Deep Dive
# ---------------------------------------------------------------------------
def _slide_affiliate() -> SlideSchema:
    return SlideSchema(
        index=12,
        name="qbr_affiliate",
        title="Affiliate Performance",
        slide_type=SlideType.DATA,
        data_source="affiliate_excel,tracker:affiliate",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qaff.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            # Headline KPIs
            DataSlot(
                name="kpi_revenue",
                slot_type=SlotType.KPI_VALUE,
                data_key="qaff.revenue",
                position=Position(left=0.5, top=1.5, width=4.0, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_CURRENCY,
                label="Revenue",
                variance_key="qaff.revenue_vs_ly",
            ),
            DataSlot(
                name="kpi_cos",
                slot_type=SlotType.KPI_VALUE,
                data_key="qaff.cos",
                position=Position(left=5.0, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_PERCENTAGE,
                label="COS",
                variance_key="qaff.cos_vs_ly",
            ),
            DataSlot(
                name="kpi_roas",
                slot_type=SlotType.KPI_VALUE,
                data_key="qaff.roas",
                position=Position(left=9.0, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_NUMBER,
                label="ROAS",
                variance_key="qaff.roas_vs_ly",
            ),
            DataSlot(
                name="kpi_orders",
                slot_type=SlotType.KPI_VALUE,
                data_key="qaff.orders",
                position=Position(left=13.0, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_INTEGER,
                label="Orders",
                variance_key="qaff.orders_vs_ly",
            ),
            DataSlot(
                name="kpi_cvr",
                slot_type=SlotType.KPI_VALUE,
                data_key="qaff.cvr",
                position=Position(left=17.0, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_PERCENTAGE,
                label="CVR",
                variance_key="qaff.cvr_vs_ly",
            ),
            # ROAS chart — monthly bars
            DataSlot(
                name="roas_chart",
                slot_type=SlotType.CHART,
                data_key="qaff.roas_chart",
                position=Position(left=0.5, top=3.5, width=10.0, height=5.0),
                chart_type=ChartType.COLUMN_CLUSTERED,
                categories_key="qaff.months",
                series=[
                    ChartSeries(name="ROAS TY", data_key="qaff.roas_monthly_ty",
                                color="#0065E0"),
                    ChartSeries(name="ROAS LY", data_key="qaff.roas_monthly_ly",
                                color="#1C2B33"),
                ],
            ),
            # Top publishers table
            DataSlot(
                name="publisher_table",
                slot_type=SlotType.TABLE,
                data_key="qaff.publisher_table",
                position=Position(left=11.0, top=3.5, width=10.0, height=7.5),
                row_data_key="qaff.publisher_rows",
                columns=[
                    TableColumn(header="Publisher", data_key="publisher_name", width_inches=3.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Revenue", data_key="revenue", width_inches=2.0,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="revenue_vs_ly", width_inches=1.3,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="COS", data_key="cos", width_inches=1.0,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Orders", data_key="orders", width_inches=1.3,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="AOV", data_key="aov", width_inches=1.2,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 13: PPC Deep Dive
# ---------------------------------------------------------------------------
def _slide_ppc() -> SlideSchema:
    return SlideSchema(
        index=13,
        name="qbr_ppc",
        title="PPC Performance",
        slide_type=SlideType.DATA,
        data_source="tracker:ppc",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qppc.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            # PPC headline KPIs
            DataSlot(
                name="kpi_revenue",
                slot_type=SlotType.KPI_VALUE,
                data_key="qppc.revenue",
                position=Position(left=0.5, top=1.5, width=4.0, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_CURRENCY,
                label="Revenue",
                variance_key="qppc.revenue_vs_ly",
            ),
            DataSlot(
                name="kpi_roas",
                slot_type=SlotType.KPI_VALUE,
                data_key="qppc.roas",
                position=Position(left=5.0, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_NUMBER,
                label="ROAS",
                variance_key="qppc.roas_vs_ly",
            ),
            DataSlot(
                name="kpi_cos",
                slot_type=SlotType.KPI_VALUE,
                data_key="qppc.cos",
                position=Position(left=9.0, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_PERCENTAGE,
                label="COS",
                variance_key="qppc.cos_vs_ly",
            ),
            DataSlot(
                name="kpi_spend",
                slot_type=SlotType.KPI_VALUE,
                data_key="qppc.spend",
                position=Position(left=13.0, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_CURRENCY,
                label="Spend",
                variance_key="qppc.spend_vs_ly",
            ),
            DataSlot(
                name="kpi_cvr",
                slot_type=SlotType.KPI_VALUE,
                data_key="qppc.cvr",
                position=Position(left=17.0, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_PERCENTAGE,
                label="CVR",
                variance_key="qppc.cvr_vs_ly",
            ),
            # Revenue + COS monthly chart
            DataSlot(
                name="ppc_chart",
                slot_type=SlotType.CHART,
                data_key="qppc.chart",
                position=Position(left=0.5, top=3.5, width=14.0, height=7.5),
                chart_type=ChartType.COLUMN_CLUSTERED,
                categories_key="qppc.months",
                series=[
                    ChartSeries(name="Revenue", data_key="qppc.revenue_monthly",
                                color="#0065E0"),
                    ChartSeries(name="Spend", data_key="qppc.spend_monthly",
                                color="#1C2B33"),
                ],
            ),
            # ROAS doughnut
            DataSlot(
                name="roas_gauge",
                slot_type=SlotType.CHART,
                data_key="qppc.roas_gauge",
                position=Position(left=15.5, top=3.5, width=5.5, height=4.0),
                chart_type=ChartType.DOUGHNUT,
                series=[
                    ChartSeries(name="ROAS", data_key="qppc.roas_achieved",
                                color="#0065E0"),
                    ChartSeries(name="Remaining", data_key="qppc.roas_remaining",
                                color="#D1D5DB"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 14: SEO Deep Dive
# ---------------------------------------------------------------------------
def _slide_seo() -> SlideSchema:
    return SlideSchema(
        index=14,
        name="qbr_seo",
        title="SEO Performance",
        slide_type=SlideType.DATA,
        data_source="tracker:organic",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qseo.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            # SEO headline KPIs
            DataSlot(
                name="kpi_revenue",
                slot_type=SlotType.KPI_VALUE,
                data_key="qseo.revenue",
                position=Position(left=0.5, top=1.5, width=4.0, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_CURRENCY,
                label="Revenue",
                variance_key="qseo.revenue_vs_ly",
            ),
            DataSlot(
                name="kpi_sessions",
                slot_type=SlotType.KPI_VALUE,
                data_key="qseo.sessions",
                position=Position(left=5.0, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_INTEGER,
                label="Sessions",
                variance_key="qseo.sessions_vs_ly",
            ),
            DataSlot(
                name="kpi_cvr",
                slot_type=SlotType.KPI_VALUE,
                data_key="qseo.cvr",
                position=Position(left=9.0, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_PERCENTAGE,
                label="CVR",
                variance_key="qseo.cvr_vs_ly",
            ),
            DataSlot(
                name="kpi_orders",
                slot_type=SlotType.KPI_VALUE,
                data_key="qseo.orders",
                position=Position(left=13.0, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_INTEGER,
                label="Orders",
                variance_key="qseo.orders_vs_ly",
            ),
            DataSlot(
                name="kpi_aov",
                slot_type=SlotType.KPI_VALUE,
                data_key="qseo.aov",
                position=Position(left=17.0, top=1.5, width=3.5, height=1.5),
                font=FontSpec(name="DM Sans", size_pt=36.0, bold=True),
                format_rule=_CURRENCY,
                label="AOV",
                variance_key="qseo.aov_vs_ly",
            ),
            # Monthly sessions trend line
            DataSlot(
                name="sessions_chart",
                slot_type=SlotType.CHART,
                data_key="qseo.sessions_chart",
                position=Position(left=0.5, top=3.5, width=14.0, height=7.5),
                chart_type=ChartType.LINE,
                categories_key="qseo.months",
                series=[
                    ChartSeries(name="Sessions TY", data_key="qseo.sessions_monthly_ty",
                                color="#0065E0"),
                    ChartSeries(name="Sessions LY", data_key="qseo.sessions_monthly_ly",
                                color="#1C2B33"),
                ],
            ),
            # Revenue doughnut
            DataSlot(
                name="revenue_gauge",
                slot_type=SlotType.CHART,
                data_key="qseo.revenue_gauge",
                position=Position(left=15.5, top=3.5, width=5.5, height=4.0),
                chart_type=ChartType.DOUGHNUT,
                series=[
                    ChartSeries(name="Revenue", data_key="qseo.revenue_achieved",
                                color="#0065E0"),
                    ChartSeries(name="Remaining", data_key="qseo.revenue_remaining",
                                color="#D1D5DB"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 15: Section Divider — Product & Promotion
# ---------------------------------------------------------------------------
def _slide_divider_product() -> SlideSchema:
    return SlideSchema(
        index=15,
        name="divider_product",
        title="Product & Promotion Performance",
        slide_type=SlideType.SECTION_DIVIDER,
        data_source="static",
        layout="Title Only",
        is_static=True,
        slots=[
            DataSlot(
                name="section_title",
                slot_type=SlotType.SECTION_DIVIDER,
                data_key="qdivider.product_title",
                position=Position(left=0.0, top=0.0, width=_W, height=_H),
                font=_DIVIDER_TITLE,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 16: Product Performance (Quarterly)
# ---------------------------------------------------------------------------
def _slide_product_performance() -> SlideSchema:
    return SlideSchema(
        index=16,
        name="qbr_product",
        title="Product Performance",
        slide_type=SlideType.DATA,
        data_source="product_sales_csv",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qproduct.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            DataSlot(
                name="product_table",
                slot_type=SlotType.TABLE,
                data_key="qproduct.table",
                position=Position(left=0.5, top=1.5, width=21.0, height=9.5),
                row_data_key="qproduct.rows",
                columns=[
                    TableColumn(header="Product", data_key="product_name", width_inches=5.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Units", data_key="units", width_inches=1.5,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="units_vs_ly", width_inches=1.3,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Revenue", data_key="revenue", width_inches=2.0,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="revenue_vs_ly", width_inches=1.3,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="AOV", data_key="aov", width_inches=1.5,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="ASP", data_key="avg_selling_price", width_inches=1.5,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Discount %", data_key="discount_pct", width_inches=1.5,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="New Cust", data_key="new_customers", width_inches=1.5,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Mix %", data_key="revenue_mix_pct", width_inches=1.5,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 17: Promotion Performance (Quarterly)
# ---------------------------------------------------------------------------
def _slide_promotion_performance() -> SlideSchema:
    return SlideSchema(
        index=17,
        name="qbr_promotion",
        title="Promotion Performance",
        slide_type=SlideType.DATA,
        data_source="offer_performance_csv",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qpromo.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            DataSlot(
                name="promotion_table",
                slot_type=SlotType.TABLE,
                data_key="qpromo.table",
                position=Position(left=0.5, top=1.5, width=21.0, height=9.5),
                row_data_key="qpromo.rows",
                columns=[
                    TableColumn(header="Promotion", data_key="promotion_name", width_inches=5.5,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Channel", data_key="channel", width_inches=2.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Redemptions", data_key="redemptions", width_inches=2.0,
                                format_rule=_INTEGER, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="redemptions_vs_ly", width_inches=1.5,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Revenue", data_key="revenue", width_inches=2.0,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="vs LY", data_key="revenue_vs_ly", width_inches=1.5,
                                format_rule=_VARIANCE, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Discount", data_key="discount_amount", width_inches=2.0,
                                format_rule=_CURRENCY, font=_TABLE_HEADER, alignment="right"),
                    TableColumn(header="Disc/Rev %", data_key="discount_revenue_pct", width_inches=1.5,
                                format_rule=_PERCENTAGE, font=_TABLE_HEADER, alignment="right"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 18: Customer Service Overview
# ---------------------------------------------------------------------------
def _slide_customer_service() -> SlideSchema:
    return SlideSchema(
        index=18,
        name="qbr_customer_service",
        title="Customer Service Overview",
        slide_type=SlideType.MANUAL,
        data_source="manual",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qcs.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            DataSlot(
                name="cs_narrative",
                slot_type=SlotType.TEXT,
                data_key="qcs.narrative",
                position=Position(left=0.5, top=1.5, width=20.5, height=9.5),
                font=_BODY,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 19: Fulfilment Overview
# ---------------------------------------------------------------------------
def _slide_fulfilment() -> SlideSchema:
    return SlideSchema(
        index=19,
        name="qbr_fulfilment",
        title="Fulfilment Overview",
        slide_type=SlideType.MANUAL,
        data_source="manual",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qfulfilment.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            DataSlot(
                name="fulfilment_narrative",
                slot_type=SlotType.TEXT,
                data_key="qfulfilment.narrative",
                position=Position(left=0.5, top=1.5, width=20.5, height=9.5),
                font=_BODY,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 20: Growth Opportunities
# ---------------------------------------------------------------------------
def _slide_growth_opportunities() -> SlideSchema:
    return SlideSchema(
        index=20,
        name="qbr_growth",
        title="Growth Opportunities",
        slide_type=SlideType.MANUAL,
        data_source="manual",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qgrowth.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            DataSlot(
                name="growth_narrative",
                slot_type=SlotType.TEXT,
                data_key="qgrowth.narrative",
                position=Position(left=0.5, top=1.5, width=20.5, height=9.5),
                font=_BODY,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 21: Section Divider — Outlook
# ---------------------------------------------------------------------------
def _slide_divider_outlook() -> SlideSchema:
    return SlideSchema(
        index=21,
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
                data_key="qdivider.outlook_title",
                position=Position(left=0.0, top=0.0, width=_W, height=_H),
                font=_DIVIDER_TITLE,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 22: Quarter Lookahead
# ---------------------------------------------------------------------------
def _slide_lookahead() -> SlideSchema:
    return SlideSchema(
        index=22,
        name="qbr_lookahead",
        title="Quarter Lookahead",
        slide_type=SlideType.MANUAL,
        data_source="manual",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qlookahead.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            DataSlot(
                name="lookahead_narrative",
                slot_type=SlotType.TEXT,
                data_key="qlookahead.narrative",
                position=Position(left=0.5, top=1.5, width=20.5, height=9.5),
                font=_BODY,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 23: Key Projects / Roadmap
# ---------------------------------------------------------------------------
def _slide_projects() -> SlideSchema:
    return SlideSchema(
        index=23,
        name="qbr_projects",
        title="Key Projects",
        slide_type=SlideType.MANUAL,
        data_source="manual",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qprojects.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            DataSlot(
                name="projects_table",
                slot_type=SlotType.TABLE,
                data_key="qprojects.table",
                position=Position(left=0.5, top=1.5, width=21.0, height=9.5),
                row_data_key="qprojects.rows",
                columns=[
                    TableColumn(header="Project", data_key="project_name", width_inches=5.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Owner", data_key="owner", width_inches=3.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Status", data_key="status", width_inches=2.0,
                                font=_TABLE_HEADER, alignment="center"),
                    TableColumn(header="Target Date", data_key="target_date", width_inches=2.5,
                                font=_TABLE_HEADER, alignment="center"),
                    TableColumn(header="Notes", data_key="notes", width_inches=5.0,
                                font=_TABLE_HEADER, alignment="left"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 24: Section Divider — Platform
# ---------------------------------------------------------------------------
def _slide_divider_platform() -> SlideSchema:
    return SlideSchema(
        index=24,
        name="divider_platform",
        title="Platform",
        slide_type=SlideType.SECTION_DIVIDER,
        data_source="static",
        layout="Title Only",
        is_static=True,
        slots=[
            DataSlot(
                name="section_title",
                slot_type=SlotType.SECTION_DIVIDER,
                data_key="qdivider.platform_title",
                position=Position(left=0.0, top=0.0, width=_W, height=_H),
                font=_DIVIDER_TITLE,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 25: Platform Roadmap
# ---------------------------------------------------------------------------
def _slide_platform_roadmap() -> SlideSchema:
    return SlideSchema(
        index=25,
        name="qbr_platform_roadmap",
        title="Platform Roadmap",
        slide_type=SlideType.MANUAL,
        data_source="manual",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qplatform.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            DataSlot(
                name="platform_narrative",
                slot_type=SlotType.TEXT,
                data_key="qplatform.narrative",
                position=Position(left=0.5, top=1.5, width=20.5, height=9.5),
                font=_BODY,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 26: Section Divider — Close
# ---------------------------------------------------------------------------
def _slide_divider_close() -> SlideSchema:
    return SlideSchema(
        index=26,
        name="divider_close",
        title="Closing",
        slide_type=SlideType.SECTION_DIVIDER,
        data_source="static",
        layout="Title Only",
        is_static=True,
        slots=[
            DataSlot(
                name="section_title",
                slot_type=SlotType.SECTION_DIVIDER,
                data_key="qdivider.close_title",
                position=Position(left=0.0, top=0.0, width=_W, height=_H),
                font=_DIVIDER_TITLE,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 27: Critical Path Planning
# ---------------------------------------------------------------------------
def _slide_critical_path() -> SlideSchema:
    return SlideSchema(
        index=27,
        name="qbr_critical_path",
        title="Critical Path",
        slide_type=SlideType.MANUAL,
        data_source="manual",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qcritical_path.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            DataSlot(
                name="critical_path_table",
                slot_type=SlotType.TABLE,
                data_key="qcritical_path.table",
                position=Position(left=0.5, top=1.5, width=21.0, height=9.5),
                row_data_key="qcritical_path.rows",
                columns=[
                    TableColumn(header="Item", data_key="item", width_inches=5.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Priority", data_key="priority", width_inches=2.0,
                                font=_TABLE_HEADER, alignment="center"),
                    TableColumn(header="Owner", data_key="owner", width_inches=3.0,
                                font=_TABLE_HEADER, alignment="left"),
                    TableColumn(header="Deadline", data_key="deadline", width_inches=2.5,
                                font=_TABLE_HEADER, alignment="center"),
                    TableColumn(header="Status", data_key="status", width_inches=2.0,
                                font=_TABLE_HEADER, alignment="center"),
                    TableColumn(header="Notes", data_key="notes", width_inches=4.0,
                                font=_TABLE_HEADER, alignment="left"),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Slide 28: Next Steps / Close
# ---------------------------------------------------------------------------
def _slide_next_steps() -> SlideSchema:
    return SlideSchema(
        index=28,
        name="qbr_next_steps",
        title="Next Steps",
        slide_type=SlideType.MANUAL,
        data_source="manual",
        layout="Title Only",
        slots=[
            DataSlot(
                name="slide_title",
                slot_type=SlotType.TEXT,
                data_key="qnext_steps.title",
                position=Position(left=0.5, top=0.3, width=20.0, height=0.8),
                font=_HEADER,
            ),
            DataSlot(
                name="action_items",
                slot_type=SlotType.TEXT,
                data_key="qnext_steps.items",
                position=Position(left=0.5, top=1.5, width=20.5, height=9.5),
                font=_BODY,
            ),
        ],
    )


# ---------------------------------------------------------------------------
# Assembled schema
# ---------------------------------------------------------------------------

def build_qbr_schema() -> TemplateSchema:
    """Build and return the complete 29-slide QBR schema."""
    return TemplateSchema(
        name="No7 US Quarterly Business Review",
        report_type="qbr",
        width_inches=_W,
        height_inches=_H,
        naming_convention="No7 US x THGi QBR - Q{quarter} {year}.pptx",
        design=DesignSystem(
            # QBR uses Ingenuity branding with slightly different colors
            brand_blue="#0065E2",
            dark_blue="#190263",
            dark_grey="#1C2B33",
            # Larger default text sizes for oversized slides
            title_size_pt=44.0,
            header_size_pt=30.0,
            body_size_pt=16.0,
            kpi_number_size_pt=60.0,
            kpi_label_size_pt=14.0,
            caption_size_pt=11.0,
        ),
        slides=[
            _slide_cover(),                  # 0
            _slide_agenda(),                 # 1
            _slide_executive_summary(),      # 2
            _slide_divider_strategy(),       # 3
            _slide_strategy_review(),        # 4
            _slide_successes(),              # 5
            _slide_challenges(),             # 6
            _slide_revenue_chart(),          # 7
            _slide_kpi_overview(),           # 8
            _slide_divider_channels(),       # 9
            _slide_channel_mix(),            # 10
            _slide_crm(),                    # 11
            _slide_affiliate(),              # 12
            _slide_ppc(),                    # 13
            _slide_seo(),                    # 14
            _slide_divider_product(),        # 15
            _slide_product_performance(),    # 16
            _slide_promotion_performance(),  # 17
            _slide_customer_service(),       # 18
            _slide_fulfilment(),             # 19
            _slide_growth_opportunities(),   # 20
            _slide_divider_outlook(),        # 21
            _slide_lookahead(),              # 22
            _slide_projects(),               # 23
            _slide_divider_platform(),       # 24
            _slide_platform_roadmap(),       # 25
            _slide_divider_close(),          # 26
            _slide_critical_path(),          # 27
            _slide_next_steps(),             # 28
        ],
    )
