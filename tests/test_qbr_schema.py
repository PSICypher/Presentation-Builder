"""Tests for the QBR report schema definition."""

import pytest

from src.schema.models import (
    ChartType,
    FormatType,
    SlideType,
    SlotType,
    TemplateSchema,
)
from src.schema.qbr_report import build_qbr_schema


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def schema():
    return build_qbr_schema()


# ---------------------------------------------------------------------------
# Top-level schema properties
# ---------------------------------------------------------------------------

class TestSchemaProperties:
    def test_name(self, schema):
        assert schema.name == "No7 US Quarterly Business Review"

    def test_report_type(self, schema):
        assert schema.report_type == "qbr"

    def test_dimensions_oversized(self, schema):
        assert schema.width_inches == pytest.approx(21.986)
        assert schema.height_inches == pytest.approx(12.368)

    def test_naming_convention(self, schema):
        assert "QBR" in schema.naming_convention
        assert "{quarter}" in schema.naming_convention
        assert "{year}" in schema.naming_convention

    def test_slide_count(self, schema):
        assert len(schema.slides) == 29


# ---------------------------------------------------------------------------
# Design system
# ---------------------------------------------------------------------------

class TestDesignSystem:
    def test_brand_blue(self, schema):
        assert schema.design.brand_blue == "#0065E2"

    def test_primary_font(self, schema):
        assert schema.design.primary_font == "DM Sans"

    def test_larger_kpi_numbers(self, schema):
        # QBR uses 60pt for KPI numbers (vs 48pt in monthly)
        assert schema.design.kpi_number_size_pt == 60.0

    def test_larger_title(self, schema):
        assert schema.design.title_size_pt == 44.0

    def test_positive_negative_colors(self, schema):
        assert schema.design.positive == "#00AA00"
        assert schema.design.negative == "#CC0000"


# ---------------------------------------------------------------------------
# Slide structure and ordering
# ---------------------------------------------------------------------------

class TestSlideStructure:
    def test_slide_indices_sequential(self, schema):
        indices = [s.index for s in schema.slides]
        assert indices == list(range(29))

    def test_slide_names_unique(self, schema):
        names = [s.name for s in schema.slides]
        assert len(names) == len(set(names))

    def test_first_slide_is_cover(self, schema):
        assert schema.slides[0].slide_type == SlideType.COVER
        assert schema.slides[0].name == "qbr_cover"

    def test_last_slide_is_next_steps(self, schema):
        assert schema.slides[28].name == "qbr_next_steps"
        assert schema.slides[28].slide_type == SlideType.MANUAL

    def test_slide_types_by_name(self, schema):
        expected = {
            "qbr_cover": SlideType.COVER,
            "qbr_agenda": SlideType.TABLE_OF_CONTENTS,
            "qbr_executive_summary": SlideType.DATA,
            "divider_strategy": SlideType.SECTION_DIVIDER,
            "qbr_strategy_review": SlideType.MANUAL,
            "qbr_successes": SlideType.MANUAL,
            "qbr_challenges": SlideType.MANUAL,
            "qbr_revenue_chart": SlideType.DATA,
            "qbr_kpi_overview": SlideType.DATA,
            "divider_channels": SlideType.SECTION_DIVIDER,
            "qbr_channel_mix": SlideType.DATA,
            "qbr_crm": SlideType.DATA,
            "qbr_affiliate": SlideType.DATA,
            "qbr_ppc": SlideType.DATA,
            "qbr_seo": SlideType.DATA,
            "divider_product": SlideType.SECTION_DIVIDER,
            "qbr_product": SlideType.DATA,
            "qbr_promotion": SlideType.DATA,
            "qbr_customer_service": SlideType.MANUAL,
            "qbr_fulfilment": SlideType.MANUAL,
            "qbr_growth": SlideType.MANUAL,
            "divider_outlook": SlideType.SECTION_DIVIDER,
            "qbr_lookahead": SlideType.MANUAL,
            "qbr_projects": SlideType.MANUAL,
            "divider_platform": SlideType.SECTION_DIVIDER,
            "qbr_platform_roadmap": SlideType.MANUAL,
            "divider_close": SlideType.SECTION_DIVIDER,
            "qbr_critical_path": SlideType.MANUAL,
            "qbr_next_steps": SlideType.MANUAL,
        }
        for name, expected_type in expected.items():
            slide = schema.get_slide(name)
            assert slide is not None, f"Slide '{name}' not found"
            assert slide.slide_type == expected_type, (
                f"Slide '{name}' type mismatch: {slide.slide_type} != {expected_type}")


# ---------------------------------------------------------------------------
# Section dividers
# ---------------------------------------------------------------------------

class TestDividers:
    def test_divider_count(self, schema):
        dividers = [s for s in schema.slides if s.slide_type == SlideType.SECTION_DIVIDER]
        assert len(dividers) == 6

    def test_all_dividers_static(self, schema):
        for s in schema.slides:
            if s.slide_type == SlideType.SECTION_DIVIDER:
                assert s.is_static, f"Divider '{s.name}' not marked static"
                assert s.data_source == "static"

    def test_divider_data_keys(self, schema):
        divider_keys = set()
        for s in schema.slides:
            if s.slide_type == SlideType.SECTION_DIVIDER:
                for slot in s.slots:
                    divider_keys.add(slot.data_key)
        expected = {
            "qdivider.strategy_title",
            "qdivider.channels_title",
            "qdivider.product_title",
            "qdivider.outlook_title",
            "qdivider.platform_title",
            "qdivider.close_title",
        }
        assert divider_keys == expected

    def test_divider_full_size(self, schema):
        for s in schema.slides:
            if s.slide_type == SlideType.SECTION_DIVIDER:
                for slot in s.slots:
                    assert slot.position.width == pytest.approx(21.986)
                    assert slot.position.height == pytest.approx(12.368)


# ---------------------------------------------------------------------------
# Data slides
# ---------------------------------------------------------------------------

class TestDataSlides:
    def test_data_slide_count(self, schema):
        data_slides = [s for s in schema.slides if s.slide_type == SlideType.DATA]
        # exec_summary, revenue_chart, kpi_overview, channel_mix,
        # crm, affiliate, ppc, seo, product, promotion = 10
        assert len(data_slides) == 10

    def test_data_slides_not_static(self, schema):
        for s in schema.slides:
            if s.slide_type == SlideType.DATA:
                assert not s.is_static, f"Data slide '{s.name}' marked static"

    def test_data_slides_have_source(self, schema):
        for s in schema.slides:
            if s.slide_type == SlideType.DATA:
                assert s.data_source != "static", (
                    f"Data slide '{s.name}' has static source")
                assert s.data_source != "", (
                    f"Data slide '{s.name}' has empty source")


# ---------------------------------------------------------------------------
# Cover slide
# ---------------------------------------------------------------------------

class TestCoverSlide:
    def test_cover_has_kpis(self, schema):
        cover = schema.get_slide("qbr_cover")
        kpi_slots = [s for s in cover.slots if s.slot_type == SlotType.KPI_VALUE]
        assert len(kpi_slots) == 6

    def test_cover_kpi_data_keys(self, schema):
        cover = schema.get_slide("qbr_cover")
        kpi_keys = {s.data_key for s in cover.slots if s.slot_type == SlotType.KPI_VALUE}
        expected = {
            "qcover.total_revenue",
            "qcover.total_orders",
            "qcover.aov",
            "qcover.new_customers",
            "qcover.cvr",
            "qcover.cos",
        }
        assert kpi_keys == expected

    def test_cover_kpis_have_variance_keys(self, schema):
        cover = schema.get_slide("qbr_cover")
        for slot in cover.slots:
            if slot.slot_type == SlotType.KPI_VALUE:
                assert slot.variance_key is not None, (
                    f"Cover KPI '{slot.name}' missing variance_key")
                assert slot.variance_key.startswith("qcover.")

    def test_cover_kpi_format_rules(self, schema):
        cover = schema.get_slide("qbr_cover")
        format_map = {}
        for slot in cover.slots:
            if slot.slot_type == SlotType.KPI_VALUE:
                format_map[slot.name] = slot.format_rule.format_type

        assert format_map["kpi_revenue"] == FormatType.CURRENCY
        assert format_map["kpi_orders"] == FormatType.NUMBER
        assert format_map["kpi_aov"] == FormatType.CURRENCY
        assert format_map["kpi_new_customers"] == FormatType.NUMBER
        assert format_map["kpi_cvr"] == FormatType.PERCENTAGE
        assert format_map["kpi_cos"] == FormatType.PERCENTAGE


# ---------------------------------------------------------------------------
# Executive summary
# ---------------------------------------------------------------------------

class TestExecutiveSummary:
    def test_exec_has_table(self, schema):
        slide = schema.get_slide("qbr_executive_summary")
        tables = [s for s in slide.slots if s.slot_type == SlotType.TABLE]
        assert len(tables) == 1

    def test_exec_table_columns(self, schema):
        slide = schema.get_slide("qbr_executive_summary")
        table = next(s for s in slide.slots if s.slot_type == SlotType.TABLE)
        headers = [c.header for c in table.columns]
        assert "Channel" in headers
        assert "Revenue" in headers
        assert "vs Target" in headers
        assert "vs LY" in headers
        assert "Orders" in headers
        assert "CVR" in headers
        assert "AOV" in headers
        assert "COS" in headers
        assert "Contribution" in headers

    def test_exec_has_contribution_column(self, schema):
        slide = schema.get_slide("qbr_executive_summary")
        table = next(s for s in slide.slots if s.slot_type == SlotType.TABLE)
        contribution = next(c for c in table.columns if c.data_key == "contribution_pct")
        assert contribution.format_rule.format_type == FormatType.PERCENTAGE

    def test_exec_has_theme_boxes(self, schema):
        slide = schema.get_slide("qbr_executive_summary")
        text_slots = [s for s in slide.slots if s.slot_type == SlotType.TEXT]
        theme_slots = [s for s in text_slots if s.data_key.startswith("qexec.theme_")]
        assert len(theme_slots) == 3


# ---------------------------------------------------------------------------
# Revenue chart slide
# ---------------------------------------------------------------------------

class TestRevenueChart:
    def test_has_column_chart(self, schema):
        slide = schema.get_slide("qbr_revenue_chart")
        charts = [s for s in slide.slots
                  if s.slot_type == SlotType.CHART and s.chart_type == ChartType.COLUMN_CLUSTERED]
        assert len(charts) >= 1

    def test_column_chart_series(self, schema):
        slide = schema.get_slide("qbr_revenue_chart")
        chart = next(s for s in slide.slots
                     if s.chart_type == ChartType.COLUMN_CLUSTERED)
        assert len(chart.series) == 3
        series_names = {s.name for s in chart.series}
        assert "Revenue TY" in series_names
        assert "Revenue LY" in series_names
        assert "Target" in series_names

    def test_has_doughnut_gauges(self, schema):
        slide = schema.get_slide("qbr_revenue_chart")
        doughnuts = [s for s in slide.slots
                     if s.slot_type == SlotType.CHART and s.chart_type == ChartType.DOUGHNUT]
        assert len(doughnuts) == 2

    def test_has_monthly_breakdown_table(self, schema):
        slide = schema.get_slide("qbr_revenue_chart")
        tables = [s for s in slide.slots if s.slot_type == SlotType.TABLE]
        assert len(tables) == 1
        table = tables[0]
        assert table.row_data_key == "qrevenue.monthly_rows"
        month_col = next(c for c in table.columns if c.data_key == "month")
        assert month_col is not None


# ---------------------------------------------------------------------------
# KPI overview slide
# ---------------------------------------------------------------------------

class TestKPIOverview:
    def test_six_gauges(self, schema):
        slide = schema.get_slide("qbr_kpi_overview")
        gauges = [s for s in slide.slots if s.chart_type == ChartType.DOUGHNUT]
        assert len(gauges) == 6

    def test_gauge_data_keys(self, schema):
        slide = schema.get_slide("qbr_kpi_overview")
        gauge_keys = {s.data_key for s in slide.slots if s.chart_type == ChartType.DOUGHNUT}
        expected = {
            "qkpi.revenue_gauge",
            "qkpi.aov_gauge",
            "qkpi.cvr_gauge",
            "qkpi.cos_gauge",
            "qkpi.nc_gauge",
            "qkpi.orders_gauge",
        }
        assert gauge_keys == expected


# ---------------------------------------------------------------------------
# Channel deep-dive slides
# ---------------------------------------------------------------------------

class TestChannelDeepDives:
    def test_crm_has_kpis_and_table(self, schema):
        slide = schema.get_slide("qbr_crm")
        kpis = [s for s in slide.slots if s.slot_type == SlotType.KPI_VALUE]
        tables = [s for s in slide.slots if s.slot_type == SlotType.TABLE]
        charts = [s for s in slide.slots if s.slot_type == SlotType.CHART]
        assert len(kpis) == 5
        assert len(tables) == 1
        assert len(charts) == 1

    def test_crm_has_next_quarter_strategy(self, schema):
        slide = schema.get_slide("qbr_crm")
        strategy = next(
            (s for s in slide.slots if s.data_key == "qcrm.next_quarter_strategy"),
            None,
        )
        assert strategy is not None
        assert strategy.slot_type == SlotType.TEXT

    def test_affiliate_has_kpis_chart_table(self, schema):
        slide = schema.get_slide("qbr_affiliate")
        kpis = [s for s in slide.slots if s.slot_type == SlotType.KPI_VALUE]
        tables = [s for s in slide.slots if s.slot_type == SlotType.TABLE]
        charts = [s for s in slide.slots if s.slot_type == SlotType.CHART]
        assert len(kpis) == 5
        assert len(tables) == 1
        assert len(charts) == 1

    def test_ppc_has_chart_and_gauge(self, schema):
        slide = schema.get_slide("qbr_ppc")
        charts = [s for s in slide.slots if s.slot_type == SlotType.CHART]
        assert len(charts) == 2
        chart_types = {c.chart_type for c in charts}
        assert ChartType.COLUMN_CLUSTERED in chart_types
        assert ChartType.DOUGHNUT in chart_types

    def test_seo_has_line_chart(self, schema):
        slide = schema.get_slide("qbr_seo")
        line_charts = [s for s in slide.slots
                       if s.chart_type == ChartType.LINE]
        assert len(line_charts) == 1
        chart = line_charts[0]
        series_names = {s.name for s in chart.series}
        assert "Sessions TY" in series_names
        assert "Sessions LY" in series_names


# ---------------------------------------------------------------------------
# Product and promotion slides
# ---------------------------------------------------------------------------

class TestProductAndPromotion:
    def test_product_table_columns(self, schema):
        slide = schema.get_slide("qbr_product")
        table = next(s for s in slide.slots if s.slot_type == SlotType.TABLE)
        headers = [c.header for c in table.columns]
        assert "Product" in headers
        assert "Units" in headers
        assert "Revenue" in headers
        assert "Mix %" in headers

    def test_product_has_mix_pct(self, schema):
        slide = schema.get_slide("qbr_product")
        table = next(s for s in slide.slots if s.slot_type == SlotType.TABLE)
        mix_col = next(c for c in table.columns if c.data_key == "revenue_mix_pct")
        assert mix_col.format_rule.format_type == FormatType.PERCENTAGE

    def test_promotion_table_columns(self, schema):
        slide = schema.get_slide("qbr_promotion")
        table = next(s for s in slide.slots if s.slot_type == SlotType.TABLE)
        headers = [c.header for c in table.columns]
        assert "Promotion" in headers
        assert "Redemptions" in headers
        assert "Revenue" in headers
        assert "Disc/Rev %" in headers

    def test_promotion_has_discount_revenue_pct(self, schema):
        slide = schema.get_slide("qbr_promotion")
        table = next(s for s in slide.slots if s.slot_type == SlotType.TABLE)
        disc_col = next(c for c in table.columns if c.data_key == "discount_revenue_pct")
        assert disc_col.format_rule.format_type == FormatType.PERCENTAGE


# ---------------------------------------------------------------------------
# Manual/operational slides
# ---------------------------------------------------------------------------

class TestManualSlides:
    def test_manual_slide_count(self, schema):
        manual_slides = [s for s in schema.slides if s.slide_type == SlideType.MANUAL]
        # strategy_review, successes, challenges, customer_service,
        # fulfilment, growth, lookahead, projects, platform_roadmap,
        # critical_path, next_steps = 11
        assert len(manual_slides) == 11

    def test_all_manual_slides_have_title(self, schema):
        for s in schema.slides:
            if s.slide_type == SlideType.MANUAL:
                title_slots = [sl for sl in s.slots if sl.slot_type == SlotType.TEXT
                               and "title" in sl.data_key]
                assert len(title_slots) >= 1, (
                    f"Manual slide '{s.name}' missing title slot")

    def test_strategy_has_four_pillars(self, schema):
        slide = schema.get_slide("qbr_strategy_review")
        text_slots = [s for s in slide.slots if s.slot_type == SlotType.TEXT]
        pillar_slots = [s for s in text_slots if "pillar_" in s.data_key]
        assert len(pillar_slots) == 4

    def test_projects_has_table(self, schema):
        slide = schema.get_slide("qbr_projects")
        tables = [s for s in slide.slots if s.slot_type == SlotType.TABLE]
        assert len(tables) == 1
        table = tables[0]
        headers = [c.header for c in table.columns]
        assert "Project" in headers
        assert "Owner" in headers
        assert "Status" in headers
        assert "Target Date" in headers

    def test_critical_path_has_table(self, schema):
        slide = schema.get_slide("qbr_critical_path")
        tables = [s for s in slide.slots if s.slot_type == SlotType.TABLE]
        assert len(tables) == 1
        table = tables[0]
        headers = [c.header for c in table.columns]
        assert "Item" in headers
        assert "Priority" in headers
        assert "Deadline" in headers


# ---------------------------------------------------------------------------
# Data key namespace consistency
# ---------------------------------------------------------------------------

class TestDataKeyNamespace:
    def test_slot_level_keys_use_q_prefix(self, schema):
        """Verify that slot-level data keys (data_key, variance_key,
        row_data_key, categories_key) use the q-prefix namespace.
        Column-level data_keys are relative within row dicts and do not
        require a prefix."""
        for slide in schema.slides:
            for slot in slide.slots:
                for key in (slot.data_key, slot.variance_key,
                            slot.row_data_key, slot.categories_key):
                    if key is None:
                        continue
                    prefix = key.split(".")[0]
                    assert prefix.startswith("q"), (
                        f"Slot key '{key}' on slide '{slide.name}' "
                        f"does not use q-prefix namespace")

    def test_no_collision_with_monthly_namespace(self, schema):
        all_keys = schema.all_data_keys()
        monthly_prefixes = {
            "cover", "toc", "divider", "exec", "daily", "promo",
            "product", "crm", "affiliate", "seo", "upcoming", "next_steps",
        }
        for key in all_keys:
            prefix = key.split(".")[0]
            assert prefix not in monthly_prefixes, (
                f"Data key '{key}' collides with monthly namespace '{prefix}'")

    def test_no_duplicate_slot_names_within_slide(self, schema):
        for slide in schema.slides:
            names = [s.name for s in slide.slots]
            assert len(names) == len(set(names)), (
                f"Slide '{slide.name}' has duplicate slot names: {names}")


# ---------------------------------------------------------------------------
# Serialization round-trip
# ---------------------------------------------------------------------------

class TestSerialization:
    def test_round_trip(self, schema):
        d = schema.to_dict()
        restored = TemplateSchema.from_dict(d)

        assert restored.name == schema.name
        assert restored.report_type == schema.report_type
        assert restored.width_inches == pytest.approx(schema.width_inches)
        assert restored.height_inches == pytest.approx(schema.height_inches)
        assert len(restored.slides) == len(schema.slides)

    def test_round_trip_slide_names(self, schema):
        d = schema.to_dict()
        restored = TemplateSchema.from_dict(d)

        original_names = [s.name for s in schema.slides]
        restored_names = [s.name for s in restored.slides]
        assert original_names == restored_names

    def test_round_trip_data_keys(self, schema):
        d = schema.to_dict()
        restored = TemplateSchema.from_dict(d)

        assert schema.all_data_keys() == restored.all_data_keys()

    def test_round_trip_design_system(self, schema):
        d = schema.to_dict()
        restored = TemplateSchema.from_dict(d)

        assert restored.design.brand_blue == schema.design.brand_blue
        assert restored.design.primary_font == schema.design.primary_font
        assert restored.design.kpi_number_size_pt == schema.design.kpi_number_size_pt

    def test_to_dict_is_serializable(self, schema):
        """Verify that to_dict() produces a JSON-serializable structure."""
        import json
        d = schema.to_dict()
        serialized = json.dumps(d)
        assert isinstance(serialized, str)
        assert len(serialized) > 1000


# ---------------------------------------------------------------------------
# Convenience methods
# ---------------------------------------------------------------------------

class TestConvenienceMethods:
    def test_get_slide_by_name(self, schema):
        slide = schema.get_slide("qbr_cover")
        assert slide is not None
        assert slide.index == 0

    def test_get_slide_missing(self, schema):
        assert schema.get_slide("nonexistent") is None

    def test_data_slides_excludes_static(self, schema):
        data_slides = schema.data_slides()
        for s in data_slides:
            assert not s.is_static

    def test_all_data_keys_not_empty(self, schema):
        keys = schema.all_data_keys()
        assert len(keys) > 100
