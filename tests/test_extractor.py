"""Tests for the template extraction engine."""

import json
from pathlib import Path

import pytest

from src.extractor.template_extractor import (
    TemplateExtractor,
    _looks_like_kpi_value,
    _slugify,
    _unique_name,
    extract_from_file,
    extract_template,
)
from src.schema.loader import load_schema, save_schema
from src.schema.models import SlideType, SlotType

# Path to the analysis output produced by TemplateAnalyzer
ANALYSIS_PATH = Path(__file__).parent.parent / "output" / "template_analysis.json"


@pytest.fixture
def analysis_data():
    """Load the template analysis JSON (all 3 templates)."""
    with open(ANALYSIS_PATH) as f:
        return json.load(f)


@pytest.fixture
def monthly_analysis(analysis_data):
    """Return only the monthly report template analysis."""
    for t in analysis_data:
        if "monthly" in t["source_file"]:
            return t
    pytest.skip("monthly-report-template.pptx analysis not found")


@pytest.fixture
def monthly_schema(monthly_analysis):
    """Extract the monthly report schema."""
    return extract_template(monthly_analysis)


# ---------------------------------------------------------------------------
# Unit tests: helper functions
# ---------------------------------------------------------------------------

class TestLooksLikeKPIValue:
    def test_currency(self):
        assert _looks_like_kpi_value("$209.2K")
        assert _looks_like_kpi_value("$3.5m")
        assert _looks_like_kpi_value("$500")

    def test_count_with_suffix(self):
        assert _looks_like_kpi_value("3.6K")
        assert _looks_like_kpi_value("12.4M")

    def test_percentage(self):
        assert _looks_like_kpi_value("3.9%")
        assert _looks_like_kpi_value("+8.2%")

    def test_pure_number(self):
        assert _looks_like_kpi_value("1,234")
        assert _looks_like_kpi_value("42")

    def test_long_text_rejected(self):
        assert not _looks_like_kpi_value("No7 US - January Monthly Performance")
        assert not _looks_like_kpi_value("This is a long narrative string")

    def test_empty(self):
        assert not _looks_like_kpi_value("")
        assert not _looks_like_kpi_value("  ")

    def test_brand_name_rejected(self):
        # "No7" has digit but is not a KPI value
        assert not _looks_like_kpi_value("No7 US Monthly Report")


class TestSlugify:
    def test_basic(self):
        assert _slugify("Hello World") == "hello_world"

    def test_special_chars(self):
        assert _slugify("CRM Performance ") == "crm_performance"

    def test_truncation(self):
        result = _slugify("A" * 50)
        assert len(result) <= 40

    def test_empty(self):
        assert _slugify("") == ""


class TestUniqueName:
    def test_first_use(self):
        assert _unique_name("foo", set()) == "foo"

    def test_collision(self):
        used = {"foo"}
        assert _unique_name("foo", used) == "foo_1"

    def test_multiple_collisions(self):
        used = {"foo", "foo_1", "foo_2"}
        assert _unique_name("foo", used) == "foo_3"


# ---------------------------------------------------------------------------
# Integration tests: slide classification
# ---------------------------------------------------------------------------

class TestSlideClassification:
    def test_monthly_slide_count(self, monthly_schema):
        assert len(monthly_schema.slides) == 14

    def test_cover_slide(self, monthly_schema):
        cover = monthly_schema.slides[0]
        assert cover.slide_type == SlideType.COVER
        assert cover.name == "cover_kpis"

    def test_toc_slide(self, monthly_schema):
        toc = monthly_schema.slides[1]
        assert toc.slide_type == SlideType.TABLE_OF_CONTENTS
        assert toc.is_static

    def test_divider_slides(self, monthly_schema):
        dividers = [s for s in monthly_schema.slides if s.slide_type == SlideType.SECTION_DIVIDER]
        assert len(dividers) == 3
        for d in dividers:
            assert d.is_static
            assert d.data_source == "static"

    def test_data_slides(self, monthly_schema):
        data_slides = [s for s in monthly_schema.slides if s.slide_type == SlideType.DATA]
        assert len(data_slides) >= 7  # exec summary, daily, promo, product, crm, affiliate, seo

    def test_manual_slides(self, monthly_schema):
        manual = [s for s in monthly_schema.slides if s.slide_type == SlideType.MANUAL]
        assert len(manual) >= 1

    def test_slide_type_distribution(self, monthly_schema):
        """Verify the expected slide type pattern: cover, toc, dividers, data, manual."""
        types = [s.slide_type for s in monthly_schema.slides]
        assert types[0] == SlideType.COVER
        assert types[1] == SlideType.TABLE_OF_CONTENTS
        assert types[2] == SlideType.SECTION_DIVIDER


# ---------------------------------------------------------------------------
# Integration tests: data source inference
# ---------------------------------------------------------------------------

class TestDataSourceInference:
    def test_cover_source(self, monthly_schema):
        assert monthly_schema.slides[0].data_source == "tracker:mtd_reporting"

    def test_divider_source(self, monthly_schema):
        for s in monthly_schema.slides:
            if s.slide_type == SlideType.SECTION_DIVIDER:
                assert s.data_source == "static"

    def test_daily_performance_source(self, monthly_schema):
        daily = next(s for s in monthly_schema.slides if "daily" in s.name)
        assert daily.data_source == "tracker:daily"

    def test_promotion_source(self, monthly_schema):
        promo = next(s for s in monthly_schema.slides if "promotion" in s.name)
        assert promo.data_source == "offer_performance"

    def test_product_source(self, monthly_schema):
        prod = next(s for s in monthly_schema.slides if "product" in s.name)
        assert prod.data_source == "product_sales"

    def test_crm_source(self, monthly_schema):
        crm = next(s for s in monthly_schema.slides if s.name == "crm_performance")
        assert crm.data_source == "crm_data"

    def test_affiliate_source(self, monthly_schema):
        aff = next(s for s in monthly_schema.slides if "affiliate" in s.name)
        assert aff.data_source == "affiliate_data"

    def test_seo_source(self, monthly_schema):
        seo = next(s for s in monthly_schema.slides if s.name == "seo_performance")
        assert seo.data_source == "tracker:organic"


# ---------------------------------------------------------------------------
# Integration tests: slot extraction
# ---------------------------------------------------------------------------

class TestSlotExtraction:
    def test_cover_has_kpi_slots(self, monthly_schema):
        cover = monthly_schema.slides[0]
        kpi_slots = [s for s in cover.slots if s.slot_type == SlotType.KPI_VALUE]
        assert len(kpi_slots) >= 3

    def test_cover_has_title(self, monthly_schema):
        cover = monthly_schema.slides[0]
        titles = [s for s in cover.slots if s.name == "report_title"]
        assert len(titles) == 1

    def test_kpi_slots_have_labels(self, monthly_schema):
        cover = monthly_schema.slides[0]
        kpi_slots = [s for s in cover.slots if s.slot_type == SlotType.KPI_VALUE]
        labeled = [s for s in kpi_slots if s.label]
        assert len(labeled) >= 2  # Most KPIs should have labels

    def test_kpi_slots_have_format_rules(self, monthly_schema):
        cover = monthly_schema.slides[0]
        kpi_slots = [s for s in cover.slots if s.slot_type == SlotType.KPI_VALUE]
        for slot in kpi_slots:
            assert slot.format_rule is not None

    def test_data_slides_have_tables(self, monthly_schema):
        table_slides = []
        for s in monthly_schema.slides:
            tables = [sl for sl in s.slots if sl.slot_type == SlotType.TABLE]
            if tables:
                table_slides.append(s)
        assert len(table_slides) >= 5  # exec summary, promo, product, crm, affiliate, seo

    def test_table_slots_have_columns(self, monthly_schema):
        for slide in monthly_schema.slides:
            for slot in slide.slots:
                if slot.slot_type == SlotType.TABLE:
                    assert len(slot.columns) > 0, f"Table {slot.name} on {slide.name} has no columns"
                    assert slot.row_data_key, f"Table {slot.name} on {slide.name} has no row_data_key"

    def test_chart_slide_has_chart(self, monthly_schema):
        daily = next(s for s in monthly_schema.slides if "daily" in s.name)
        charts = [s for s in daily.slots if s.slot_type == SlotType.CHART]
        assert len(charts) == 1

    def test_divider_slots(self, monthly_schema):
        for s in monthly_schema.slides:
            if s.slide_type == SlideType.SECTION_DIVIDER:
                assert len(s.slots) == 1
                assert s.slots[0].slot_type == SlotType.SECTION_DIVIDER

    def test_all_slots_have_positions(self, monthly_schema):
        for slide in monthly_schema.slides:
            for slot in slide.slots:
                assert slot.position is not None
                assert slot.position.width > 0
                assert slot.position.height > 0

    def test_all_slots_have_data_keys(self, monthly_schema):
        for slide in monthly_schema.slides:
            for slot in slide.slots:
                assert slot.data_key, f"Slot {slot.name} on {slide.name} has no data_key"
                assert "." in slot.data_key  # Should be namespaced: slide.field


# ---------------------------------------------------------------------------
# Integration tests: uniqueness constraints
# ---------------------------------------------------------------------------

class TestUniqueness:
    def test_unique_slide_names(self, analysis_data):
        """No two slides should share the same name within a template."""
        for template in analysis_data:
            schema = extract_template(template)
            names = [s.name for s in schema.slides]
            assert len(names) == len(set(names)), \
                f"{schema.name}: duplicate slide names: {[n for n in names if names.count(n) > 1]}"

    def test_unique_slot_names_per_slide(self, analysis_data):
        """No two slots should share the same name within a single slide."""
        for template in analysis_data:
            schema = extract_template(template)
            for slide in schema.slides:
                names = [s.name for s in slide.slots]
                assert len(names) == len(set(names)), \
                    f"{schema.name} slide {slide.name}: duplicate slot names: {[n for n in names if names.count(n) > 1]}"


# ---------------------------------------------------------------------------
# Integration tests: design system extraction
# ---------------------------------------------------------------------------

class TestDesignSystem:
    def test_primary_font(self, monthly_schema):
        assert monthly_schema.design.primary_font == "DM Sans"

    def test_brand_blue(self, monthly_schema):
        blue = monthly_schema.design.brand_blue.lower()
        assert "0065" in blue or "0064" in blue  # Brand blue variant

    def test_has_color_palette(self, monthly_schema):
        d = monthly_schema.design
        assert d.dark_text == "#000000"
        assert d.white == "#FFFFFF"
        assert d.positive == "#00AA00"
        assert d.negative == "#CC0000"


# ---------------------------------------------------------------------------
# Integration tests: schema metadata
# ---------------------------------------------------------------------------

class TestSchemaMetadata:
    def test_monthly_report_type(self, monthly_schema):
        assert monthly_schema.report_type == "monthly"

    def test_monthly_dimensions(self, monthly_schema):
        assert monthly_schema.width_inches == 13.333
        assert monthly_schema.height_inches == 7.5

    def test_naming_convention(self, monthly_schema):
        assert "{month}" in monthly_schema.naming_convention
        assert "{year}" in monthly_schema.naming_convention

    def test_qbr_report_type(self, analysis_data):
        for t in analysis_data:
            if "qbr" in t["source_file"].lower():
                schema = extract_template(t)
                assert schema.report_type == "qbr"
                break


# ---------------------------------------------------------------------------
# Integration tests: round-trip serialization
# ---------------------------------------------------------------------------

class TestRoundTrip:
    def test_yaml_round_trip(self, monthly_schema, tmp_path):
        """Save to YAML and load back — all fields should be preserved."""
        path = tmp_path / "test_schema.yaml"
        save_schema(monthly_schema, path)
        loaded = load_schema(path)

        assert loaded.name == monthly_schema.name
        assert loaded.report_type == monthly_schema.report_type
        assert loaded.width_inches == monthly_schema.width_inches
        assert loaded.height_inches == monthly_schema.height_inches
        assert len(loaded.slides) == len(monthly_schema.slides)

        for orig, rt in zip(monthly_schema.slides, loaded.slides):
            assert orig.name == rt.name
            assert orig.slide_type == rt.slide_type
            assert orig.is_static == rt.is_static
            assert len(orig.slots) == len(rt.slots)
            for os, rs in zip(orig.slots, rt.slots):
                assert os.name == rs.name
                assert os.slot_type == rs.slot_type
                assert os.data_key == rs.data_key

    def test_all_data_keys_preserved(self, monthly_schema, tmp_path):
        """All data keys should survive serialization."""
        path = tmp_path / "keys_test.yaml"
        save_schema(monthly_schema, path)
        loaded = load_schema(path)
        assert loaded.all_data_keys() == monthly_schema.all_data_keys()


# ---------------------------------------------------------------------------
# Integration tests: multi-template extraction
# ---------------------------------------------------------------------------

class TestMultiTemplate:
    def test_extract_all_templates(self, analysis_data):
        """All 3 templates should extract without errors."""
        schemas = [extract_template(t) for t in analysis_data]
        assert len(schemas) == 3

    def test_extract_from_file(self):
        """extract_from_file should load and extract all templates."""
        schemas = extract_from_file(ANALYSIS_PATH)
        assert len(schemas) == 3
        names = [s.name for s in schemas]
        assert any("monthly" in n.lower() for n in names)
        assert any("qbr" in n.lower() for n in names)

    def test_data_slides_method(self, monthly_schema):
        """data_slides() should return only non-static slides."""
        data = monthly_schema.data_slides()
        for s in data:
            assert not s.is_static
