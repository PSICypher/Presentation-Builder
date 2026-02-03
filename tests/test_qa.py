"""Tests for the QA validation module."""

import io
import math

import pytest
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

from src.generator.pptx_builder import PPTXBuilder
from src.qa.validator import (
    Issue,
    QAResult,
    QAValidator,
    validate_presentation,
    _is_missing,
    _all_text_on_slide,
    _table_shapes,
    _chart_shapes,
)
from src.schema.models import (
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
from src.schema.monthly_report import build_monthly_report_schema


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def design():
    return DesignSystem()


@pytest.fixture
def minimal_schema(design):
    """Single-slide schema for focused tests."""
    return TemplateSchema(
        name="Test Report",
        report_type="monthly",
        width_inches=13.333,
        height_inches=7.5,
        design=design,
        slides=[
            SlideSchema(
                index=0,
                name="test_slide",
                title="Test Slide",
                slide_type=SlideType.DATA,
                data_source="test",
                slots=[],
            ),
        ],
    )


@pytest.fixture
def kpi_schema(design):
    """Schema with a single KPI slot."""
    return TemplateSchema(
        name="KPI Test",
        report_type="monthly",
        width_inches=13.333,
        height_inches=7.5,
        design=design,
        slides=[
            SlideSchema(
                index=0,
                name="kpi_slide",
                title="KPI Slide",
                slide_type=SlideType.DATA,
                data_source="test",
                slots=[
                    DataSlot(
                        name="revenue",
                        slot_type=SlotType.KPI_VALUE,
                        data_key="test.revenue",
                        position=Position(left=0.5, top=1.0, width=2.0, height=1.5),
                        font=FontSpec(name="DM Sans", size_pt=48.0, bold=True),
                        format_rule=FormatRule(FormatType.CURRENCY),
                        label="Revenue",
                        variance_key="test.revenue_var",
                    ),
                ],
            ),
        ],
    )


@pytest.fixture
def table_schema(design):
    """Schema with a single table slot."""
    return TemplateSchema(
        name="Table Test",
        report_type="monthly",
        width_inches=13.333,
        height_inches=7.5,
        design=design,
        slides=[
            SlideSchema(
                index=0,
                name="table_slide",
                title="Table Slide",
                slide_type=SlideType.DATA,
                data_source="test",
                slots=[
                    DataSlot(
                        name="test_table",
                        slot_type=SlotType.TABLE,
                        data_key="test.table",
                        position=Position(left=0.3, top=0.9, width=12.0, height=4.0),
                        row_data_key="test.rows",
                        columns=[
                            TableColumn(
                                header="Channel",
                                data_key="channel",
                                width_inches=2.0,
                                alignment="left",
                            ),
                            TableColumn(
                                header="Revenue",
                                data_key="revenue",
                                width_inches=1.5,
                                format_rule=FormatRule(FormatType.CURRENCY),
                                alignment="right",
                            ),
                            TableColumn(
                                header="vs Target",
                                data_key="vs_target",
                                width_inches=1.0,
                                format_rule=FormatRule(FormatType.VARIANCE_PERCENTAGE),
                                alignment="right",
                            ),
                        ],
                    ),
                ],
            ),
        ],
    )


@pytest.fixture
def chart_schema(design):
    """Schema with a column chart slot."""
    return TemplateSchema(
        name="Chart Test",
        report_type="monthly",
        width_inches=13.333,
        height_inches=7.5,
        design=design,
        slides=[
            SlideSchema(
                index=0,
                name="chart_slide",
                title="Chart Slide",
                slide_type=SlideType.DATA,
                data_source="test",
                slots=[
                    DataSlot(
                        name="daily_chart",
                        slot_type=SlotType.CHART,
                        data_key="test.chart",
                        position=Position(left=0.3, top=0.9, width=8.0, height=4.0),
                        chart_type=ChartType.COLUMN_CLUSTERED,
                        categories_key="test.dates",
                        series=[
                            ChartSeries(
                                name="Revenue",
                                data_key="test.revenue_series",
                                color="#0065E0",
                            ),
                            ChartSeries(
                                name="Target",
                                data_key="test.target_series",
                                color="#D1D5DB",
                            ),
                        ],
                    ),
                ],
            ),
        ],
    )


@pytest.fixture
def doughnut_schema(design):
    """Schema with a doughnut chart slot."""
    return TemplateSchema(
        name="Doughnut Test",
        report_type="monthly",
        width_inches=13.333,
        height_inches=7.5,
        design=design,
        slides=[
            SlideSchema(
                index=0,
                name="gauge_slide",
                title="Gauge Slide",
                slide_type=SlideType.DATA,
                data_source="test",
                slots=[
                    DataSlot(
                        name="gauge",
                        slot_type=SlotType.CHART,
                        data_key="test.gauge",
                        position=Position(left=0.5, top=5.5, width=2.0, height=1.5),
                        chart_type=ChartType.DOUGHNUT,
                        series=[
                            ChartSeries(name="Achieved", data_key="test.achieved", color="#0065E0"),
                            ChartSeries(name="Remaining", data_key="test.remaining", color="#D1D5DB"),
                        ],
                    ),
                ],
            ),
        ],
    )


@pytest.fixture
def divider_schema(design):
    """Schema with a section divider slide."""
    return TemplateSchema(
        name="Divider Test",
        report_type="monthly",
        width_inches=13.333,
        height_inches=7.5,
        design=design,
        slides=[
            SlideSchema(
                index=0,
                name="divider",
                title="eComm Performance",
                slide_type=SlideType.SECTION_DIVIDER,
                data_source="static",
                is_static=True,
                slots=[
                    DataSlot(
                        name="section_title",
                        slot_type=SlotType.SECTION_DIVIDER,
                        data_key="divider.title",
                        position=Position(left=0.0, top=0.0, width=13.333, height=7.5),
                        font=FontSpec(name="DM Sans", size_pt=36.0, bold=True, color="#FFFFFF"),
                    ),
                ],
            ),
        ],
    )


@pytest.fixture
def text_schema(design):
    """Schema with text slots."""
    return TemplateSchema(
        name="Text Test",
        report_type="monthly",
        width_inches=13.333,
        height_inches=7.5,
        design=design,
        slides=[
            SlideSchema(
                index=0,
                name="text_slide",
                title="Text Slide",
                slide_type=SlideType.DATA,
                data_source="test",
                slots=[
                    DataSlot(
                        name="title",
                        slot_type=SlotType.TEXT,
                        data_key="test.title",
                        position=Position(left=0.3, top=0.2, width=12.0, height=0.5),
                        font=FontSpec(name="DM Sans", size_pt=24.0, bold=True),
                    ),
                    DataSlot(
                        name="body",
                        slot_type=SlotType.TEXT,
                        data_key="test.body",
                        position=Position(left=0.3, top=1.0, width=12.0, height=5.0),
                        font=FontSpec(name="DM Sans", size_pt=14.0),
                    ),
                ],
            ),
        ],
    )


@pytest.fixture
def full_schema():
    return build_monthly_report_schema()


def _build(schema, payload):
    """Build PPTX bytes from schema+payload."""
    return PPTXBuilder(schema).build(payload)


# ---------------------------------------------------------------------------
# Helper function tests
# ---------------------------------------------------------------------------

class TestIsMissing:
    def test_none(self):
        assert _is_missing(None) is True

    def test_nan(self):
        assert _is_missing(float("nan")) is True

    def test_zero(self):
        assert _is_missing(0) is False

    def test_string(self):
        assert _is_missing("hello") is False

    def test_empty_list(self):
        assert _is_missing([]) is False

    def test_number(self):
        assert _is_missing(42.5) is False


# ---------------------------------------------------------------------------
# Issue and QAResult tests
# ---------------------------------------------------------------------------

class TestIssue:
    def test_str_format(self):
        issue = Issue(
            severity="error",
            slide_index=3,
            slide_name="exec",
            slot_name="table",
            category="table_row_count",
            message="Expected 5 rows, got 3",
        )
        s = str(issue)
        assert "[ERROR]" in s
        assert "slide 3" in s
        assert "exec" in s
        assert "table" in s
        assert "Expected 5 rows" in s

    def test_str_no_slot(self):
        issue = Issue(
            severity="warning",
            slide_index=-1,
            slide_name="",
            slot_name="",
            category="slide_count",
            message="Wrong count",
        )
        s = str(issue)
        assert "[WARNING]" in s
        assert "slide -1" in s


class TestQAResult:
    def test_empty_result_passes(self):
        result = QAResult()
        assert result.passed is True
        assert result.error_count == 0
        assert result.warning_count == 0

    def test_warnings_still_pass(self):
        result = QAResult(issues=[
            Issue("warning", -1, "", "", "test", "A warning"),
        ])
        assert result.passed is True
        assert result.warning_count == 1

    def test_errors_fail(self):
        result = QAResult(issues=[
            Issue("error", 0, "slide", "", "test", "An error"),
        ])
        assert result.passed is False
        assert result.error_count == 1

    def test_mixed_issues(self):
        result = QAResult(issues=[
            Issue("error", 0, "", "", "a", "err"),
            Issue("warning", 1, "", "", "b", "warn"),
            Issue("error", 2, "", "", "c", "err2"),
        ])
        assert result.passed is False
        assert result.error_count == 2
        assert result.warning_count == 1

    def test_summary_pass(self):
        result = QAResult()
        assert "PASS" in result.summary()

    def test_summary_fail(self):
        result = QAResult(issues=[
            Issue("error", 0, "", "", "a", "err"),
        ])
        assert "FAIL" in result.summary()

    def test_report(self):
        result = QAResult(issues=[
            Issue("error", 0, "s", "sl", "a", "message"),
        ])
        report = result.report()
        assert "FAIL" in report
        assert "message" in report


# ---------------------------------------------------------------------------
# Slide count validation
# ---------------------------------------------------------------------------

class TestSlideCount:
    def test_correct_slide_count(self, minimal_schema):
        pptx_bytes = _build(minimal_schema, {})
        result = QAValidator(minimal_schema).validate(pptx_bytes, {})
        slide_count_errors = [
            i for i in result.issues if i.category == "slide_count"
        ]
        assert len(slide_count_errors) == 0

    def test_wrong_slide_count_detected(self, design):
        """Build with 1-slide schema but validate against 2-slide schema."""
        one_slide = TemplateSchema(
            name="One", report_type="monthly",
            width_inches=13.333, height_inches=7.5,
            design=design,
            slides=[
                SlideSchema(index=0, name="s1", title="S1",
                            slide_type=SlideType.DATA, data_source="test"),
            ],
        )
        two_slide = TemplateSchema(
            name="Two", report_type="monthly",
            width_inches=13.333, height_inches=7.5,
            design=design,
            slides=[
                SlideSchema(index=0, name="s1", title="S1",
                            slide_type=SlideType.DATA, data_source="test"),
                SlideSchema(index=1, name="s2", title="S2",
                            slide_type=SlideType.DATA, data_source="test"),
            ],
        )
        pptx_bytes = _build(one_slide, {})
        result = QAValidator(two_slide).validate(pptx_bytes, {})
        errors = [i for i in result.errors if i.category == "slide_count"]
        assert len(errors) == 1
        assert "Expected 2" in errors[0].message
        assert "got 1" in errors[0].message

    def test_full_schema_slide_count(self, full_schema):
        pptx_bytes = _build(full_schema, {})
        result = QAValidator(full_schema).validate(pptx_bytes, {})
        slide_count_errors = [
            i for i in result.errors if i.category == "slide_count"
        ]
        assert len(slide_count_errors) == 0


# ---------------------------------------------------------------------------
# Dimension validation
# ---------------------------------------------------------------------------

class TestDimensions:
    def test_correct_dimensions(self, minimal_schema):
        pptx_bytes = _build(minimal_schema, {})
        result = QAValidator(minimal_schema).validate(pptx_bytes, {})
        dim_errors = [
            i for i in result.errors if i.category == "dimensions"
        ]
        assert len(dim_errors) == 0

    def test_wrong_dimensions_detected(self, design):
        """Build with standard dims but validate against QBR dims."""
        standard = TemplateSchema(
            name="Std", report_type="monthly",
            width_inches=13.333, height_inches=7.5,
            design=design,
            slides=[
                SlideSchema(index=0, name="s1", title="S1",
                            slide_type=SlideType.DATA, data_source="test"),
            ],
        )
        qbr_dims = TemplateSchema(
            name="QBR", report_type="qbr",
            width_inches=21.986, height_inches=12.368,
            design=design,
            slides=[
                SlideSchema(index=0, name="s1", title="S1",
                            slide_type=SlideType.DATA, data_source="test"),
            ],
        )
        pptx_bytes = _build(standard, {})
        result = QAValidator(qbr_dims).validate(pptx_bytes, {})
        dim_errors = [
            i for i in result.errors if i.category == "dimensions"
        ]
        assert len(dim_errors) == 2  # width + height


# ---------------------------------------------------------------------------
# Payload coverage validation
# ---------------------------------------------------------------------------

class TestPayloadCoverage:
    def test_full_payload_no_warnings(self, kpi_schema):
        payload = {"test.revenue": 1000, "test.revenue_var": 5.0}
        pptx_bytes = _build(kpi_schema, payload)
        result = QAValidator(kpi_schema).validate(pptx_bytes, payload)
        missing = [
            i for i in result.issues if i.category == "payload_missing"
        ]
        assert len(missing) == 0

    def test_missing_key_warns(self, kpi_schema):
        payload = {"test.revenue": 1000}  # Missing variance key
        pptx_bytes = _build(kpi_schema, payload)
        result = QAValidator(kpi_schema).validate(pptx_bytes, payload)
        missing = [
            i for i in result.issues if i.category == "payload_missing"
        ]
        assert len(missing) == 1
        assert "test.revenue_var" in missing[0].message

    def test_empty_payload_warns_all(self, kpi_schema):
        payload = {}
        pptx_bytes = _build(kpi_schema, payload)
        result = QAValidator(kpi_schema).validate(pptx_bytes, payload)
        missing = [
            i for i in result.issues if i.category == "payload_missing"
        ]
        assert len(missing) == 2  # revenue + variance_key

    def test_table_payload_keys(self, table_schema):
        payload = {"test.rows": [{"channel": "X", "revenue": 100, "vs_target": 1.0}]}
        pptx_bytes = _build(table_schema, payload)
        result = QAValidator(table_schema).validate(pptx_bytes, payload)
        missing = [
            i for i in result.issues if i.category == "payload_missing"
        ]
        # test.table is the data_key (not in payload), test.rows is present
        table_key_missing = [m for m in missing if "test.table" in m.message]
        rows_key_missing = [m for m in missing if "test.rows" in m.message]
        assert len(rows_key_missing) == 0

    def test_chart_series_keys_tracked(self, chart_schema):
        payload = {}
        pptx_bytes = _build(chart_schema, payload)
        result = QAValidator(chart_schema).validate(pptx_bytes, payload)
        missing = [
            i for i in result.issues if i.category == "payload_missing"
        ]
        missing_keys = {m.message.split("'")[1] for m in missing}
        assert "test.dates" in missing_keys
        assert "test.revenue_series" in missing_keys
        assert "test.target_series" in missing_keys


# ---------------------------------------------------------------------------
# Payload type validation
# ---------------------------------------------------------------------------

class TestPayloadTypes:
    def test_table_rows_must_be_list(self, table_schema):
        payload = {"test.rows": "not a list"}
        result = QAValidator(table_schema).validate_payload(payload)
        type_errors = [
            i for i in result.errors if i.category == "type_error"
        ]
        assert len(type_errors) == 1
        assert "list" in type_errors[0].message

    def test_table_rows_list_is_valid(self, table_schema):
        payload = {"test.rows": [{"channel": "X", "revenue": 100, "vs_target": 0}]}
        result = QAValidator(table_schema).validate_payload(payload)
        type_errors = [
            i for i in result.errors if i.category == "type_error"
        ]
        assert len(type_errors) == 0

    def test_table_column_key_missing_warns(self, table_schema):
        payload = {"test.rows": [{"channel": "X"}]}  # Missing revenue, vs_target
        result = QAValidator(table_schema).validate_payload(payload)
        col_warns = [
            i for i in result.warnings if i.category == "column_key_missing"
        ]
        assert len(col_warns) == 2  # revenue + vs_target

    def test_chart_series_must_be_list(self, chart_schema):
        payload = {
            "test.dates": ["1/1", "1/2"],
            "test.revenue_series": "not a list",
            "test.target_series": [1, 2],
        }
        result = QAValidator(chart_schema).validate_payload(payload)
        type_errors = [
            i for i in result.errors if i.category == "type_error"
        ]
        assert len(type_errors) == 1
        assert "revenue_series" in type_errors[0].message

    def test_chart_series_length_mismatch(self, chart_schema):
        payload = {
            "test.dates": ["1/1", "1/2", "1/3"],
            "test.revenue_series": [100, 200],  # 2 values, 3 categories
            "test.target_series": [150, 150, 150],
        }
        result = QAValidator(chart_schema).validate_payload(payload)
        length_errors = [
            i for i in result.errors
            if i.category == "series_length_mismatch"
        ]
        assert len(length_errors) == 1

    def test_doughnut_series_scalars_ok(self, doughnut_schema):
        payload = {"test.achieved": 75.0, "test.remaining": 25.0}
        result = QAValidator(doughnut_schema).validate_payload(payload)
        type_errors = [
            i for i in result.errors if i.category == "type_error"
        ]
        assert len(type_errors) == 0

    def test_kpi_value_type(self, kpi_schema):
        payload = {"test.revenue": [1, 2, 3]}  # Should be numeric
        result = QAValidator(kpi_schema).validate_payload(payload)
        type_errors = [
            i for i in result.errors if i.category == "type_error"
        ]
        assert len(type_errors) == 1


# ---------------------------------------------------------------------------
# KPI slot validation
# ---------------------------------------------------------------------------

class TestKPIValidation:
    def test_kpi_value_present(self, kpi_schema):
        payload = {"test.revenue": 209200, "test.revenue_var": 5.2}
        pptx_bytes = _build(kpi_schema, payload)
        result = QAValidator(kpi_schema).validate(pptx_bytes, payload)
        kpi_errors = [
            i for i in result.errors if i.category == "kpi_value_missing"
        ]
        assert len(kpi_errors) == 0

    def test_kpi_formatted_value_on_slide(self, kpi_schema):
        payload = {"test.revenue": 1234567, "test.revenue_var": 0}
        pptx_bytes = _build(kpi_schema, payload)
        result = QAValidator(kpi_schema).validate(pptx_bytes, payload)
        kpi_errors = [
            i for i in result.errors if i.category == "kpi_value_missing"
        ]
        assert len(kpi_errors) == 0

    def test_kpi_missing_shows_na(self, kpi_schema):
        payload = {}
        pptx_bytes = _build(kpi_schema, payload)
        result = QAValidator(kpi_schema).validate(pptx_bytes, payload)
        # N/A should be rendered, so no missing_na warning
        na_warns = [
            i for i in result.warnings if i.category == "kpi_missing_na"
        ]
        assert len(na_warns) == 0

    def test_kpi_label_present(self, kpi_schema):
        payload = {"test.revenue": 100000, "test.revenue_var": 0}
        pptx_bytes = _build(kpi_schema, payload)
        result = QAValidator(kpi_schema).validate(pptx_bytes, payload)
        label_warns = [
            i for i in result.warnings if i.category == "kpi_label_missing"
        ]
        assert len(label_warns) == 0

    def test_kpi_positive_variance_color(self, kpi_schema):
        payload = {"test.revenue": 100000, "test.revenue_var": 5.2}
        pptx_bytes = _build(kpi_schema, payload)
        result = QAValidator(kpi_schema).validate(pptx_bytes, payload)
        color_errors = [
            i for i in result.errors if i.category == "variance_color"
        ]
        assert len(color_errors) == 0

    def test_kpi_negative_variance_color(self, kpi_schema):
        payload = {"test.revenue": 100000, "test.revenue_var": -3.1}
        pptx_bytes = _build(kpi_schema, payload)
        result = QAValidator(kpi_schema).validate(pptx_bytes, payload)
        color_errors = [
            i for i in result.errors if i.category == "variance_color"
        ]
        assert len(color_errors) == 0

    def test_kpi_zero_variance_color(self, kpi_schema):
        payload = {"test.revenue": 100000, "test.revenue_var": 0.0}
        pptx_bytes = _build(kpi_schema, payload)
        result = QAValidator(kpi_schema).validate(pptx_bytes, payload)
        color_errors = [
            i for i in result.errors if i.category == "variance_color"
        ]
        assert len(color_errors) == 0


# ---------------------------------------------------------------------------
# Table slot validation
# ---------------------------------------------------------------------------

class TestTableValidation:
    def test_table_row_count_matches(self, table_schema):
        payload = {
            "test.rows": [
                {"channel": "DIRECT", "revenue": 45000, "vs_target": 3.2},
                {"channel": "PPC", "revenue": 32000, "vs_target": -1.5},
            ],
        }
        pptx_bytes = _build(table_schema, payload)
        result = QAValidator(table_schema).validate(pptx_bytes, payload)
        row_errors = [
            i for i in result.errors if i.category == "table_row_count"
        ]
        assert len(row_errors) == 0

    def test_table_column_count_matches(self, table_schema):
        payload = {
            "test.rows": [
                {"channel": "X", "revenue": 100, "vs_target": 0},
            ],
        }
        pptx_bytes = _build(table_schema, payload)
        result = QAValidator(table_schema).validate(pptx_bytes, payload)
        col_errors = [
            i for i in result.errors if i.category == "table_column_count"
        ]
        assert len(col_errors) == 0

    def test_table_headers_correct(self, table_schema):
        payload = {
            "test.rows": [
                {"channel": "X", "revenue": 100, "vs_target": 0},
            ],
        }
        pptx_bytes = _build(table_schema, payload)
        result = QAValidator(table_schema).validate(pptx_bytes, payload)
        header_errors = [
            i for i in result.errors if i.category == "table_header"
        ]
        assert len(header_errors) == 0

    def test_table_cell_formatting(self, table_schema):
        payload = {
            "test.rows": [
                {"channel": "DIRECT", "revenue": 45000, "vs_target": 3.2},
            ],
        }
        pptx_bytes = _build(table_schema, payload)
        result = QAValidator(table_schema).validate(pptx_bytes, payload)
        format_errors = [
            i for i in result.errors if i.category == "table_cell_format"
        ]
        assert len(format_errors) == 0

    def test_table_variance_coloring(self, table_schema):
        payload = {
            "test.rows": [
                {"channel": "DIRECT", "revenue": 50000, "vs_target": 5.0},
                {"channel": "PPC", "revenue": 30000, "vs_target": -2.5},
            ],
        }
        pptx_bytes = _build(table_schema, payload)
        result = QAValidator(table_schema).validate(pptx_bytes, payload)
        color_errors = [
            i for i in result.errors
            if i.category == "table_variance_color"
        ]
        assert len(color_errors) == 0

    def test_table_empty_data_no_crash(self, table_schema):
        payload = {}
        pptx_bytes = _build(table_schema, payload)
        result = QAValidator(table_schema).validate(pptx_bytes, payload)
        # Should not error on missing table (no data)
        table_missing = [
            i for i in result.errors if i.category == "table_missing"
        ]
        assert len(table_missing) == 0

    def test_table_multiple_rows(self, table_schema):
        rows = [
            {"channel": f"CH{i}", "revenue": 1000 * i, "vs_target": i * 0.5}
            for i in range(1, 11)
        ]
        payload = {"test.rows": rows}
        pptx_bytes = _build(table_schema, payload)
        result = QAValidator(table_schema).validate(pptx_bytes, payload)
        row_errors = [
            i for i in result.errors if i.category == "table_row_count"
        ]
        assert len(row_errors) == 0

    def test_table_missing_cell_value(self, table_schema):
        payload = {
            "test.rows": [
                {"channel": "DIRECT", "revenue": None, "vs_target": None},
            ],
        }
        pptx_bytes = _build(table_schema, payload)
        result = QAValidator(table_schema).validate(pptx_bytes, payload)
        # N/A should be rendered for missing values — no format error
        format_errors = [
            i for i in result.errors if i.category == "table_cell_format"
        ]
        assert len(format_errors) == 0


# ---------------------------------------------------------------------------
# Chart slot validation
# ---------------------------------------------------------------------------

class TestChartValidation:
    def test_column_chart_type_correct(self, chart_schema):
        payload = {
            "test.dates": ["1/1", "1/2", "1/3"],
            "test.revenue_series": [10000, 20000, 15000],
            "test.target_series": [15000, 15000, 15000],
        }
        pptx_bytes = _build(chart_schema, payload)
        result = QAValidator(chart_schema).validate(pptx_bytes, payload)
        type_errors = [
            i for i in result.errors if i.category == "chart_type"
        ]
        assert len(type_errors) == 0

    def test_chart_series_count_correct(self, chart_schema):
        payload = {
            "test.dates": ["1/1", "1/2"],
            "test.revenue_series": [10000, 20000],
            "test.target_series": [15000, 15000],
        }
        pptx_bytes = _build(chart_schema, payload)
        result = QAValidator(chart_schema).validate(pptx_bytes, payload)
        series_warns = [
            i for i in result.warnings
            if i.category == "chart_series_count"
        ]
        assert len(series_warns) == 0

    def test_chart_data_length_mismatch(self, chart_schema):
        payload = {
            "test.dates": ["1/1", "1/2", "1/3"],
            "test.revenue_series": [10000, 20000],  # Mismatch!
            "test.target_series": [15000, 15000, 15000],
        }
        pptx_bytes = _build(chart_schema, payload)
        result = QAValidator(chart_schema).validate(pptx_bytes, payload)
        length_errors = [
            i for i in result.errors if i.category == "chart_data_length"
        ]
        assert len(length_errors) == 1

    def test_doughnut_chart_renders(self, doughnut_schema):
        payload = {"test.achieved": 75.0, "test.remaining": 25.0}
        pptx_bytes = _build(doughnut_schema, payload)
        result = QAValidator(doughnut_schema).validate(pptx_bytes, payload)
        type_errors = [
            i for i in result.errors if i.category == "chart_type"
        ]
        assert len(type_errors) == 0

    def test_chart_missing_data_no_crash(self, chart_schema):
        payload = {}
        pptx_bytes = _build(chart_schema, payload)
        result = QAValidator(chart_schema).validate(pptx_bytes, payload)
        # No chart_missing error since no data was supplied
        chart_missing = [
            i for i in result.errors if i.category == "chart_missing"
        ]
        assert len(chart_missing) == 0


# ---------------------------------------------------------------------------
# Section divider validation
# ---------------------------------------------------------------------------

class TestDividerValidation:
    def test_divider_background_correct(self, divider_schema):
        payload = {"divider.title": "eComm Performance"}
        pptx_bytes = _build(divider_schema, payload)
        result = QAValidator(divider_schema).validate(pptx_bytes, payload)
        bg_errors = [
            i for i in result.errors if i.category == "divider_background"
        ]
        assert len(bg_errors) == 0

    def test_divider_text_present(self, divider_schema):
        payload = {"divider.title": "eComm Performance"}
        pptx_bytes = _build(divider_schema, payload)
        result = QAValidator(divider_schema).validate(pptx_bytes, payload)
        text_warns = [
            i for i in result.warnings if i.category == "text_content"
        ]
        assert len(text_warns) == 0


# ---------------------------------------------------------------------------
# Text slot validation
# ---------------------------------------------------------------------------

class TestTextValidation:
    def test_text_present_on_slide(self, text_schema):
        payload = {
            "test.title": "Executive Summary",
            "test.body": "Revenue increased by 5%.",
        }
        pptx_bytes = _build(text_schema, payload)
        result = QAValidator(text_schema).validate(pptx_bytes, payload)
        text_warns = [
            i for i in result.warnings if i.category == "text_content"
        ]
        assert len(text_warns) == 0

    def test_text_list_items_present(self, text_schema):
        payload = {
            "test.title": "TOC",
            "test.body": ["Item 1", "Item 2", "Item 3"],
        }
        pptx_bytes = _build(text_schema, payload)
        result = QAValidator(text_schema).validate(pptx_bytes, payload)
        text_warns = [
            i for i in result.warnings if i.category == "text_content"
        ]
        assert len(text_warns) == 0

    def test_missing_text_no_error(self, text_schema):
        payload = {}
        pptx_bytes = _build(text_schema, payload)
        result = QAValidator(text_schema).validate(pptx_bytes, payload)
        text_warns = [
            i for i in result.warnings if i.category == "text_content"
        ]
        assert len(text_warns) == 0  # Missing data = nothing to validate


# ---------------------------------------------------------------------------
# Convenience function test
# ---------------------------------------------------------------------------

class TestConvenience:
    def test_validate_presentation_function(self, minimal_schema):
        pptx_bytes = _build(minimal_schema, {})
        result = validate_presentation(minimal_schema, pptx_bytes, {})
        assert isinstance(result, QAResult)
        assert result.passed is True


# ---------------------------------------------------------------------------
# Payload-only validation
# ---------------------------------------------------------------------------

class TestValidatePayload:
    def test_valid_payload(self, kpi_schema):
        payload = {"test.revenue": 100000, "test.revenue_var": 5.0}
        result = QAValidator(kpi_schema).validate_payload(payload)
        assert len(result.errors) == 0

    def test_invalid_table_type(self, table_schema):
        payload = {"test.rows": "string"}
        result = QAValidator(table_schema).validate_payload(payload)
        assert len(result.errors) > 0

    def test_missing_column_keys(self, table_schema):
        payload = {"test.rows": [{"channel": "X"}]}
        result = QAValidator(table_schema).validate_payload(payload)
        col_warns = [
            i for i in result.warnings if i.category == "column_key_missing"
        ]
        assert len(col_warns) == 2


# ---------------------------------------------------------------------------
# Full 14-slide integration tests
# ---------------------------------------------------------------------------

class TestFullIntegration:
    def _sample_payload(self):
        """Representative payload covering all 14 slides."""
        return {
            "cover.report_title": "No7 US Monthly eComm Report",
            "cover.report_period": "January 2026 Overview",
            "cover.total_revenue": 1234567,
            "cover.total_orders": 12345,
            "cover.aov": 100.0,
            "cover.new_customers": 4500,
            "cover.cvr": 3.6,
            "cover.cos": 12.5,
            "cover.revenue_vs_target": 5.2,
            "cover.orders_vs_target": 3.1,
            "cover.aov_vs_target": -1.2,
            "cover.nc_vs_target": 8.0,
            "cover.cvr_vs_target": 0.5,
            "cover.cos_vs_target": -0.3,
            "toc.items": [
                "eComm Performance Overview",
                "Daily Performance",
                "Promotion Performance",
            ],
            "divider.ecomm_title": "eComm Performance",
            "divider.channels_title": "Channel Deep Dives",
            "divider.outlook_title": "Outlook",
            "exec.title": "Executive Summary",
            "exec.performance_rows": [
                {
                    "channel": "Total",
                    "revenue": 1234567,
                    "revenue_vs_target": 5.2,
                    "revenue_vs_ly": 12.3,
                    "orders": 12345,
                    "sessions": 345678,
                    "cvr": 3.6,
                    "aov": 100.0,
                    "cos": 12.5,
                    "new_customers": 4500,
                },
            ],
            "exec.narrative": "Strong month.",
            "daily.title": "Daily Performance",
            "daily.dates": ["1/1", "1/2", "1/3"],
            "daily.revenue_actual": [40000, 45000, 38000],
            "daily.revenue_target": [42000, 42000, 42000],
            "daily.revenue_ly": [35000, 38000, 32000],
            "daily.campaign_rows": [
                {"date": "1/1", "activity": "New Year Sale"},
            ],
            "daily.revenue_achieved_pct": 75.0,
            "daily.revenue_remaining_pct": 25.0,
            "promo.title": "Promotion Performance",
            "promo.rows": [
                {
                    "promotion_name": "New Year Sale",
                    "channel": "All",
                    "redemptions": 5000,
                    "redemptions_vs_ly": 12.5,
                    "revenue": 250000,
                    "revenue_vs_ly": 8.3,
                    "discount_amount": 45000,
                },
            ],
            "product.title": "Product Performance",
            "product.rows": [
                {
                    "product_name": "No7 Serum",
                    "units": 3500,
                    "units_vs_ly": 15.2,
                    "revenue": 175000,
                    "revenue_vs_ly": 18.1,
                    "aov": 50.0,
                    "avg_selling_price": 50.0,
                    "discount_pct": 5.0,
                    "new_customers": 800,
                },
            ],
            "crm.title": "CRM Performance",
            "crm.emails_sent": 250000,
            "crm.emails_sent_vs_ly": 5.0,
            "crm.open_rate": 22.5,
            "crm.open_rate_vs_ly": 1.2,
            "crm.ctr": 3.8,
            "crm.ctr_vs_ly": -0.5,
            "crm.revenue": 180000,
            "crm.revenue_vs_ly": 12.0,
            "crm.cvr": 4.2,
            "crm.cvr_vs_ly": 0.3,
            "crm.aov": 95.0,
            "crm.aov_vs_ly": -2.1,
            "crm.detail_rows": [
                {
                    "campaign_type": "Manual",
                    "emails_sent": 150000,
                    "open_rate": 25.0,
                    "ctr": 4.2,
                    "sessions": 6300,
                    "orders": 252,
                    "cvr": 4.0,
                    "revenue": 120000,
                    "aov": 476.19,
                    "revenue_vs_ly": 10.0,
                },
            ],
            "affiliate.title": "Affiliate Performance",
            "affiliate.revenue": 95000,
            "affiliate.revenue_vs_ly": 7.5,
            "affiliate.cos": 8.0,
            "affiliate.cos_vs_ly": -1.0,
            "affiliate.roas": 12.5,
            "affiliate.roas_vs_ly": 1.2,
            "affiliate.orders": 950,
            "affiliate.orders_vs_ly": 5.0,
            "affiliate.cvr": 2.8,
            "affiliate.cvr_vs_ly": 0.2,
            "affiliate.publisher_rows": [
                {
                    "publisher_name": "Publisher A",
                    "revenue": 30000,
                    "revenue_vs_ly": 10.0,
                    "commission": 2400,
                    "cos": 8.0,
                    "orders": 300,
                    "cvr": 3.0,
                    "sessions": 10000,
                    "aov": 100.0,
                },
            ],
            "seo.title": "SEO Performance",
            "seo.revenue": 120000,
            "seo.revenue_vs_ly": 6.0,
            "seo.sessions": 80000,
            "seo.sessions_vs_ly": 4.5,
            "seo.cvr": 3.0,
            "seo.cvr_vs_ly": 0.1,
            "seo.orders": 2400,
            "seo.orders_vs_ly": 6.5,
            "seo.aov": 50.0,
            "seo.aov_vs_ly": -0.5,
            "seo.narrative": "Organic traffic grew steadily.",
            "upcoming.title": "Upcoming Promotions",
            "upcoming.rows": [],
            "next_steps.title": "Next Steps",
            "next_steps.items": "Review Feb targets",
        }

    def test_full_14_slide_passes(self, full_schema):
        payload = self._sample_payload()
        pptx_bytes = _build(full_schema, payload)
        result = QAValidator(full_schema).validate(pptx_bytes, payload)
        # No slide-count or dimension errors
        structural = [
            i for i in result.errors
            if i.category in ("slide_count", "dimensions")
        ]
        assert len(structural) == 0

    def test_full_14_slide_empty_payload(self, full_schema):
        pptx_bytes = _build(full_schema, {})
        result = QAValidator(full_schema).validate(pptx_bytes, {})
        # Should have no errors (only warnings for missing data)
        structural = [
            i for i in result.errors
            if i.category in ("slide_count", "dimensions")
        ]
        assert len(structural) == 0

    def test_full_14_slide_count(self, full_schema):
        payload = self._sample_payload()
        pptx_bytes = _build(full_schema, payload)
        result = QAValidator(full_schema).validate(pptx_bytes, payload)
        count_errors = [
            i for i in result.errors if i.category == "slide_count"
        ]
        assert len(count_errors) == 0

    def test_full_divider_backgrounds(self, full_schema):
        payload = self._sample_payload()
        pptx_bytes = _build(full_schema, payload)
        result = QAValidator(full_schema).validate(pptx_bytes, payload)
        bg_errors = [
            i for i in result.errors if i.category == "divider_background"
        ]
        assert len(bg_errors) == 0

    def test_full_exec_table(self, full_schema):
        payload = self._sample_payload()
        pptx_bytes = _build(full_schema, payload)
        result = QAValidator(full_schema).validate(pptx_bytes, payload)
        table_errors = [
            i for i in result.errors
            if i.category in ("table_row_count", "table_column_count",
                              "table_header", "table_missing")
            and "exec" in i.slide_name
        ]
        assert len(table_errors) == 0

    def test_full_cover_kpis(self, full_schema):
        payload = self._sample_payload()
        pptx_bytes = _build(full_schema, payload)
        result = QAValidator(full_schema).validate(pptx_bytes, payload)
        kpi_errors = [
            i for i in result.errors
            if i.category == "kpi_value_missing"
            and "cover" in i.slide_name
        ]
        assert len(kpi_errors) == 0

    def test_full_chart_validation(self, full_schema):
        payload = self._sample_payload()
        pptx_bytes = _build(full_schema, payload)
        result = QAValidator(full_schema).validate(pptx_bytes, payload)
        chart_type_errors = [
            i for i in result.errors if i.category == "chart_type"
        ]
        assert len(chart_type_errors) == 0

    def test_report_output(self, full_schema):
        payload = self._sample_payload()
        pptx_bytes = _build(full_schema, payload)
        result = QAValidator(full_schema).validate(pptx_bytes, payload)
        report = result.report()
        assert "QA" in report
        assert "error" in report.lower() or "warning" in report.lower() or "PASS" in report


# ---------------------------------------------------------------------------
# Edge cases
# ---------------------------------------------------------------------------

class TestEdgeCases:
    def test_nan_values_in_payload(self, kpi_schema):
        payload = {"test.revenue": float("nan"), "test.revenue_var": float("nan")}
        pptx_bytes = _build(kpi_schema, payload)
        result = QAValidator(kpi_schema).validate(pptx_bytes, payload)
        # NaN treated as missing — N/A should render
        assert result.passed or all(i.severity == "warning" for i in result.issues)

    def test_very_large_values(self, kpi_schema):
        payload = {"test.revenue": 999999999, "test.revenue_var": 999.9}
        pptx_bytes = _build(kpi_schema, payload)
        result = QAValidator(kpi_schema).validate(pptx_bytes, payload)
        kpi_errors = [
            i for i in result.errors if i.category == "kpi_value_missing"
        ]
        assert len(kpi_errors) == 0

    def test_negative_values(self, kpi_schema):
        payload = {"test.revenue": -50000, "test.revenue_var": -15.3}
        pptx_bytes = _build(kpi_schema, payload)
        result = QAValidator(kpi_schema).validate(pptx_bytes, payload)
        kpi_errors = [
            i for i in result.errors if i.category == "kpi_value_missing"
        ]
        assert len(kpi_errors) == 0

    def test_zero_value(self, kpi_schema):
        payload = {"test.revenue": 0, "test.revenue_var": 0}
        pptx_bytes = _build(kpi_schema, payload)
        result = QAValidator(kpi_schema).validate(pptx_bytes, payload)
        kpi_errors = [
            i for i in result.errors if i.category == "kpi_value_missing"
        ]
        assert len(kpi_errors) == 0

    def test_empty_table_rows(self, table_schema):
        payload = {"test.rows": []}
        pptx_bytes = _build(table_schema, payload)
        result = QAValidator(table_schema).validate(pptx_bytes, payload)
        # Empty rows = no table rendered = no table error
        table_missing = [
            i for i in result.errors if i.category == "table_missing"
        ]
        assert len(table_missing) == 0

    def test_single_row_table(self, table_schema):
        payload = {
            "test.rows": [
                {"channel": "ONLY", "revenue": 500, "vs_target": 0.0},
            ],
        }
        pptx_bytes = _build(table_schema, payload)
        result = QAValidator(table_schema).validate(pptx_bytes, payload)
        row_errors = [
            i for i in result.errors if i.category == "table_row_count"
        ]
        assert len(row_errors) == 0

    def test_find_slide_for_key(self, kpi_schema):
        validator = QAValidator(kpi_schema)
        assert validator._find_slide_for_key("test.revenue") == "kpi_slide"
        assert validator._find_slide_for_key("test.revenue_var") == "kpi_slide"
        assert validator._find_slide_for_key("nonexistent") == ""

    def test_chart_no_series_no_crash(self, design):
        schema = TemplateSchema(
            name="Empty Chart",
            report_type="monthly",
            width_inches=13.333,
            height_inches=7.5,
            design=design,
            slides=[
                SlideSchema(
                    index=0,
                    name="empty_chart",
                    title="Empty",
                    slide_type=SlideType.DATA,
                    data_source="test",
                    slots=[
                        DataSlot(
                            name="chart",
                            slot_type=SlotType.CHART,
                            data_key="test.chart",
                            position=Position(left=0, top=0, width=8, height=4),
                            chart_type=ChartType.COLUMN_CLUSTERED,
                            series=[],
                        ),
                    ],
                ),
            ],
        )
        pptx_bytes = _build(schema, {})
        result = QAValidator(schema).validate(pptx_bytes, {})
        # Should not crash on empty series
        assert isinstance(result, QAResult)
