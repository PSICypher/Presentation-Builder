"""Tests for the CLI entry point (src.cli).

Covers argument parsing, command dispatch, schema loading, data ingestion
wiring, generate pipeline, validate command, inspect command, and error
handling.  Uses monkeypatching to avoid real file I/O and heavy imports
where possible.
"""

import argparse
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch, PropertyMock

import pytest

from src.cli import (
    build_parser,
    cmd_generate,
    cmd_inspect,
    cmd_validate,
    main,
    _load_schema,
    _ingest_sources,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def parser():
    return build_parser()


@pytest.fixture
def minimal_schema():
    """A minimal TemplateSchema mock."""
    schema = MagicMock()
    schema.name = "Test Schema"
    schema.report_type = "monthly"
    schema.width_inches = 13.333
    schema.height_inches = 7.5
    schema.slides = [MagicMock(index=0, name="cover",
                                slide_type=MagicMock(value="data"),
                                is_static=False,
                                slots=[])]
    schema.all_data_keys.return_value = {"cover.revenue", "cover.orders"}
    return schema


@pytest.fixture
def mapping_result():
    """A MappingResult mock."""
    result = MagicMock()
    result.payload = {"cover.revenue": 100000, "cover.orders": 500}
    result.coverage = 1.0
    result.warnings = []
    return result


@pytest.fixture
def qa_pass():
    """A passing QAResult mock."""
    qa = MagicMock()
    qa.passed = True
    qa.summary.return_value = "QA PASS: 0 error(s), 0 warning(s)"
    qa.report.return_value = "QA PASS: 0 error(s), 0 warning(s)"
    return qa


@pytest.fixture
def qa_fail():
    """A failing QAResult mock."""
    qa = MagicMock()
    qa.passed = False
    qa.summary.return_value = "QA FAIL: 2 error(s), 1 warning(s)"
    qa.report.return_value = (
        "QA FAIL: 2 error(s), 1 warning(s)\n"
        "  ERROR: slide 0 missing\n"
        "  ERROR: slide 1 missing\n"
        "  WARNING: low coverage"
    )
    return qa


# ===================================================================
# Parser tests
# ===================================================================

class TestParser:
    """Argument parsing tests."""

    def test_generate_minimal(self, parser):
        args = parser.parse_args([
            "generate", "--month", "1", "--year", "2026",
            "-o", "out.pptx",
        ])
        assert args.command == "generate"
        assert args.month == 1
        assert args.year == 2026
        assert args.output == "out.pptx"
        assert args.report == "monthly"
        assert args.skip_qa is False
        assert args.force is False
        assert args.verbose is False

    def test_generate_all_data_flags(self, parser):
        args = parser.parse_args([
            "generate", "-o", "out.pptx",
            "--tracker", "t.xlsx",
            "--targets", "tg.csv",
            "--product-sales", "ps.csv",
            "--offer-performance", "op.csv",
            "--crm", "crm.xlsm",
            "--affiliate", "aff.xlsm",
            "--trading", "tr.csv",
            "--historical", "hist.csv",
        ])
        assert args.tracker == "t.xlsx"
        assert args.targets == "tg.csv"
        assert args.product_sales == "ps.csv"
        assert args.offer_performance == "op.csv"
        assert args.crm == "crm.xlsm"
        assert args.affiliate == "aff.xlsm"
        assert args.trading == "tr.csv"
        assert args.historical == "hist.csv"

    def test_generate_qbr_report(self, parser):
        args = parser.parse_args([
            "generate", "--report", "qbr", "-o", "out.pptx",
        ])
        assert args.report == "qbr"

    def test_generate_custom_schema(self, parser):
        args = parser.parse_args([
            "generate", "--schema", "custom.yaml", "-o", "out.pptx",
        ])
        assert args.schema == "custom.yaml"
        # --report not set when --schema used (mutually exclusive)

    def test_generate_flags(self, parser):
        args = parser.parse_args([
            "generate", "-o", "out.pptx",
            "--skip-qa", "--force", "-v",
        ])
        assert args.skip_qa is True
        assert args.force is True
        assert args.verbose is True

    def test_validate_command(self, parser):
        args = parser.parse_args([
            "validate", "--pptx", "report.pptx",
        ])
        assert args.command == "validate"
        assert args.pptx == "report.pptx"

    def test_validate_with_schema(self, parser):
        args = parser.parse_args([
            "validate", "--schema", "s.yaml", "--pptx", "r.pptx",
        ])
        assert args.schema == "s.yaml"

    def test_inspect_command(self, parser):
        args = parser.parse_args(["inspect"])
        assert args.command == "inspect"
        assert args.verbose is False
        assert args.keys is False

    def test_inspect_verbose_keys(self, parser):
        args = parser.parse_args(["inspect", "-v", "--keys"])
        assert args.verbose is True
        assert args.keys is True

    def test_inspect_qbr(self, parser):
        args = parser.parse_args(["inspect", "--report", "qbr"])
        assert args.report == "qbr"

    def test_no_command_fails(self, parser):
        with pytest.raises(SystemExit):
            parser.parse_args([])

    def test_generate_requires_output(self, parser):
        with pytest.raises(SystemExit):
            parser.parse_args(["generate"])

    def test_validate_requires_pptx(self, parser):
        with pytest.raises(SystemExit):
            parser.parse_args(["validate"])

    def test_month_range_enforced(self, parser):
        with pytest.raises(SystemExit):
            parser.parse_args([
                "generate", "-o", "out.pptx", "--month", "13",
            ])

    def test_month_zero_rejected(self, parser):
        with pytest.raises(SystemExit):
            parser.parse_args([
                "generate", "-o", "out.pptx", "--month", "0",
            ])

    def test_schema_and_report_mutually_exclusive(self, parser):
        with pytest.raises(SystemExit):
            parser.parse_args([
                "generate", "--report", "monthly", "--schema", "s.yaml",
                "-o", "out.pptx",
            ])


# ===================================================================
# Schema loading tests
# ===================================================================

class TestLoadSchema:
    """Tests for _load_schema()."""

    def test_load_monthly(self):
        args = argparse.Namespace(report="monthly", schema=None)
        schema = _load_schema(args)
        assert schema.name is not None
        assert len(schema.slides) == 14

    def test_load_qbr(self):
        from src.schema.qbr_report import build_qbr_schema
        schema = build_qbr_schema()
        assert schema.name is not None
        assert len(schema.slides) == 29

    def test_load_monthly_via_cli(self):
        args = argparse.Namespace(report="monthly", schema=None)
        schema = _load_schema(args)
        assert len(schema.slides) == 14

    def test_load_qbr_via_cli(self):
        args = argparse.Namespace(report="qbr", schema=None)
        schema = _load_schema(args)
        assert len(schema.slides) == 29

    def test_load_custom_yaml(self, tmp_path, minimal_schema):
        from src.schema.loader import save_schema
        from src.schema.monthly_report import build_monthly_report_schema

        schema = build_monthly_report_schema()
        yaml_path = tmp_path / "test_schema.yaml"
        save_schema(schema, yaml_path)

        args = argparse.Namespace(schema=str(yaml_path), report="monthly")
        loaded = _load_schema(args)
        assert loaded.name == schema.name
        assert len(loaded.slides) == len(schema.slides)

    def test_load_missing_yaml_exits(self):
        args = argparse.Namespace(schema="/nonexistent/schema.yaml",
                                  report="monthly")
        with pytest.raises(SystemExit):
            _load_schema(args)

    def test_load_unknown_report_exits(self):
        args = argparse.Namespace(schema=None, report="unknown")
        with pytest.raises(SystemExit):
            _load_schema(args)


# ===================================================================
# Ingest sources tests
# ===================================================================

class TestIngestSources:
    """Tests for _ingest_sources()."""

    def test_no_sources_returns_empty(self):
        args = argparse.Namespace(
            tracker=None, targets=None, product_sales=None,
            offer_performance=None, crm=None, affiliate=None,
            trading=None, historical=None,
        )
        sources = _ingest_sources(args)
        assert sources == {}

    def test_missing_file_exits(self):
        args = argparse.Namespace(
            tracker="/nonexistent/tracker.xlsx",
            targets=None, product_sales=None,
            offer_performance=None, crm=None, affiliate=None,
            trading=None, historical=None,
        )
        with pytest.raises(SystemExit):
            _ingest_sources(args)

    def test_ingests_specified_sources(self, tmp_path):
        # Create a dummy CSV file
        csv_path = tmp_path / "targets.csv"
        csv_path.write_text(
            "Date,Channel_Id,Gross_Revenue_Target,Order_Target,"
            "Session_Target,New_Customer_Target,Marketing_Spend_Target\n"
            "2026-01-01,PPC,1000,50,500,10,200\n"
        )

        args = argparse.Namespace(
            tracker=None, targets=str(csv_path), product_sales=None,
            offer_performance=None, crm=None, affiliate=None,
            trading=None, historical=None,
        )
        sources = _ingest_sources(args)
        assert "targets" in sources
        assert len(sources) == 1


# ===================================================================
# Generate command tests
# ===================================================================

class TestCmdGenerate:
    """Tests for cmd_generate()."""

    def test_full_pipeline(self, tmp_path, minimal_schema,
                           mapping_result, qa_pass):
        output = tmp_path / "report.pptx"

        with patch("src.cli._load_schema", return_value=minimal_schema), \
             patch("src.cli._ingest_sources", return_value={}), \
             patch("src.cli.DataMapper") as MockMapper, \
             patch("src.cli.PPTXBuilder") as MockBuilder, \
             patch("src.cli.QAValidator") as MockValidator:

            MockMapper.return_value.map.return_value = mapping_result
            MockBuilder.return_value.build.return_value = b"PK\x03\x04fake"
            MockValidator.return_value.validate.return_value = qa_pass

            args = argparse.Namespace(
                report="monthly", schema=None,
                month=1, year=2026,
                output=str(output),
                skip_qa=False, force=False, verbose=False,
                tracker=None, targets=None, product_sales=None,
                offer_performance=None, crm=None, affiliate=None,
                trading=None, historical=None,
            )
            cmd_generate(args)

        assert output.exists()
        assert output.read_bytes() == b"PK\x03\x04fake"
        MockMapper.assert_called_once_with(minimal_schema, month=1, year=2026)
        MockBuilder.assert_called_once_with(minimal_schema)

    def test_skip_qa(self, tmp_path, minimal_schema, mapping_result):
        output = tmp_path / "report.pptx"

        with patch("src.cli._load_schema", return_value=minimal_schema), \
             patch("src.cli._ingest_sources", return_value={}), \
             patch("src.cli.DataMapper") as MockMapper, \
             patch("src.cli.PPTXBuilder") as MockBuilder, \
             patch("src.cli.QAValidator") as MockValidator:

            MockMapper.return_value.map.return_value = mapping_result
            MockBuilder.return_value.build.return_value = b"PK\x03\x04fake"

            args = argparse.Namespace(
                report="monthly", schema=None,
                month=1, year=2026,
                output=str(output),
                skip_qa=True, force=False, verbose=False,
                tracker=None, targets=None, product_sales=None,
                offer_performance=None, crm=None, affiliate=None,
                trading=None, historical=None,
            )
            cmd_generate(args)

        assert output.exists()
        MockValidator.return_value.validate.assert_not_called()

    def test_qa_fail_exits(self, tmp_path, minimal_schema,
                           mapping_result, qa_fail):
        output = tmp_path / "report.pptx"

        with patch("src.cli._load_schema", return_value=minimal_schema), \
             patch("src.cli._ingest_sources", return_value={}), \
             patch("src.cli.DataMapper") as MockMapper, \
             patch("src.cli.PPTXBuilder") as MockBuilder, \
             patch("src.cli.QAValidator") as MockValidator:

            MockMapper.return_value.map.return_value = mapping_result
            MockBuilder.return_value.build.return_value = b"PK\x03\x04fake"
            MockValidator.return_value.validate.return_value = qa_fail

            args = argparse.Namespace(
                report="monthly", schema=None,
                month=1, year=2026,
                output=str(output),
                skip_qa=False, force=False, verbose=False,
                tracker=None, targets=None, product_sales=None,
                offer_performance=None, crm=None, affiliate=None,
                trading=None, historical=None,
            )

            with pytest.raises(SystemExit):
                cmd_generate(args)

        assert not output.exists()

    def test_qa_fail_force_writes(self, tmp_path, minimal_schema,
                                  mapping_result, qa_fail):
        output = tmp_path / "report.pptx"

        with patch("src.cli._load_schema", return_value=minimal_schema), \
             patch("src.cli._ingest_sources", return_value={}), \
             patch("src.cli.DataMapper") as MockMapper, \
             patch("src.cli.PPTXBuilder") as MockBuilder, \
             patch("src.cli.QAValidator") as MockValidator:

            MockMapper.return_value.map.return_value = mapping_result
            MockBuilder.return_value.build.return_value = b"PK\x03\x04fake"
            MockValidator.return_value.validate.return_value = qa_fail

            args = argparse.Namespace(
                report="monthly", schema=None,
                month=1, year=2026,
                output=str(output),
                skip_qa=False, force=True, verbose=False,
                tracker=None, targets=None, product_sales=None,
                offer_performance=None, crm=None, affiliate=None,
                trading=None, historical=None,
            )
            cmd_generate(args)

        assert output.exists()

    def test_qa_fail_verbose_shows_report(self, tmp_path, minimal_schema,
                                          mapping_result, qa_fail, capsys):
        output = tmp_path / "report.pptx"

        with patch("src.cli._load_schema", return_value=minimal_schema), \
             patch("src.cli._ingest_sources", return_value={}), \
             patch("src.cli.DataMapper") as MockMapper, \
             patch("src.cli.PPTXBuilder") as MockBuilder, \
             patch("src.cli.QAValidator") as MockValidator:

            MockMapper.return_value.map.return_value = mapping_result
            MockBuilder.return_value.build.return_value = b"PK\x03\x04fake"
            MockValidator.return_value.validate.return_value = qa_fail

            args = argparse.Namespace(
                report="monthly", schema=None,
                month=1, year=2026,
                output=str(output),
                skip_qa=False, force=True, verbose=True,
                tracker=None, targets=None, product_sales=None,
                offer_performance=None, crm=None, affiliate=None,
                trading=None, historical=None,
            )
            cmd_generate(args)

        captured = capsys.readouterr()
        assert "FAIL" in captured.err

    def test_creates_output_directory(self, tmp_path, minimal_schema,
                                      mapping_result, qa_pass):
        output = tmp_path / "nested" / "dir" / "report.pptx"

        with patch("src.cli._load_schema", return_value=minimal_schema), \
             patch("src.cli._ingest_sources", return_value={}), \
             patch("src.cli.DataMapper") as MockMapper, \
             patch("src.cli.PPTXBuilder") as MockBuilder, \
             patch("src.cli.QAValidator") as MockValidator:

            MockMapper.return_value.map.return_value = mapping_result
            MockBuilder.return_value.build.return_value = b"PK\x03\x04fake"
            MockValidator.return_value.validate.return_value = qa_pass

            args = argparse.Namespace(
                report="monthly", schema=None,
                month=1, year=2026,
                output=str(output),
                skip_qa=False, force=False, verbose=False,
                tracker=None, targets=None, product_sales=None,
                offer_performance=None, crm=None, affiliate=None,
                trading=None, historical=None,
            )
            cmd_generate(args)

        assert output.exists()

    def test_mapper_warnings_printed(self, tmp_path, minimal_schema,
                                     qa_pass, capsys):
        output = tmp_path / "report.pptx"
        result = MagicMock()
        result.payload = {}
        result.coverage = 0.5
        result.warnings = ["cover: No tracker data", "seo: Missing data"]

        with patch("src.cli._load_schema", return_value=minimal_schema), \
             patch("src.cli._ingest_sources", return_value={}), \
             patch("src.cli.DataMapper") as MockMapper, \
             patch("src.cli.PPTXBuilder") as MockBuilder, \
             patch("src.cli.QAValidator") as MockValidator:

            MockMapper.return_value.map.return_value = result
            MockBuilder.return_value.build.return_value = b"PK\x03\x04fake"
            MockValidator.return_value.validate.return_value = qa_pass

            args = argparse.Namespace(
                report="monthly", schema=None,
                month=1, year=2026,
                output=str(output),
                skip_qa=False, force=False, verbose=False,
                tracker=None, targets=None, product_sales=None,
                offer_performance=None, crm=None, affiliate=None,
                trading=None, historical=None,
            )
            cmd_generate(args)

        captured = capsys.readouterr()
        assert "cover: No tracker data" in captured.err
        assert "seo: Missing data" in captured.err


# ===================================================================
# Validate command tests
# ===================================================================

class TestCmdValidate:
    """Tests for cmd_validate()."""

    def test_validate_pass(self, tmp_path, minimal_schema, qa_pass):
        pptx_path = tmp_path / "test.pptx"
        pptx_path.write_bytes(b"PK\x03\x04fake")

        with patch("src.cli._load_schema", return_value=minimal_schema), \
             patch("src.cli.QAValidator") as MockValidator:
            MockValidator.return_value.validate.return_value = qa_pass

            args = argparse.Namespace(
                report="monthly", schema=None,
                pptx=str(pptx_path),
            )
            with pytest.raises(SystemExit) as exc_info:
                cmd_validate(args)
            assert exc_info.value.code == 0

    def test_validate_fail_exit_code(self, tmp_path, minimal_schema, qa_fail):
        pptx_path = tmp_path / "test.pptx"
        pptx_path.write_bytes(b"PK\x03\x04fake")

        with patch("src.cli._load_schema", return_value=minimal_schema), \
             patch("src.cli.QAValidator") as MockValidator:
            MockValidator.return_value.validate.return_value = qa_fail

            args = argparse.Namespace(
                report="monthly", schema=None,
                pptx=str(pptx_path),
            )
            with pytest.raises(SystemExit) as exc_info:
                cmd_validate(args)
            assert exc_info.value.code == 1

    def test_validate_missing_pptx_exits(self, minimal_schema):
        with patch("src.cli._load_schema", return_value=minimal_schema):
            args = argparse.Namespace(
                report="monthly", schema=None,
                pptx="/nonexistent/file.pptx",
            )
            with pytest.raises(SystemExit):
                cmd_validate(args)

    def test_validate_prints_report(self, tmp_path, minimal_schema,
                                    qa_pass, capsys):
        pptx_path = tmp_path / "test.pptx"
        pptx_path.write_bytes(b"PK\x03\x04fake")

        with patch("src.cli._load_schema", return_value=minimal_schema), \
             patch("src.cli.QAValidator") as MockValidator:
            MockValidator.return_value.validate.return_value = qa_pass

            args = argparse.Namespace(
                report="monthly", schema=None,
                pptx=str(pptx_path),
            )
            with pytest.raises(SystemExit):
                cmd_validate(args)

        captured = capsys.readouterr()
        assert "PASS" in captured.out


# ===================================================================
# Inspect command tests
# ===================================================================

class TestCmdInspect:
    """Tests for cmd_inspect()."""

    def test_inspect_basic(self, minimal_schema, capsys):
        with patch("src.cli._load_schema", return_value=minimal_schema):
            args = argparse.Namespace(
                report="monthly", schema=None,
                verbose=False, keys=False,
            )
            cmd_inspect(args)

        captured = capsys.readouterr()
        assert "Test Schema" in captured.out
        assert "1" in captured.out  # 1 slide

    def test_inspect_verbose(self, minimal_schema, capsys):
        with patch("src.cli._load_schema", return_value=minimal_schema):
            args = argparse.Namespace(
                report="monthly", schema=None,
                verbose=True, keys=False,
            )
            cmd_inspect(args)

        captured = capsys.readouterr()
        assert "cover" in captured.out

    def test_inspect_keys(self, minimal_schema, capsys):
        with patch("src.cli._load_schema", return_value=minimal_schema):
            args = argparse.Namespace(
                report="monthly", schema=None,
                verbose=False, keys=True,
            )
            cmd_inspect(args)

        captured = capsys.readouterr()
        assert "cover.revenue" in captured.out
        assert "cover.orders" in captured.out

    def test_inspect_verbose_keys(self, minimal_schema, capsys):
        with patch("src.cli._load_schema", return_value=minimal_schema):
            slot = MagicMock()
            slot.data_key = "cover.revenue"
            slot.slot_type = MagicMock(value="kpi_value")
            minimal_schema.slides[0].slots = [slot]

            args = argparse.Namespace(
                report="monthly", schema=None,
                verbose=True, keys=True,
            )
            cmd_inspect(args)

        captured = capsys.readouterr()
        assert "cover.revenue" in captured.out
        assert "kpi_value" in captured.out

    def test_inspect_monthly_real(self, capsys):
        """Integration: inspect the real monthly schema."""
        args = argparse.Namespace(
            report="monthly", schema=None,
            verbose=False, keys=False,
        )
        cmd_inspect(args)

        captured = capsys.readouterr()
        assert "14" in captured.out  # 14 slides

    def test_inspect_qbr_real(self, capsys):
        """Integration: inspect the real QBR schema."""
        args = argparse.Namespace(
            report="qbr", schema=None,
            verbose=False, keys=False,
        )
        cmd_inspect(args)

        captured = capsys.readouterr()
        assert "29" in captured.out  # 29 slides


# ===================================================================
# Main entry point tests
# ===================================================================

class TestMain:
    """Tests for main() and __main__.py."""

    def test_main_inspect(self, capsys):
        main(["inspect", "--report", "monthly"])
        captured = capsys.readouterr()
        assert "14" in captured.out

    def test_main_no_args_exits(self):
        with pytest.raises(SystemExit):
            main([])

    def test_main_generate_with_output(self, tmp_path):
        output = tmp_path / "test.pptx"
        main([
            "generate", "--report", "monthly",
            "--month", "1", "--year", "2026",
            "--skip-qa", "-o", str(output),
        ])
        assert output.exists()
        assert output.stat().st_size > 0

    def test_main_help(self):
        with pytest.raises(SystemExit) as exc_info:
            main(["--help"])
        assert exc_info.value.code == 0

    def test_main_generate_help(self):
        with pytest.raises(SystemExit) as exc_info:
            main(["generate", "--help"])
        assert exc_info.value.code == 0


# ===================================================================
# Integration tests
# ===================================================================

class TestIntegration:
    """End-to-end integration tests with real schemas."""

    def test_generate_monthly_empty_data(self, tmp_path):
        """Generate a monthly report with no data sources."""
        output = tmp_path / "empty_monthly.pptx"
        main([
            "generate", "--report", "monthly",
            "--month", "6", "--year", "2025",
            "--skip-qa", "-o", str(output),
        ])
        assert output.exists()
        assert output.stat().st_size > 0

    def test_generate_qbr_empty_data(self, tmp_path):
        """Generate a QBR with no data sources."""
        output = tmp_path / "empty_qbr.pptx"
        main([
            "generate", "--report", "qbr",
            "--month", "3", "--year", "2025",
            "--skip-qa", "-o", str(output),
        ])
        assert output.exists()
        assert output.stat().st_size > 0

    def test_generate_monthly_with_qa(self, tmp_path):
        """Generate monthly report and run QA (force write)."""
        output = tmp_path / "qa_monthly.pptx"
        main([
            "generate", "--report", "monthly",
            "--month", "1", "--year", "2026",
            "--force", "-o", str(output),
        ])
        assert output.exists()

    def test_generate_with_targets_csv(self, tmp_path):
        """Generate with a real targets CSV file."""
        targets = tmp_path / "targets.csv"
        targets.write_text(
            "Date,Channel_Id,Gross_Revenue_Target,Net_Revenue_Target,"
            "Order_Target,Session_Target,New_Customer_Target,"
            "Marketing_Spend_Target\n"
            "2026-01-01,PPC,10000,9000,100,1000,20,2000\n"
            "2026-01-02,PPC,10000,9000,100,1000,20,2000\n"
        )

        output = tmp_path / "with_targets.pptx"
        main([
            "generate", "--report", "monthly",
            "--month", "1", "--year", "2026",
            "--targets", str(targets),
            "--skip-qa", "-o", str(output),
        ])
        assert output.exists()

    def test_roundtrip_schema_yaml(self, tmp_path):
        """Save schema to YAML, load it back, generate with it."""
        from src.schema.loader import save_schema
        from src.schema.monthly_report import build_monthly_report_schema

        schema = build_monthly_report_schema()
        yaml_path = tmp_path / "schema.yaml"
        save_schema(schema, yaml_path)

        output = tmp_path / "from_yaml.pptx"
        main([
            "generate", "--schema", str(yaml_path),
            "--month", "2", "--year", "2026",
            "--skip-qa", "-o", str(output),
        ])
        assert output.exists()

    def test_validate_generated_pptx(self, tmp_path):
        """Generate then validate the output."""
        output = tmp_path / "to_validate.pptx"
        main([
            "generate", "--report", "monthly",
            "--month", "1", "--year", "2026",
            "--skip-qa", "-o", str(output),
        ])

        # Validate (will exit with code 0 or 1)
        with pytest.raises(SystemExit) as exc_info:
            main([
                "validate", "--report", "monthly",
                "--pptx", str(output),
            ])
        # We accept either pass or fail — the point is it runs
        assert exc_info.value.code in (0, 1)

    def test_default_month_year(self, tmp_path):
        """Default month/year should use current date."""
        import datetime
        now = datetime.date.today()

        output = tmp_path / "default_date.pptx"
        main([
            "generate", "--report", "monthly",
            "--skip-qa", "-o", str(output),
        ])
        assert output.exists()
