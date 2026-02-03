"""CLI entry point for Presentation Builder.

Orchestrates the full pipeline: schema selection, data ingestion,
mapping, PPTX generation, and QA validation.

Usage::

    # Generate a monthly report
    python -m src.cli generate \\
        --report monthly --month 1 --year 2026 \\
        --tracker data/tracker.xlsx \\
        --targets data/targets.csv \\
        --output output/jan_2026.pptx

    # Generate a QBR report
    python -m src.cli generate \\
        --report qbr --month 3 --year 2026 \\
        --tracker data/tracker.xlsx \\
        --output output/q1_2026.pptx

    # Validate an existing PPTX against its schema
    python -m src.cli validate \\
        --report monthly \\
        --pptx output/jan_2026.pptx

    # Inspect a schema (show slide count, data keys, etc.)
    python -m src.cli inspect --report monthly

    # Use a custom YAML schema
    python -m src.cli generate \\
        --schema output/custom_schema.yaml \\
        --month 1 --year 2026 \\
        --tracker data/tracker.xlsx \\
        --output output/custom.pptx
"""

import argparse
import sys
from pathlib import Path

from src.generator.pptx_builder import PPTXBuilder
from src.processor.ingestion import ingest
from src.processor.mapper import DataMapper
from src.qa.validator import QAValidator
from src.schema.loader import load_schema
from src.schema.monthly_report import build_monthly_report_schema
from src.schema.qbr_report import build_qbr_schema


# ---------------------------------------------------------------------------
# Schema loading
# ---------------------------------------------------------------------------

def _load_schema(args):
    """Load a TemplateSchema from CLI args (--schema or --report)."""
    if hasattr(args, "schema") and args.schema:
        path = Path(args.schema)
        if not path.exists():
            _error(f"Schema file not found: {path}")
        return load_schema(path)

    report_type = getattr(args, "report", "monthly")
    if report_type == "monthly":
        return build_monthly_report_schema()
    elif report_type == "qbr":
        return build_qbr_schema()
    else:
        _error(f"Unknown report type: {report_type!r}. Use 'monthly' or 'qbr'.")


# ---------------------------------------------------------------------------
# Data ingestion
# ---------------------------------------------------------------------------

_SOURCE_FLAGS = {
    "tracker": "tracker",
    "targets": "targets",
    "product_sales": "product_sales",
    "offer_performance": "offer_performance",
    "crm": "crm",
    "affiliate": "affiliate",
    "trading": "trading",
    "historical": "historical",
}


def _ingest_sources(args):
    """Ingest all data sources specified via CLI flags.

    Returns a dict mapping source type to ingested data.
    """
    sources = {}
    for flag, source_type in _SOURCE_FLAGS.items():
        path = getattr(args, flag, None)
        if path is None:
            continue
        p = Path(path)
        if not p.exists():
            _error(f"Data file not found: {p}")
        _info(f"Ingesting {source_type} from {p}")
        sources[source_type] = ingest(p, source_type)

    if not sources:
        _warn("No data sources specified — generating with empty payload")

    return sources


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------

def cmd_generate(args):
    """Generate a PPTX presentation."""
    schema = _load_schema(args)
    _info(f"Schema: {schema.name} ({len(schema.slides)} slides)")

    # Ingest data
    sources = _ingest_sources(args)

    # Map data to payload
    _info("Mapping data to schema keys...")
    mapper = DataMapper(schema, month=args.month, year=args.year)
    result = mapper.map(sources)

    _info(f"Payload coverage: {result.coverage:.0%} "
          f"({int(result.coverage * len(schema.all_data_keys()))}"
          f"/{len(schema.all_data_keys())} keys)")

    if result.warnings:
        for w in result.warnings:
            _warn(w)

    # Generate PPTX
    _info("Building PPTX...")
    builder = PPTXBuilder(schema)
    pptx_bytes = builder.build(result.payload)

    # QA validation
    if not args.skip_qa:
        _info("Running QA validation...")
        validator = QAValidator(schema)
        qa_result = validator.validate(pptx_bytes, result.payload)

        if qa_result.passed:
            _info(qa_result.summary())
        else:
            _warn(qa_result.summary())
            if args.verbose:
                print(qa_result.report(), file=sys.stderr)

            if not args.force:
                _error("QA validation failed. Use --force to write anyway, "
                       "or --skip-qa to skip validation.")
    else:
        _info("QA validation skipped (--skip-qa)")

    # Write output
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(pptx_bytes)
    _info(f"Written: {output} ({len(pptx_bytes):,} bytes)")


def cmd_validate(args):
    """Validate an existing PPTX against its schema."""
    schema = _load_schema(args)
    pptx_path = Path(args.pptx)
    if not pptx_path.exists():
        _error(f"PPTX file not found: {pptx_path}")

    pptx_bytes = pptx_path.read_bytes()
    _info(f"Validating {pptx_path} against {schema.name}")

    validator = QAValidator(schema)

    # Payload validation requires a payload — run structural checks only
    qa_result = validator.validate(pptx_bytes, {})

    print(qa_result.report())
    sys.exit(0 if qa_result.passed else 1)


def cmd_inspect(args):
    """Show schema information."""
    schema = _load_schema(args)

    print(f"Schema:      {schema.name}")
    print(f"Report type: {schema.report_type}")
    print(f"Dimensions:  {schema.width_inches}\" x {schema.height_inches}\"")
    print(f"Slides:      {len(schema.slides)}")
    print()

    all_keys = sorted(schema.all_data_keys())
    print(f"Data keys:   {len(all_keys)}")

    if args.verbose:
        print()
        for slide in schema.slides:
            slot_count = len(slide.slots)
            static = " (static)" if slide.is_static else ""
            print(f"  [{slide.index:2d}] {slide.name}"
                  f" — {slide.slide_type.value}{static}"
                  f" — {slot_count} slot(s)")
            if args.keys:
                for slot in slide.slots:
                    print(f"       {slot.data_key}"
                          f" ({slot.slot_type.value})")

    if args.keys and not args.verbose:
        print()
        for key in all_keys:
            print(f"  {key}")


# ---------------------------------------------------------------------------
# Output helpers
# ---------------------------------------------------------------------------

def _info(msg):
    print(f"  {msg}", file=sys.stderr)


def _warn(msg):
    print(f"  WARNING: {msg}", file=sys.stderr)


def _error(msg):
    print(f"  ERROR: {msg}", file=sys.stderr)
    sys.exit(1)


# ---------------------------------------------------------------------------
# Argument parsing
# ---------------------------------------------------------------------------

def build_parser():
    """Build the argument parser."""
    parser = argparse.ArgumentParser(
        prog="presentation-builder",
        description="Generate eCommerce report presentations from data sources.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # ---- generate ----
    gen = subparsers.add_parser(
        "generate",
        help="Generate a PPTX presentation from data sources.",
    )
    _add_schema_args(gen)
    _add_period_args(gen)
    _add_data_args(gen)
    gen.add_argument(
        "-o", "--output",
        required=True,
        help="Output PPTX file path.",
    )
    gen.add_argument(
        "--skip-qa",
        action="store_true",
        default=False,
        help="Skip QA validation after generation.",
    )
    gen.add_argument(
        "--force",
        action="store_true",
        default=False,
        help="Write output even if QA validation fails.",
    )
    gen.add_argument(
        "-v", "--verbose",
        action="store_true",
        default=False,
        help="Show detailed output (full QA report on failure).",
    )
    gen.set_defaults(func=cmd_generate)

    # ---- validate ----
    val = subparsers.add_parser(
        "validate",
        help="Validate an existing PPTX against its schema.",
    )
    _add_schema_args(val)
    val.add_argument(
        "--pptx",
        required=True,
        help="Path to the PPTX file to validate.",
    )
    val.set_defaults(func=cmd_validate)

    # ---- inspect ----
    insp = subparsers.add_parser(
        "inspect",
        help="Show schema structure and data keys.",
    )
    _add_schema_args(insp)
    insp.add_argument(
        "-v", "--verbose",
        action="store_true",
        default=False,
        help="Show per-slide detail.",
    )
    insp.add_argument(
        "--keys",
        action="store_true",
        default=False,
        help="List all data keys.",
    )
    insp.set_defaults(func=cmd_inspect)

    return parser


def _add_schema_args(parser):
    """Add --report / --schema args to a subparser."""
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--report",
        choices=["monthly", "qbr"],
        default="monthly",
        help="Built-in report type (default: monthly).",
    )
    group.add_argument(
        "--schema",
        help="Path to a custom YAML schema file.",
    )


def _add_period_args(parser):
    """Add --month / --year args to a subparser."""
    import datetime
    now = datetime.date.today()
    parser.add_argument(
        "--month",
        type=int,
        choices=range(1, 13),
        metavar="1-12",
        default=now.month,
        help=f"Report month (default: {now.month}).",
    )
    parser.add_argument(
        "--year",
        type=int,
        default=now.year,
        help=f"Report year (default: {now.year}).",
    )


def _add_data_args(parser):
    """Add data source file path arguments."""
    data = parser.add_argument_group("data sources")
    data.add_argument(
        "--tracker",
        help="Internal Performance Tracker (.xlsx).",
    )
    data.add_argument(
        "--targets",
        help="Target Phasing (.csv).",
    )
    data.add_argument(
        "--product-sales",
        dest="product_sales",
        help="Product Sales Report (.csv, UTF-16 LE).",
    )
    data.add_argument(
        "--offer-performance",
        dest="offer_performance",
        help="Offer Performance (.csv, UTF-16 LE).",
    )
    data.add_argument(
        "--crm",
        help="CRM/Email Performance (.xlsm).",
    )
    data.add_argument(
        "--affiliate",
        help="Affiliate Publisher Performance (.xlsm).",
    )
    data.add_argument(
        "--trading",
        help="Trading/Pricing (.csv).",
    )
    data.add_argument(
        "--historical",
        help="Historical Comparison (.csv, UTF-16 LE).",
    )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main(argv=None):
    """CLI entry point."""
    parser = build_parser()
    args = parser.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    main()
