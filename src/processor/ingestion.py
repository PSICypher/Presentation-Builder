"""Data ingestion module for Presentation Builder.

Handles reading, cleaning, and normalizing data from multiple source types:
- Internal Performance Tracker (Excel .xlsx)
- Product Sales Report (CSV, UTF-16 LE, tab-delimited)
- Offer Performance (CSV, UTF-16 LE, tab-delimited)
- CRM/Email Performance (Excel .xlsm)
- Affiliate Publisher Performance (Excel .xlsm)
- Target Phasing (CSV, UTF-8, comma-delimited)
- Trading/Pricing (CSV, UTF-8, comma-delimited)
- Historical Comparison (CSV, UTF-16 LE, tab-delimited)
"""

import re
from pathlib import Path

import pandas as pd


# ---------------------------------------------------------------------------
# Value parsers
# ---------------------------------------------------------------------------

def parse_numeric(value):
    """Parse a numeric value that may contain commas or float artifacts.

    Examples:
        "63,571" -> 63571.0
        "49,156.000000000" -> 49156.0
        "1,138,771" -> 1138771.0
        42 -> 42.0
        NaN -> NaN
    """
    if pd.isna(value):
        return float("nan")
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(",", "")
    if not s:
        return float("nan")
    try:
        f = float(s)
        if f == int(f) and "." not in str(value).replace(",", "").rstrip("0"):
            return float(int(f))
        return f
    except ValueError:
        return float("nan")


def parse_percentage(value):
    """Parse percentage/variance strings into decimal floats.

    Handles multiple formats found across data sources:
        "170.6%"      -> 1.706   (relative change)
        "-3.52 ppts"  -> -0.0352 (absolute point change)
        "-21.6% ↓"    -> -0.216  (with unicode arrow)
        "96.90%"      -> 0.969   (absolute rate)
        "+5.2%"       -> 0.052

    Returns NaN for unparseable values.
    """
    if pd.isna(value):
        return float("nan")
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    # Strip unicode arrows
    s = re.sub(r"[↑↓]", "", s).strip()
    if not s:
        return float("nan")

    # Handle "ppts" format: "-3.52 ppts" -> -0.0352
    ppts_match = re.match(r"^([+-]?\d+\.?\d*)\s*ppts?$", s, re.IGNORECASE)
    if ppts_match:
        return float(ppts_match.group(1)) / 100

    # Handle percentage: "170.6%" -> 1.706, "-21.6%" -> -0.216
    pct_match = re.match(r"^([+-]?\d+\.?\d*)\s*%$", s)
    if pct_match:
        return float(pct_match.group(1)) / 100

    return float("nan")


# ---------------------------------------------------------------------------
# Column cleaning
# ---------------------------------------------------------------------------

def clean_columns(df):
    """Strip whitespace from column names and deduplicate."""
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    return df


# ---------------------------------------------------------------------------
# Encoding detection and CSV reading
# ---------------------------------------------------------------------------

def detect_encoding(path):
    """Detect whether a file is UTF-16 LE (with BOM) or UTF-8.

    Returns (encoding, delimiter) tuple.
    """
    with open(path, "rb") as f:
        raw = f.read(4)
    if raw[:2] == b"\xff\xfe":
        return "utf-16-le", "\t"
    return "utf-8", ","


def read_csv_auto(path):
    """Read a CSV file with automatic encoding and delimiter detection."""
    encoding, sep = detect_encoding(path)
    df = pd.read_csv(path, encoding=encoding, sep=sep)
    return clean_columns(df)


# ---------------------------------------------------------------------------
# Numeric column cleaning
# ---------------------------------------------------------------------------

def clean_numeric_columns(df, columns=None):
    """Apply parse_numeric to specified columns (or all object columns)."""
    if columns is None:
        columns = df.select_dtypes(include=["object", "string"]).columns.tolist()
    for col in columns:
        if col in df.columns:
            df[col] = df[col].apply(parse_numeric)
    return df


def clean_percentage_columns(df, columns):
    """Apply parse_percentage to specified columns."""
    for col in columns:
        if col in df.columns:
            df[col] = df[col].apply(parse_percentage)
    return df


# ---------------------------------------------------------------------------
# Source-specific ingestors
# ---------------------------------------------------------------------------

def ingest_tracker(path):
    """Ingest Internal Performance Tracker (.xlsx).

    Returns a dict of DataFrames keyed by sheet name. Reads the most
    useful sheets for report generation.

    Priority sheets:
        - RAW DATA: Daily channel-level performance (primary)
        - Daily: Day-by-day breakdown with targets
        - MTD Reporting: Month-to-date summary
        - Channel Deep Dive: Aggregated channel YoY variance
        - Data Dump: Raw MTD data with channel breakdowns
        - Manual update- Daily Target: Daily revenue targets
        - Manual update- Channel Targets: Monthly channel targets
    """
    path = Path(path)
    priority_sheets = [
        "RAW DATA",
        "Daily",
        "MTD Reporting",
        "Channel Deep Dive",
        "Data Dump",
        "Manual update- Daily Target",
        "Manual update- Channel Targets",
    ]
    xl = pd.ExcelFile(path, engine="openpyxl")
    available = xl.sheet_names
    result = {}

    for sheet in priority_sheets:
        if sheet not in available:
            continue
        df = xl.parse(sheet)
        df = clean_columns(df)
        result[sheet] = df

    xl.close()
    return result


def ingest_product_sales(path):
    """Ingest Product Sales Report CSV (UTF-16 LE, tab-delimited).

    Cleans column names, parses numeric and percentage columns.
    Returns a single DataFrame with cleaned types.
    """
    df = read_csv_auto(path)

    # Identify column groups
    dimension_cols = [c for c in df.columns if c.startswith("Dimension")]
    analysis_cols = [c for c in df.columns if "(Analysis)" in c]
    comparison_cols = [c for c in df.columns if "(Comparison)" in c]
    vs_comp_cols = [c for c in df.columns if "(vs. Comp)" in c]

    # Parse numeric analysis/comparison columns (comma-formatted strings)
    clean_numeric_columns(df, analysis_cols + comparison_cols)

    # Parse percentage variance columns
    clean_percentage_columns(df, vs_comp_cols)

    return df


def ingest_offer_performance(path):
    """Ingest Offer Performance CSV (UTF-16 LE, tab-delimited).

    Cleans column names, parses numeric and percentage columns.
    """
    df = read_csv_auto(path)

    numeric_cols = ["Redemptions", "Revenue", "Discount Amount"]
    pct_cols = ["% Change Redemptions", "% Change Revenue", "% Change Discount Amount"]

    clean_numeric_columns(df, [c for c in numeric_cols if c in df.columns])
    clean_percentage_columns(df, [c for c in pct_cols if c in df.columns])

    return df


def ingest_crm(path, sheet_name="Custom Dates Dimension Table"):
    """Ingest CRM/Email Performance (.xlsm).

    Reads the specified sheet and normalizes column names.
    Handles the known typo 'Unsibscribe Rate' -> 'Unsubscribe Rate'.
    """
    df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    df = clean_columns(df)

    # Fix known typos
    rename_map = {}
    for col in df.columns:
        if "Unsibscribe" in col:
            rename_map[col] = col.replace("Unsibscribe", "Unsubscribe")
    if rename_map:
        df = df.rename(columns=rename_map)

    return df


def ingest_affiliate(path, sheet_name="Table - Custom Dates"):
    """Ingest Affiliate Publisher Performance (.xlsm).

    Reads the specified sheet, normalizes column names, and fixes
    known typos.
    """
    df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    df = clean_columns(df)

    # Fix known typos
    rename_map = {}
    for col in df.columns:
        if "Sale-Actvie" in col:
            rename_map[col] = col.replace("Sale-Actvie", "Sale-Active")
    if rename_map:
        df = df.rename(columns=rename_map)

    return df


def ingest_targets(path):
    """Ingest Target Phasing CSV (UTF-8, comma-delimited).

    Clean, well-structured data — daily granularity per channel.
    """
    df = read_csv_auto(path)
    # Parse Date column if present
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    return df


def ingest_trading(path):
    """Ingest Trading/Pricing CSV (UTF-8, comma-delimited).

    SKU-level promotional pricing data.
    """
    df = read_csv_auto(path)
    return df


def ingest_historical(path):
    """Ingest Historical Comparison CSV (UTF-16 LE, tab-delimited).

    Prior-year daily data for YoY calculations.
    """
    df = read_csv_auto(path)

    numeric_cols = ["Orders", "New Customers", "Cost", "Revenue"]
    pct_cols = ["CoS", "Phased", "Phased.1"]

    clean_numeric_columns(df, [c for c in numeric_cols if c in df.columns])
    clean_percentage_columns(df, [c for c in pct_cols if c in df.columns])

    return df


# ---------------------------------------------------------------------------
# Source type registry
# ---------------------------------------------------------------------------

SOURCE_TYPES = {
    "tracker": ingest_tracker,
    "product_sales": ingest_product_sales,
    "offer_performance": ingest_offer_performance,
    "crm": ingest_crm,
    "affiliate": ingest_affiliate,
    "targets": ingest_targets,
    "trading": ingest_trading,
    "historical": ingest_historical,
}


def ingest(path, source_type):
    """Ingest a data file by source type.

    Args:
        path: Path to the data file.
        source_type: One of 'tracker', 'product_sales', 'offer_performance',
            'crm', 'affiliate', 'targets', 'trading', 'historical'.

    Returns:
        DataFrame or dict of DataFrames (for tracker).

    Raises:
        ValueError: If source_type is not recognized.
    """
    if source_type not in SOURCE_TYPES:
        raise ValueError(
            f"Unknown source type '{source_type}'. "
            f"Valid types: {', '.join(sorted(SOURCE_TYPES))}"
        )
    return SOURCE_TYPES[source_type](path)
