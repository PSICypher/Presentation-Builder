"""Data processor module for Presentation Builder."""

from .ingestion import (
    ingest,
    ingest_affiliate,
    ingest_crm,
    ingest_historical,
    ingest_offer_performance,
    ingest_product_sales,
    ingest_targets,
    ingest_tracker,
    ingest_trading,
    parse_numeric,
    parse_percentage,
    clean_columns,
    read_csv_auto,
    detect_encoding,
    SOURCE_TYPES,
)
from .transform import (
    DataTransformer,
    ReportContext,
)
