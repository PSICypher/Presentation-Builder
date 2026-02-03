"""Template schema package — typed models for presentation structure.

Provides the contract between the template analyzer, data processor,
and presentation generator:

- models.py: Core dataclasses (TemplateSchema, SlideSchema, DataSlot, etc.)
- design_system.py: Value formatting functions (currency, percentage, etc.)
- monthly_report.py: The 14-slide monthly eComm report schema definition
- loader.py: YAML serialization/deserialization
"""

from .design_system import (
    format_currency,
    format_integer,
    format_number,
    format_percentage,
    format_points_change,
    format_value,
    format_variance_percentage,
    variance_color,
)
from .loader import load_schema, save_schema
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
from .monthly_report import build_monthly_report_schema

__all__ = [
    # Models
    "ChartSeries",
    "ChartType",
    "DataSlot",
    "DesignSystem",
    "FontSpec",
    "FormatRule",
    "FormatType",
    "Position",
    "SlideSchema",
    "SlideType",
    "SlotType",
    "TableColumn",
    "TemplateSchema",
    # Schema builders
    "build_monthly_report_schema",
    # Loader
    "load_schema",
    "save_schema",
    # Formatting
    "format_currency",
    "format_integer",
    "format_number",
    "format_percentage",
    "format_points_change",
    "format_value",
    "format_variance_percentage",
    "variance_color",
]
