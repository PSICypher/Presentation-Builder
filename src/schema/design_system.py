"""Design system utilities — value formatting functions.

Implements the formatting rules from the CoWork QBR study notes:
- Currency: <$1k=$XXX, $1k-$999k=$XXXk, $1m+=$X.Xm
- Percentages: X.X%
- Variances: +X.X% / -X.X%
- Points: +X.X ppts
- Numbers: <1k=XXX, 1k-999k=X,XXX, 1m+=X.Xm
"""

import math

from .models import FormatType


def format_currency(value: float | int | None) -> str:
    """Format a dollar value using tiered abbreviation.

    <$1k   -> $XXX
    $1k-$999k -> $XXXk
    $1m+   -> $X.Xm
    """
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "N/A"
    v = abs(value)
    sign = "-" if value < 0 else ""
    if v < 1_000:
        return f"{sign}${v:,.0f}"
    if v < 1_000_000:
        k = v / 1_000
        if k == int(k):
            return f"{sign}${int(k)}k"
        return f"{sign}${k:.1f}k".rstrip("0").rstrip(".")  + "k" * (not f"{k:.1f}".rstrip("0").rstrip(".").endswith("k"))
    m = v / 1_000_000
    return f"{sign}${m:.1f}m"


def _format_currency(value: float | int | None) -> str:
    """Format a dollar value using tiered abbreviation."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "N/A"
    v = abs(value)
    sign = "-" if value < 0 else ""
    if v < 1_000:
        return f"{sign}${v:,.0f}"
    if v < 999_950:
        k = v / 1_000
        # Clean trailing zeros: $12.0k -> $12k, $12.5k stays
        formatted = f"{k:.1f}"
        if formatted.endswith(".0"):
            formatted = formatted[:-2]
        return f"{sign}${formatted}k"
    m = v / 1_000_000
    return f"{sign}${m:.1f}m"


# Replace the initial attempt with the clean version
format_currency = _format_currency


def format_percentage(value: float | int | None) -> str:
    """Format a rate as X.X% (no sign prefix)."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "N/A"
    return f"{value:.1f}%"


def format_variance_percentage(value: float | int | None) -> str:
    """Format a variance as +X.X% or -X.X%."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "N/A"
    sign = "+" if value > 0 else ""
    return f"{sign}{value:.1f}%"


def format_points_change(value: float | int | None) -> str:
    """Format a percentage-point change as +X.X ppts or -X.X ppts."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "N/A"
    sign = "+" if value > 0 else ""
    return f"{sign}{value:.1f} ppts"


def format_number(value: float | int | None) -> str:
    """Format a number using tiered abbreviation.

    <1k    -> XXX
    1k-999k -> X,XXX (or XXXk for large values)
    1m+    -> X.Xm
    """
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "N/A"
    v = abs(value)
    sign = "-" if value < 0 else ""
    if v < 1_000:
        return f"{sign}{v:,.0f}"
    if v < 1_000_000:
        return f"{sign}{v:,.0f}"
    m = v / 1_000_000
    return f"{sign}{m:.1f}m"


def format_integer(value: float | int | None) -> str:
    """Format a whole number with comma separators."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "N/A"
    return f"{int(value):,}"


def format_value(value: float | int | str | None, format_type: FormatType) -> str:
    """Format a value according to its FormatType."""
    if isinstance(value, str):
        return value
    formatters = {
        FormatType.CURRENCY: format_currency,
        FormatType.PERCENTAGE: format_percentage,
        FormatType.VARIANCE_PERCENTAGE: format_variance_percentage,
        FormatType.POINTS_CHANGE: format_points_change,
        FormatType.NUMBER: format_number,
        FormatType.INTEGER: format_integer,
        FormatType.TEXT: lambda v: str(v) if v is not None else "N/A",
    }
    formatter = formatters.get(format_type, str)
    return formatter(value)


def variance_color(value: float | None, positive: str = "#00AA00",
                   negative: str = "#CC0000", neutral: str = "#000000") -> str:
    """Return the appropriate color hex for a variance value."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return neutral
    if value > 0:
        return positive
    if value < 0:
        return negative
    return neutral
