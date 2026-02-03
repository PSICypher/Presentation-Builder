"""Template schema models - the contract between analyzer, processor, and generator.

Defines the typed structure of a presentation template: what slides exist,
what data slots each slide contains, how values should be formatted, and
where shapes are positioned on the canvas.
"""

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


# ---------------------------------------------------------------------------
# Enums
# ---------------------------------------------------------------------------

class SlotType(Enum):
    """What kind of content a data slot holds."""
    KPI_VALUE = "kpi_value"              # Single metric (number + label + variance)
    TABLE = "table"                      # Data table with rows/columns
    CHART = "chart"                      # Chart (column, doughnut, line)
    TEXT = "text"                        # Narrative text or bullet points
    STATIC = "static"                    # Fixed content (unchanged per report)
    IMAGE = "image"                      # Logo, icon, background image
    SECTION_DIVIDER = "section_divider"  # Full-slide section break


class ChartType(Enum):
    """Supported chart types in templates."""
    COLUMN_CLUSTERED = "column_clustered"
    DOUGHNUT = "doughnut"
    DOUGHNUT_EXPLODED = "doughnut_exploded"
    LINE = "line"


class FormatType(Enum):
    """How to format a numeric value for display."""
    CURRENCY = "currency"                      # <1k=$XXX, 1k-999k=$XXXk, 1m+=$X.Xm
    PERCENTAGE = "percentage"                  # X.X%
    VARIANCE_PERCENTAGE = "variance_percentage"  # +X.X% / -X.X%
    POINTS_CHANGE = "points_change"            # +X.X ppts
    NUMBER = "number"                          # <1k=XXX, 1k-999k=X,XXX, 1m+=X.Xm
    TEXT = "text"                              # Plain text, no formatting
    INTEGER = "integer"                        # Whole number with comma separators


class SlideType(Enum):
    """Categorises a slide's role in the presentation."""
    COVER = "cover"
    TABLE_OF_CONTENTS = "toc"
    SECTION_DIVIDER = "divider"
    DATA = "data"            # Contains data-bound content
    MANUAL = "manual"        # Human-authored content (upcoming promos, next steps)


# ---------------------------------------------------------------------------
# Position and styling primitives
# ---------------------------------------------------------------------------

@dataclass
class Position:
    """Shape position and dimensions in inches."""
    left: float
    top: float
    width: float
    height: float

    def to_dict(self) -> dict:
        return {"left": self.left, "top": self.top,
                "width": self.width, "height": self.height}

    @classmethod
    def from_dict(cls, d: dict) -> "Position":
        return cls(left=d["left"], top=d["top"],
                   width=d["width"], height=d["height"])


@dataclass
class FontSpec:
    """Typography specification for a text element."""
    name: str = "DM Sans"
    size_pt: float = 14.0
    bold: bool = False
    italic: bool = False
    color: str = "#000000"

    def to_dict(self) -> dict:
        d: dict[str, Any] = {"name": self.name, "size_pt": self.size_pt}
        if self.bold:
            d["bold"] = True
        if self.italic:
            d["italic"] = True
        if self.color != "#000000":
            d["color"] = self.color
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "FontSpec":
        return cls(
            name=d.get("name", "DM Sans"),
            size_pt=d.get("size_pt", 14.0),
            bold=d.get("bold", False),
            italic=d.get("italic", False),
            color=d.get("color", "#000000"),
        )


@dataclass
class FormatRule:
    """How to format and color a data value."""
    format_type: FormatType
    positive_color: str = "#00AA00"
    negative_color: str = "#CC0000"
    neutral_color: str = "#000000"

    def to_dict(self) -> dict:
        d: dict[str, Any] = {"format_type": self.format_type.value}
        if self.positive_color != "#00AA00":
            d["positive_color"] = self.positive_color
        if self.negative_color != "#CC0000":
            d["negative_color"] = self.negative_color
        if self.neutral_color != "#000000":
            d["neutral_color"] = self.neutral_color
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "FormatRule":
        return cls(
            format_type=FormatType(d["format_type"]),
            positive_color=d.get("positive_color", "#00AA00"),
            negative_color=d.get("negative_color", "#CC0000"),
            neutral_color=d.get("neutral_color", "#000000"),
        )


# ---------------------------------------------------------------------------
# Column definition for tables
# ---------------------------------------------------------------------------

@dataclass
class TableColumn:
    """Definition of a single column in a data table."""
    header: str              # Display header text
    data_key: str            # Key in the data payload
    width_inches: float | None = None
    format_rule: FormatRule | None = None
    font: FontSpec | None = None
    alignment: str = "left"  # left, center, right

    def to_dict(self) -> dict:
        d: dict[str, Any] = {
            "header": self.header,
            "data_key": self.data_key,
            "alignment": self.alignment,
        }
        if self.width_inches is not None:
            d["width_inches"] = self.width_inches
        if self.format_rule:
            d["format_rule"] = self.format_rule.to_dict()
        if self.font:
            d["font"] = self.font.to_dict()
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "TableColumn":
        return cls(
            header=d["header"],
            data_key=d["data_key"],
            width_inches=d.get("width_inches"),
            format_rule=FormatRule.from_dict(d["format_rule"]) if d.get("format_rule") else None,
            font=FontSpec.from_dict(d["font"]) if d.get("font") else None,
            alignment=d.get("alignment", "left"),
        )


# ---------------------------------------------------------------------------
# Chart series configuration
# ---------------------------------------------------------------------------

@dataclass
class ChartSeries:
    """Configuration for a single data series in a chart."""
    name: str             # Series display name
    data_key: str         # Key in the data payload for this series' values
    color: str | None = None  # Override color for this series

    def to_dict(self) -> dict:
        d: dict[str, Any] = {"name": self.name, "data_key": self.data_key}
        if self.color:
            d["color"] = self.color
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "ChartSeries":
        return cls(name=d["name"], data_key=d["data_key"], color=d.get("color"))


# ---------------------------------------------------------------------------
# DataSlot — a named, positioned location for data on a slide
# ---------------------------------------------------------------------------

@dataclass
class DataSlot:
    """A single addressable location on a slide where data is rendered.

    Each slot has a unique name within its slide, a type describing what kind
    of content it holds, and a data_key that the processor uses to supply the
    correct value from the data payload.
    """
    name: str                            # Unique within slide, e.g. "total_revenue"
    slot_type: SlotType
    data_key: str                        # Binding key, e.g. "cover.total_revenue"
    position: Position

    # Text/KPI styling
    font: FontSpec | None = None
    format_rule: FormatRule | None = None

    # KPI-specific
    label: str | None = None             # Static label text below/above the value
    variance_key: str | None = None      # Data key for the variance indicator

    # Table-specific
    columns: list[TableColumn] = field(default_factory=list)
    row_data_key: str | None = None      # Data key for the list of row dicts

    # Chart-specific
    chart_type: ChartType | None = None
    series: list[ChartSeries] = field(default_factory=list)
    categories_key: str | None = None    # Data key for category labels

    # Shape reference (for matching to analyzer output)
    shape_name: str | None = None        # python-pptx shape name from template

    def to_dict(self) -> dict:
        d: dict[str, Any] = {
            "name": self.name,
            "slot_type": self.slot_type.value,
            "data_key": self.data_key,
            "position": self.position.to_dict(),
        }
        if self.font:
            d["font"] = self.font.to_dict()
        if self.format_rule:
            d["format_rule"] = self.format_rule.to_dict()
        if self.label:
            d["label"] = self.label
        if self.variance_key:
            d["variance_key"] = self.variance_key
        if self.columns:
            d["columns"] = [c.to_dict() for c in self.columns]
        if self.row_data_key:
            d["row_data_key"] = self.row_data_key
        if self.chart_type:
            d["chart_type"] = self.chart_type.value
        if self.series:
            d["series"] = [s.to_dict() for s in self.series]
        if self.categories_key:
            d["categories_key"] = self.categories_key
        if self.shape_name:
            d["shape_name"] = self.shape_name
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "DataSlot":
        return cls(
            name=d["name"],
            slot_type=SlotType(d["slot_type"]),
            data_key=d["data_key"],
            position=Position.from_dict(d["position"]),
            font=FontSpec.from_dict(d["font"]) if d.get("font") else None,
            format_rule=FormatRule.from_dict(d["format_rule"]) if d.get("format_rule") else None,
            label=d.get("label"),
            variance_key=d.get("variance_key"),
            columns=[TableColumn.from_dict(c) for c in d.get("columns", [])],
            row_data_key=d.get("row_data_key"),
            chart_type=ChartType(d["chart_type"]) if d.get("chart_type") else None,
            series=[ChartSeries.from_dict(s) for s in d.get("series", [])],
            categories_key=d.get("categories_key"),
            shape_name=d.get("shape_name"),
        )


# ---------------------------------------------------------------------------
# SlideSchema — one slide in the presentation
# ---------------------------------------------------------------------------

@dataclass
class SlideSchema:
    """Schema for a single slide in the presentation template."""
    index: int                           # 0-based slide position
    name: str                            # Machine name, e.g. "cover_kpis"
    title: str                           # Human-readable title
    slide_type: SlideType
    data_source: str                     # Which data source feeds this slide
    layout: str = "Title Only"           # PowerPoint layout name
    slots: list[DataSlot] = field(default_factory=list)
    is_static: bool = False              # True for TOC, dividers (no data binding)

    def to_dict(self) -> dict:
        d: dict[str, Any] = {
            "index": self.index,
            "name": self.name,
            "title": self.title,
            "slide_type": self.slide_type.value,
            "data_source": self.data_source,
            "layout": self.layout,
        }
        if self.is_static:
            d["is_static"] = True
        if self.slots:
            d["slots"] = [s.to_dict() for s in self.slots]
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "SlideSchema":
        return cls(
            index=d["index"],
            name=d["name"],
            title=d["title"],
            slide_type=SlideType(d["slide_type"]),
            data_source=d["data_source"],
            layout=d.get("layout", "Title Only"),
            slots=[DataSlot.from_dict(s) for s in d.get("slots", [])],
            is_static=d.get("is_static", False),
        )


# ---------------------------------------------------------------------------
# DesignSystem — global styling rules
# ---------------------------------------------------------------------------

@dataclass
class DesignSystem:
    """Brand design system applied across all slides."""
    # Colors
    brand_blue: str = "#0065E0"
    dark_text: str = "#000000"
    white: str = "#FFFFFF"
    dark_blue: str = "#190263"
    dark_grey: str = "#1C2B33"
    accent_green: str = "#00E167"
    positive: str = "#00AA00"
    negative: str = "#CC0000"
    light_gray: str = "#D1D5DB"
    divider_bg: str = "#0065E0"  # Section divider background

    # Typography
    primary_font: str = "DM Sans"
    title_size_pt: float = 36.0
    header_size_pt: float = 24.0
    body_size_pt: float = 14.0
    kpi_number_size_pt: float = 48.0
    kpi_label_size_pt: float = 12.0
    caption_size_pt: float = 9.0

    def to_dict(self) -> dict:
        return {
            "colors": {
                "brand_blue": self.brand_blue,
                "dark_text": self.dark_text,
                "white": self.white,
                "dark_blue": self.dark_blue,
                "dark_grey": self.dark_grey,
                "accent_green": self.accent_green,
                "positive": self.positive,
                "negative": self.negative,
                "light_gray": self.light_gray,
                "divider_bg": self.divider_bg,
            },
            "typography": {
                "primary_font": self.primary_font,
                "title_size_pt": self.title_size_pt,
                "header_size_pt": self.header_size_pt,
                "body_size_pt": self.body_size_pt,
                "kpi_number_size_pt": self.kpi_number_size_pt,
                "kpi_label_size_pt": self.kpi_label_size_pt,
                "caption_size_pt": self.caption_size_pt,
            },
        }

    @classmethod
    def from_dict(cls, d: dict) -> "DesignSystem":
        colors = d.get("colors", {})
        typo = d.get("typography", {})
        return cls(
            brand_blue=colors.get("brand_blue", "#0065E0"),
            dark_text=colors.get("dark_text", "#000000"),
            white=colors.get("white", "#FFFFFF"),
            dark_blue=colors.get("dark_blue", "#190263"),
            dark_grey=colors.get("dark_grey", "#1C2B33"),
            accent_green=colors.get("accent_green", "#00E167"),
            positive=colors.get("positive", "#00AA00"),
            negative=colors.get("negative", "#CC0000"),
            light_gray=colors.get("light_gray", "#D1D5DB"),
            divider_bg=colors.get("divider_bg", "#0065E0"),
            primary_font=typo.get("primary_font", "DM Sans"),
            title_size_pt=typo.get("title_size_pt", 36.0),
            header_size_pt=typo.get("header_size_pt", 24.0),
            body_size_pt=typo.get("body_size_pt", 14.0),
            kpi_number_size_pt=typo.get("kpi_number_size_pt", 48.0),
            kpi_label_size_pt=typo.get("kpi_label_size_pt", 12.0),
            caption_size_pt=typo.get("caption_size_pt", 9.0),
        )


# ---------------------------------------------------------------------------
# TemplateSchema — top-level container
# ---------------------------------------------------------------------------

@dataclass
class TemplateSchema:
    """Complete schema for a presentation template.

    This is the central artefact that ties together dimensions, design system,
    and per-slide schemas with their data slots. The generator consumes this
    along with a data payload to produce the final PPTX.
    """
    name: str
    report_type: str                     # "monthly" or "qbr"
    width_inches: float
    height_inches: float
    design: DesignSystem
    slides: list[SlideSchema]
    naming_convention: str = ""          # Output filename template

    def get_slide(self, name: str) -> SlideSchema | None:
        """Look up a slide by its machine name."""
        for s in self.slides:
            if s.name == name:
                return s
        return None

    def data_slides(self) -> list[SlideSchema]:
        """Return only slides that require data binding."""
        return [s for s in self.slides if not s.is_static]

    def all_data_keys(self) -> set[str]:
        """Collect every data_key referenced across all slots."""
        keys: set[str] = set()
        for slide in self.slides:
            for slot in slide.slots:
                keys.add(slot.data_key)
                if slot.variance_key:
                    keys.add(slot.variance_key)
                if slot.row_data_key:
                    keys.add(slot.row_data_key)
                if slot.categories_key:
                    keys.add(slot.categories_key)
                for col in slot.columns:
                    keys.add(col.data_key)
                for series in slot.series:
                    keys.add(series.data_key)
        return keys

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "report_type": self.report_type,
            "dimensions": {
                "width_inches": self.width_inches,
                "height_inches": self.height_inches,
            },
            "naming_convention": self.naming_convention,
            "design": self.design.to_dict(),
            "slides": [s.to_dict() for s in self.slides],
        }

    @classmethod
    def from_dict(cls, d: dict) -> "TemplateSchema":
        dims = d.get("dimensions", {})
        return cls(
            name=d["name"],
            report_type=d["report_type"],
            width_inches=dims.get("width_inches", 13.333),
            height_inches=dims.get("height_inches", 7.5),
            design=DesignSystem.from_dict(d.get("design", {})),
            slides=[SlideSchema.from_dict(s) for s in d.get("slides", [])],
            naming_convention=d.get("naming_convention", ""),
        )
