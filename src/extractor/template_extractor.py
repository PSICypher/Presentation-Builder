"""Template Extraction Engine — derives TemplateSchema from PPTX analysis.

Takes the raw JSON output from TemplateAnalyzer and produces a typed
TemplateSchema by classifying slides, identifying data-bearing shapes,
extracting the design system, and mapping shapes to DataSlots.

Classification heuristics (from template analysis of monthly-report, qbr,
and ingenuity-qbr templates):

Slide classification:
    - Dividers:  ≤3 shapes, FREEFORM with brand-blue fill, large centered text
    - Cover:     First slide, large KPI values (≥40pt bold), report title
    - TOC:       Near start, "Content" title, bullet-list text
    - Data:      Contains tables, charts, or structured data text
    - Manual:    Tail slides without heavy data binding

Shape-to-slot mapping:
    - TABLE shapes      → SlotType.TABLE  (columns from headers)
    - CHART shapes      → SlotType.CHART  (type + series count)
    - ≥40pt bold text   → SlotType.KPI_VALUE (on cover slides)
    - Narrative text    → SlotType.TEXT
    - Images            → SlotType.IMAGE
"""

from __future__ import annotations

import json
import re
from collections import Counter
from pathlib import Path
from typing import Any

from src.schema.models import (
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

# Brand-blue hex variants seen across templates (case-insensitive matching)
_BRAND_BLUES = {"0065e2", "0065e0", "0064e0", "0063e0", "0066cc"}

# Minimum font size (pt) to qualify as a KPI number
_KPI_FONT_MIN_PT = 36.0

# Maximum shape count for a slide to be classified as a divider
_DIVIDER_MAX_SHAPES = 4

# Chart type string mapping from python-pptx to our enum
_CHART_TYPE_MAP = {
    "COLUMN_CLUSTERED (51)": ChartType.COLUMN_CLUSTERED,
    "DOUGHNUT (-4120)": ChartType.DOUGHNUT,
    "DOUGHNUT_EXPLODED (-4121)": ChartType.DOUGHNUT_EXPLODED,
    "LINE (4)": ChartType.LINE,
}


def _normalize_hex(color: str) -> str:
    """Normalize a hex color to lowercase without '#' prefix."""
    return color.lstrip("#").lower()


def _has_brand_blue_fill(shape: dict) -> bool:
    """Check if a shape has a brand-blue fill color."""
    for c in shape.get("fill_colors", []):
        if _normalize_hex(c) in _BRAND_BLUES:
            return True
    return False


def _is_picture(shape: dict) -> bool:
    return shape.get("is_picture", False)


def _is_freeform(shape: dict) -> bool:
    return "FREEFORM" in shape.get("shape_type", "")


def _has_table(shape: dict) -> bool:
    return "table" in shape


def _has_chart(shape: dict) -> bool:
    return "chart" in shape


def _get_text(shape: dict) -> str:
    """Extract concatenated text from a shape's paragraphs."""
    parts = []
    for para in shape.get("text", []):
        t = para.get("text", "").strip()
        if t:
            parts.append(t)
    return " ".join(parts)


def _get_first_run_font(shape: dict) -> dict | None:
    """Extract the font dict from the first non-empty run."""
    for para in shape.get("text", []):
        for run in para.get("runs", []):
            if run.get("font"):
                return run["font"]
    return None


def _font_size(shape: dict) -> float:
    """Get the font size of the first run, or 0."""
    font = _get_first_run_font(shape)
    return font.get("size_pt", 0) if font else 0


def _is_bold(shape: dict) -> bool:
    font = _get_first_run_font(shape)
    return font.get("bold", False) if font else False


def _make_position(shape: dict) -> Position:
    """Convert a shape's position dict to a Position dataclass."""
    pos = shape["position"]
    return Position(
        left=pos.get("left", 0),
        top=pos.get("top", 0),
        width=pos.get("width", 0),
        height=pos.get("height", 0),
    )


def _make_font_spec(font_dict: dict | None) -> FontSpec | None:
    """Convert a raw font dict to a FontSpec, or None."""
    if not font_dict:
        return None
    color = font_dict.get("color", "000000")
    if not color.startswith("#"):
        color = "#" + color
    return FontSpec(
        name=font_dict.get("name", "DM Sans"),
        size_pt=font_dict.get("size_pt", 14.0),
        bold=font_dict.get("bold", False),
        italic=font_dict.get("italic", False),
        color=color,
    )


def _slugify(text: str) -> str:
    """Convert text to a snake_case slug for data keys."""
    text = text.lower().strip()
    text = re.sub(r"[^a-z0-9\s]", "", text)
    text = re.sub(r"\s+", "_", text)
    return text[:40].rstrip("_")


def _unique_name(base: str, used: set[str]) -> str:
    """Return a unique name by appending an index suffix if needed."""
    if base not in used:
        return base
    idx = 1
    while f"{base}_{idx}" in used:
        idx += 1
    return f"{base}_{idx}"


def _infer_format_type(text: str) -> FormatType:
    """Guess the FormatType from sample text content."""
    text = text.strip()
    if text.startswith("$"):
        return FormatType.CURRENCY
    if text.endswith("%"):
        if text.startswith("+") or text.startswith("-"):
            return FormatType.VARIANCE_PERCENTAGE
        return FormatType.PERCENTAGE
    if text.endswith("ppts"):
        return FormatType.POINTS_CHANGE
    # Check if it looks numeric
    cleaned = text.replace(",", "").replace("K", "").replace("k", "").replace("M", "").replace("m", "")
    try:
        float(cleaned)
        return FormatType.NUMBER
    except ValueError:
        return FormatType.TEXT


class TemplateExtractor:
    """Extracts a TemplateSchema from TemplateAnalyzer output.

    Takes a single template's analysis dict (as produced by
    TemplateAnalyzer.analyze()) and produces a TemplateSchema with
    classified slides, mapped DataSlots, and an extracted DesignSystem.
    """

    def __init__(self, analysis: dict[str, Any]):
        self.analysis = analysis
        self.source_file = analysis.get("source_file", "unknown")
        self.slides_data: list[dict] = analysis.get("slides", [])
        self.summary: dict = analysis.get("summary", {})
        self.dimensions: dict = analysis.get("dimensions", {})

        # Counters populated during extraction
        self._all_fonts: Counter = Counter()
        self._all_colors: Counter = Counter()
        self._all_font_sizes: Counter = Counter()

    def extract(self) -> TemplateSchema:
        """Run full extraction and return a TemplateSchema."""
        # Collect font/color stats from summary
        self._all_fonts = Counter(self.summary.get("fonts", {}))
        self._all_colors = Counter(self.summary.get("colors_hex", {}))
        self._all_font_sizes = Counter(self.summary.get("font_sizes_pt", {}))

        # Classify and extract each slide (with unique name tracking)
        slide_schemas = []
        used_slide_names: set[str] = set()
        for slide_data in self.slides_data:
            slide_schema = self._extract_slide(slide_data)
            # Ensure unique slide names across the presentation
            slide_schema.name = _unique_name(slide_schema.name, used_slide_names)
            used_slide_names.add(slide_schema.name)
            slide_schemas.append(slide_schema)

        # Extract design system from collected statistics
        design = self._extract_design_system()

        # Determine report type from filename or dimensions
        report_type = self._infer_report_type()

        # Build naming convention
        naming = self._infer_naming_convention()

        return TemplateSchema(
            name=self._infer_name(),
            report_type=report_type,
            width_inches=self.dimensions.get("width_inches", 13.333),
            height_inches=self.dimensions.get("height_inches", 7.5),
            design=design,
            slides=slide_schemas,
            naming_convention=naming,
        )

    def _classify_slide(self, slide_data: dict) -> SlideType:
        """Classify a slide based on its shapes and position."""
        idx = slide_data.get("index", 0)
        shapes = slide_data.get("shapes", [])
        total_slides = len(self.slides_data)

        # Count content-bearing shapes (exclude pictures used as logos/backgrounds)
        non_picture_shapes = [s for s in shapes if not _is_picture(s)]
        text_shapes = [s for s in shapes if s.get("text")]
        table_count = sum(1 for s in shapes if _has_table(s))
        chart_count = sum(1 for s in shapes if _has_chart(s))

        # --- Divider detection ---
        # Dividers have very few shapes, typically a freeform with brand-blue
        # fill and a single text element, plus possibly a logo image
        if len(non_picture_shapes) <= 2:
            has_freeform_blue = any(
                _is_freeform(s) and _has_brand_blue_fill(s)
                for s in shapes
            )
            if has_freeform_blue:
                return SlideType.SECTION_DIVIDER

        # Minimal slides with just brand-colored background shapes
        if len(non_picture_shapes) <= 1 and len(shapes) <= _DIVIDER_MAX_SHAPES:
            # Check if the non-picture shapes have large centered text
            for s in non_picture_shapes:
                if _font_size(s) >= 30 and _has_brand_blue_fill(s):
                    return SlideType.SECTION_DIVIDER

        # --- Cover detection ---
        # First slide with large KPI-style numbers
        if idx == 0:
            kpi_shapes = [
                s for s in text_shapes
                if _font_size(s) >= _KPI_FONT_MIN_PT and _is_bold(s)
            ]
            if kpi_shapes:
                return SlideType.COVER

        # --- TOC detection ---
        # Second or third slide with "Content" in title text
        if idx in (1, 2):
            for s in text_shapes:
                text = _get_text(s).lower()
                if "content" in text or "table of contents" in text or "agenda" in text:
                    return SlideType.TABLE_OF_CONTENTS

        # --- Manual slide detection ---
        # Last 2 slides with sparse data
        if idx >= total_slides - 2 and table_count == 0 and chart_count == 0:
            # Check for "next steps" / "upcoming" / "action" type content
            all_text = " ".join(_get_text(s).lower() for s in text_shapes)
            if any(kw in all_text for kw in ("next step", "upcoming", "action", "promotional plan")):
                return SlideType.MANUAL

        # --- Data slide (default for anything with content) ---
        if table_count > 0 or chart_count > 0 or len(text_shapes) >= 3:
            return SlideType.DATA

        # Last resort: if it has very few shapes and is near the end, manual
        if idx >= total_slides - 2:
            return SlideType.MANUAL

        return SlideType.DATA

    def _extract_slide(self, slide_data: dict) -> SlideSchema:
        """Extract a SlideSchema from a single slide's analysis data."""
        idx = slide_data.get("index", 0)
        shapes = slide_data.get("shapes", [])
        layout = slide_data.get("layout", "Title Only")
        slide_type = self._classify_slide(slide_data)

        # Derive a human name and machine name from content
        title, name = self._derive_slide_names(slide_data, slide_type, idx)

        # Determine data source hint
        data_source = self._infer_data_source(slide_data, slide_type)

        # Extract slots based on slide type
        slots = self._extract_slots(slide_data, slide_type, name)

        is_static = slide_type in (SlideType.SECTION_DIVIDER, SlideType.TABLE_OF_CONTENTS)

        return SlideSchema(
            index=idx,
            name=name,
            title=title,
            slide_type=slide_type,
            data_source=data_source,
            layout=layout,
            slots=slots,
            is_static=is_static,
        )

    def _derive_slide_names(self, slide_data: dict, slide_type: SlideType, idx: int) -> tuple[str, str]:
        """Derive a human title and machine name for a slide."""
        shapes = slide_data.get("shapes", [])

        # Find the most prominent text (largest font, or first text shape)
        title_text = ""
        best_size = 0
        for s in shapes:
            text = _get_text(s)
            size = _font_size(s)
            if text and size > best_size:
                title_text = text
                best_size = size

        # Fallback: any text shape
        if not title_text:
            for s in shapes:
                text = _get_text(s)
                if text and not text.startswith("*Data"):
                    title_text = text
                    break

        # Generate names from slide type and content
        if slide_type == SlideType.COVER:
            return (title_text or "Cover + KPIs", "cover_kpis")
        if slide_type == SlideType.TABLE_OF_CONTENTS:
            return ("Table of Contents", "toc")
        if slide_type == SlideType.SECTION_DIVIDER:
            slug = _slugify(title_text) if title_text else f"divider_{idx}"
            return (title_text or f"Section Divider {idx}", slug)

        # Data/manual slides: use prominent text as title
        if title_text:
            slug = _slugify(title_text)
            return (title_text, slug or f"slide_{idx}")

        return (f"Slide {idx}", f"slide_{idx}")

    def _infer_data_source(self, slide_data: dict, slide_type: SlideType) -> str:
        """Infer the data source for a slide from its title text.

        Only uses the slide title (largest font text, typically ≥20pt) for
        keyword matching. Narrative body text contains false positives
        (e.g., executive summary mentioning "crm" or "email" in analysis).
        """
        if slide_type == SlideType.SECTION_DIVIDER:
            return "static"
        if slide_type == SlideType.TABLE_OF_CONTENTS:
            return "static"
        if slide_type == SlideType.COVER:
            return "tracker:mtd_reporting"
        if slide_type == SlideType.MANUAL:
            return "manual"

        # Collect only title-level text (≥20pt or shape names like "object 3")
        shapes = slide_data.get("shapes", [])
        title_texts = []
        for s in shapes:
            text = _get_text(s)
            if not text:
                continue
            # Title shapes: large font or "object" name pattern (slide titles)
            if _font_size(s) >= 20 or s.get("name", "").startswith("object"):
                title_texts.append(text.lower())

        title_text = " ".join(title_texts)

        if "crm" in title_text or "email" in title_text:
            return "crm_data"
        if "affiliate" in title_text:
            return "affiliate_data"
        if "seo" in title_text or "organic" in title_text:
            return "tracker:organic"
        if "promotion" in title_text or "offer" in title_text:
            return "offer_performance"
        if "product" in title_text:
            return "product_sales"
        if "daily" in title_text:
            return "tracker:daily"

        return "tracker"

    def _extract_slots(self, slide_data: dict, slide_type: SlideType, slide_name: str) -> list[DataSlot]:
        """Extract DataSlots from a slide's shapes."""
        if slide_type == SlideType.SECTION_DIVIDER:
            return self._extract_divider_slots(slide_data, slide_name)
        if slide_type == SlideType.COVER:
            return self._extract_cover_slots(slide_data, slide_name)
        if slide_type == SlideType.TABLE_OF_CONTENTS:
            return self._extract_toc_slots(slide_data, slide_name)
        if slide_type == SlideType.DATA:
            return self._extract_data_slots(slide_data, slide_name)
        if slide_type == SlideType.MANUAL:
            return self._extract_manual_slots(slide_data, slide_name)
        return []

    def _extract_divider_slots(self, slide_data: dict, slide_name: str) -> list[DataSlot]:
        """Extract slots from a section divider slide (typically just a title)."""
        slots = []
        for s in slide_data.get("shapes", []):
            text = _get_text(s)
            if text and not _is_picture(s):
                slots.append(DataSlot(
                    name="divider_title",
                    slot_type=SlotType.SECTION_DIVIDER,
                    data_key=f"{slide_name}.title",
                    position=_make_position(s),
                    font=_make_font_spec(_get_first_run_font(s)),
                    shape_name=s.get("name"),
                ))
                break
        return slots

    def _extract_cover_slots(self, slide_data: dict, slide_name: str) -> list[DataSlot]:
        """Extract slots from a cover slide with KPIs."""
        slots = []
        shapes = slide_data.get("shapes", [])

        # Separate shapes by role
        kpi_values = []
        kpi_labels = []
        title_shape = None
        other_text = []

        for s in shapes:
            if _is_picture(s):
                continue

            text = _get_text(s)
            if not text:
                continue

            size = _font_size(s)
            bold = _is_bold(s)

            # Title: largest non-KPI text, typically ≥30pt
            if size >= 30 and bold and not _looks_like_kpi_value(text):
                if title_shape is None or size > _font_size(title_shape):
                    title_shape = s
            # KPI value: large bold text with numeric content
            elif size >= _KPI_FONT_MIN_PT and bold and _looks_like_kpi_value(text):
                kpi_values.append(s)
            # KPI label: small text near KPI values
            elif size <= 16 and not text.startswith("*Data"):
                kpi_labels.append(s)
            elif not text.startswith("*Data"):
                other_text.append(s)

        # Add title slot
        if title_shape:
            slots.append(DataSlot(
                name="report_title",
                slot_type=SlotType.TEXT,
                data_key=f"{slide_name}.report_title",
                position=_make_position(title_shape),
                font=_make_font_spec(_get_first_run_font(title_shape)),
                shape_name=title_shape.get("name"),
            ))

        # Pair KPI values with their labels (by horizontal position proximity)
        kpi_values.sort(key=lambda s: s["position"]["left"])
        kpi_labels.sort(key=lambda s: s["position"]["left"])

        # Deduplicate labels (template may have overlapping shapes)
        seen_labels = set()
        unique_labels = []
        for label in kpi_labels:
            text = _get_text(label)
            if text not in seen_labels:
                seen_labels.add(text)
                unique_labels.append(label)
        kpi_labels = unique_labels

        for i, kpi_shape in enumerate(kpi_values):
            kpi_text = _get_text(kpi_shape)
            fmt = _infer_format_type(kpi_text)

            # Find the nearest label below this KPI
            label_text = None
            kpi_left = kpi_shape["position"]["left"]
            best_label = None
            best_dist = float("inf")
            for label in kpi_labels:
                label_left = label["position"]["left"]
                dist = abs(label_left - kpi_left)
                if dist < best_dist:
                    best_dist = dist
                    best_label = label
                    label_text = _get_text(label)

            if best_label and best_dist < 2.0:
                kpi_labels.remove(best_label)
            else:
                label_text = None

            slot_name = _slugify(label_text) if label_text else f"kpi_{i}"
            slots.append(DataSlot(
                name=slot_name,
                slot_type=SlotType.KPI_VALUE,
                data_key=f"{slide_name}.{slot_name}",
                position=_make_position(kpi_shape),
                font=_make_font_spec(_get_first_run_font(kpi_shape)),
                format_rule=FormatRule(format_type=fmt),
                label=label_text.strip() if label_text else None,
                variance_key=f"{slide_name}.{slot_name}_variance" if label_text else None,
                shape_name=kpi_shape.get("name"),
            ))

        return slots

    def _extract_toc_slots(self, slide_data: dict, slide_name: str) -> list[DataSlot]:
        """Extract slots from a table-of-contents slide."""
        slots = []
        shapes = slide_data.get("shapes", [])

        for s in shapes:
            if _is_picture(s):
                continue
            text = _get_text(s)
            if not text:
                continue

            size = _font_size(s)
            if size >= 30:
                # Title
                slots.append(DataSlot(
                    name="toc_title",
                    slot_type=SlotType.STATIC,
                    data_key=f"{slide_name}.title",
                    position=_make_position(s),
                    font=_make_font_spec(_get_first_run_font(s)),
                    shape_name=s.get("name"),
                ))
            elif len(text) > 20:
                # Bullet list content
                slots.append(DataSlot(
                    name="toc_items",
                    slot_type=SlotType.STATIC,
                    data_key=f"{slide_name}.items",
                    position=_make_position(s),
                    font=_make_font_spec(_get_first_run_font(s)),
                    shape_name=s.get("name"),
                ))

        return slots

    def _extract_data_slots(self, slide_data: dict, slide_name: str) -> list[DataSlot]:
        """Extract slots from a data slide (tables, charts, text)."""
        slots = []
        shapes = slide_data.get("shapes", [])
        table_idx = 0
        chart_idx = 0
        text_idx = 0
        used_names: set[str] = set()

        for s in shapes:
            if _is_picture(s):
                continue

            text = _get_text(s)

            # Table shape
            if _has_table(s):
                table_info = s["table"]
                columns = self._extract_table_columns(table_info, slide_name, table_idx)
                slot_name = f"table_{table_idx}" if table_idx > 0 else "main_table"
                slots.append(DataSlot(
                    name=slot_name,
                    slot_type=SlotType.TABLE,
                    data_key=f"{slide_name}.{slot_name}",
                    position=_make_position(s),
                    row_data_key=f"{slide_name}.{slot_name}_rows",
                    columns=columns,
                    shape_name=s.get("name"),
                ))
                used_names.add(slot_name)
                table_idx += 1
                continue

            # Chart shape
            if _has_chart(s):
                chart_info = s["chart"]
                chart_type_str = chart_info.get("chart_type", "")
                chart_type = _CHART_TYPE_MAP.get(chart_type_str)
                slot_name = f"chart_{chart_idx}" if chart_idx > 0 else "main_chart"
                slots.append(DataSlot(
                    name=slot_name,
                    slot_type=SlotType.CHART,
                    data_key=f"{slide_name}.{slot_name}",
                    position=_make_position(s),
                    chart_type=chart_type,
                    categories_key=f"{slide_name}.{slot_name}_categories",
                    shape_name=s.get("name"),
                ))
                used_names.add(slot_name)
                chart_idx += 1
                continue

            # Text shapes
            if not text:
                continue

            # Skip data source notes
            if text.startswith("*Data"):
                continue

            size = _font_size(s)

            # Slide title (large text, typically the "object 3" shape)
            if size >= 20 and "slide_title" not in used_names:
                slots.append(DataSlot(
                    name="slide_title",
                    slot_type=SlotType.TEXT,
                    data_key=f"{slide_name}.title",
                    position=_make_position(s),
                    font=_make_font_spec(_get_first_run_font(s)),
                    shape_name=s.get("name"),
                ))
                used_names.add("slide_title")
            elif "key call-out" in text.lower() or "overview" in text.lower().replace("oveview", "overview"):
                slot_name = _unique_name("callout_header", used_names)
                slots.append(DataSlot(
                    name=slot_name,
                    slot_type=SlotType.TEXT,
                    data_key=f"{slide_name}.{slot_name}",
                    position=_make_position(s),
                    font=_make_font_spec(_get_first_run_font(s)),
                    shape_name=s.get("name"),
                ))
                used_names.add(slot_name)
            elif len(text) > 30:
                slot_name = _unique_name("narrative", used_names)
                slots.append(DataSlot(
                    name=slot_name,
                    slot_type=SlotType.TEXT,
                    data_key=f"{slide_name}.{slot_name}",
                    position=_make_position(s),
                    font=_make_font_spec(_get_first_run_font(s)),
                    shape_name=s.get("name"),
                ))
                used_names.add(slot_name)

        return slots

    def _extract_manual_slots(self, slide_data: dict, slide_name: str) -> list[DataSlot]:
        """Extract slots from a manual-entry slide."""
        slots = []
        shapes = slide_data.get("shapes", [])
        used_names: set[str] = set()

        for s in shapes:
            if _is_picture(s):
                continue

            text = _get_text(s)
            if not text:
                continue

            size = _font_size(s)

            if size >= 20 or _is_bold(s):
                base_name = _slugify(text)[:20] or "section_header"
                slot_name = _unique_name(base_name, used_names)
                slots.append(DataSlot(
                    name=slot_name,
                    slot_type=SlotType.TEXT,
                    data_key=f"{slide_name}.{slot_name}",
                    position=_make_position(s),
                    font=_make_font_spec(_get_first_run_font(s)),
                    shape_name=s.get("name"),
                ))
                used_names.add(slot_name)
            elif len(text) > 10:
                slot_name = _unique_name("content", used_names)
                slots.append(DataSlot(
                    name=slot_name,
                    slot_type=SlotType.TEXT,
                    data_key=f"{slide_name}.{slot_name}",
                    position=_make_position(s),
                    font=_make_font_spec(_get_first_run_font(s)),
                    shape_name=s.get("name"),
                ))
                used_names.add(slot_name)

        return slots

    def _extract_table_columns(self, table_info: dict, slide_name: str, table_idx: int) -> list[TableColumn]:
        """Extract column definitions from a table's header row."""
        columns = []
        headers = table_info.get("headers", [])
        col_widths = table_info.get("col_widths_inches", [])

        for i, header in enumerate(headers):
            header_text = header.get("text", f"Column {i}").strip()
            if not header_text:
                header_text = f"col_{i}"

            data_key = _slugify(header_text) or f"col_{i}"
            width = col_widths[i] if i < len(col_widths) else None

            font = None
            if header.get("font"):
                font = _make_font_spec(header["font"])

            columns.append(TableColumn(
                header=header_text,
                data_key=data_key,
                width_inches=width,
                font=font,
                alignment="left",
            ))

        return columns

    def _extract_design_system(self) -> DesignSystem:
        """Extract the design system from template-wide statistics."""
        colors = self.summary.get("colors_hex", {})
        fonts = self.summary.get("fonts", {})

        # Find brand blue (most common blue-ish color)
        brand_blue = "#0065E0"
        for color, count in sorted(colors.items(), key=lambda x: -x[1]):
            if _normalize_hex(color) in _BRAND_BLUES:
                brand_blue = f"#{_normalize_hex(color).upper()}"
                break

        # Primary font: most frequently used
        primary_font = "DM Sans"
        if fonts:
            primary_font = max(fonts, key=fonts.get)

        # Extract theme data if available
        theme = self.analysis.get("theme", {})
        theme_colors = {}
        if isinstance(theme, dict):
            theme_colors = theme.get("theme_colors", {})
            if "masters" in theme:
                # Multiple masters — use the first one with useful colors
                for master_theme in theme.get("masters", []):
                    tc = master_theme.get("theme_colors", {})
                    if tc:
                        theme_colors = tc
                        break

        return DesignSystem(
            brand_blue=brand_blue,
            dark_text="#000000",
            white="#FFFFFF",
            dark_blue=f"#{theme_colors.get('dk2', '190263').upper()}" if theme_colors.get("dk2") else "#190263",
            dark_grey="#1C2B33",
            accent_green="#00E167",
            positive="#00AA00",
            negative="#CC0000",
            light_gray="#D1D5DB",
            divider_bg=brand_blue,
            primary_font=primary_font,
        )

    def _infer_report_type(self) -> str:
        """Infer report type from filename."""
        name = self.source_file.lower()
        if "qbr" in name:
            return "qbr"
        return "monthly"

    def _infer_name(self) -> str:
        """Infer a descriptive name from the source file."""
        stem = Path(self.source_file).stem
        # Convert kebab-case to title case
        return stem.replace("-", " ").replace("_", " ").title()

    def _infer_naming_convention(self) -> str:
        """Infer output naming convention."""
        if "monthly" in self.source_file.lower():
            return "No7 US x THGi Monthly eComm Report - {month} {year} Overview.pptx"
        if "qbr" in self.source_file.lower():
            return "No7 US x THGi QBR - {quarter} {year}.pptx"
        return "{name} - {month} {year}.pptx"


def _looks_like_kpi_value(text: str) -> bool:
    """Check if text looks like a KPI numeric value (e.g., $209.2K, 3.6K, 3.9%).

    KPI values are short (≤15 chars), mostly numeric, and typically formatted
    as currency ($209.2K), counts (3.6K), or rates (3.9%).
    """
    text = text.strip()
    if not text or len(text) > 15:
        return False
    # Starts with $
    if text.startswith("$"):
        return True
    # Mostly digits with optional suffix (K, M, %)
    digit_count = sum(1 for c in text if c.isdigit())
    if digit_count == 0:
        return False
    # At least 30% digits for short numeric strings
    if digit_count / len(text) < 0.3:
        return False
    if any(c in text.upper() for c in "KM%"):
        return True
    # Pure number
    try:
        float(text.replace(",", ""))
        return True
    except ValueError:
        return False


def extract_template(analysis: dict[str, Any]) -> TemplateSchema:
    """Convenience function: extract a TemplateSchema from analysis data."""
    extractor = TemplateExtractor(analysis)
    return extractor.extract()


def extract_from_file(analysis_path: str | Path) -> list[TemplateSchema]:
    """Extract TemplateSchemas from a JSON analysis file.

    The file may contain a single template analysis dict or a list of them.
    Returns a list of TemplateSchema objects.
    """
    path = Path(analysis_path)
    with open(path) as f:
        data = json.load(f)

    if isinstance(data, list):
        return [extract_template(d) for d in data]
    return [extract_template(data)]


if __name__ == "__main__":
    import argparse
    import sys

    parser = argparse.ArgumentParser(description="Extract TemplateSchema from template analysis")
    parser.add_argument("input", help="Path to template analysis JSON file")
    parser.add_argument("-o", "--output", help="Output path for extracted schema(s)")
    parser.add_argument("-f", "--format", choices=["json", "yaml"], default="yaml")
    parser.add_argument("--template", help="Extract only this template (by filename)")
    parser.add_argument("--summary", action="store_true", help="Print summary only")
    args = parser.parse_args()

    schemas = extract_from_file(args.input)

    if args.template:
        schemas = [s for s in schemas if args.template.lower() in s.name.lower()]
        if not schemas:
            print(f"No template matching '{args.template}' found", file=sys.stderr)
            sys.exit(1)

    for schema in schemas:
        print(f"\n{'='*60}")
        print(f"  {schema.name} ({schema.report_type})")
        print(f"{'='*60}")
        print(f"  Dimensions: {schema.width_inches}\" x {schema.height_inches}\"")
        print(f"  Slides: {len(schema.slides)}")
        print(f"  Data slides: {len(schema.data_slides())}")
        print(f"  Data keys: {len(schema.all_data_keys())}")
        print(f"  Naming: {schema.naming_convention}")

        if not args.summary:
            for slide in schema.slides:
                marker = "*" if not slide.is_static else " "
                print(f"  {marker} [{slide.index:2d}] {slide.slide_type.value:8s} | {slide.name:30s} | {len(slide.slots)} slots | src={slide.data_source}")
                for slot in slide.slots:
                    print(f"         └─ {slot.slot_type.value:12s} {slot.name:20s} → {slot.data_key}")

    if args.output:
        from src.schema.loader import save_schema
        if len(schemas) == 1:
            save_schema(schemas[0], args.output)
        else:
            # Save each schema separately
            out_dir = Path(args.output)
            out_dir.mkdir(parents=True, exist_ok=True)
            for schema in schemas:
                slug = schema.name.lower().replace(" ", "_")
                ext = "yaml" if args.format == "yaml" else "json"
                save_schema(schema, out_dir / f"{slug}_schema.{ext}")
        print(f"\n✓ Schema(s) saved to {args.output}")
