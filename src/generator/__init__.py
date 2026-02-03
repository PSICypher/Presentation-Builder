"""Presentation generator package — PPTX builder engine.

Consumes a TemplateSchema and data payload to produce PowerPoint files.

Modules:
    pptx_builder: Core PPTX generation
    charts: Chart generation (column, line, doughnut)
"""

from .pptx_builder import PPTXBuilder, build_presentation
from .charts import add_chart, add_slide_charts

__all__ = [
    "PPTXBuilder",
    "build_presentation",
    "add_chart",
    "add_slide_charts",
]
