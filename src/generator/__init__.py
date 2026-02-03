"""Presentation generator package — renders data into PPTX shapes.

Modules:
    charts: Chart generation (column, line, doughnut)
"""

from .charts import add_chart, add_slide_charts

__all__ = [
    "add_chart",
    "add_slide_charts",
]
