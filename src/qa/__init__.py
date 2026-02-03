"""QA validation package for Presentation Builder.

Validates generated PPTX output against template schemas — checks slide
count, data slot population, formatting rules, table row counts, chart
series, and variance coloring.
"""

from .validator import (
    Issue,
    QAResult,
    QAValidator,
    validate_presentation,
)

__all__ = [
    "Issue",
    "QAResult",
    "QAValidator",
    "validate_presentation",
]
