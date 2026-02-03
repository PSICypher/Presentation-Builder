"""Presentation generator package — PPTX builder engine.

Consumes a TemplateSchema and data payload to produce PowerPoint files.
"""

from .pptx_builder import PPTXBuilder, build_presentation

__all__ = [
    "PPTXBuilder",
    "build_presentation",
]
