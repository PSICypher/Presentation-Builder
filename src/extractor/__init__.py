"""Template extraction engine — derives TemplateSchema from PPTX templates.

Uses the TemplateAnalyzer output to classify slides, identify data-bearing
shapes, and produce a typed TemplateSchema for the generator.
"""

from .template_extractor import TemplateExtractor, extract_template

__all__ = ["TemplateExtractor", "extract_template"]
