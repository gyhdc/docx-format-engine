"""Shared conversion constants and supported values.

Purpose:
    Keep approximate Word-unit conversions in one auditable place.
MVP scope:
    Character indentation is approximated from the paragraph's representative
    font size because OOXML does not expose a portable physical "character"
    unit for every document.
"""

from __future__ import annotations

POINTS_PER_INCH = 72.0
CM_PER_INCH = 2.54
CM_PER_POINT = CM_PER_INCH / POINTS_PER_INCH
DEFAULT_FONT_SIZE_PT = 12.0
FLOAT_TOLERANCE = 0.05
FONT_SIZE_TOLERANCE_PT = 0.1
SPACING_TOLERANCE_PT = 0.2
INDENT_TOLERANCE_CHARS = 0.15
MARGIN_TOLERANCE_CM = 0.05

SUPPORTED_ROLES = (
    "title",
    "abstract",
    "abstract_heading_zh",
    "abstract_body_zh",
    "abstract_heading_en",
    "abstract_body_en",
    "keywords",
    "keywords_zh",
    "keywords_en",
    "heading_1",
    "heading_2",
    "heading_3",
    "body",
    "toc_heading",
    "toc_entry_1",
    "toc_entry_2",
    "toc_entry_3",
    "figure_caption",
    "table_caption",
    "reference",
    "unknown",
)

FORMAT_ROLES = tuple(role for role in SUPPORTED_ROLES if role != "unknown")


def chars_to_cm(characters: float, font_size_pt: float | None) -> float:
    """Approximate a character count as centimeters for Word indentation."""

    size = font_size_pt or DEFAULT_FONT_SIZE_PT
    return characters * size * CM_PER_POINT


def cm_to_chars(centimeters: float, font_size_pt: float | None) -> float:
    """Approximate centimeters as a character count for report comparison."""

    size = font_size_pt or DEFAULT_FONT_SIZE_PT
    width_cm = size * CM_PER_POINT
    return centimeters / width_cm if width_cm else 0.0
