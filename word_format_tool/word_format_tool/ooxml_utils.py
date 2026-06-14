"""Read and apply effective paragraph/run formatting, including OOXML fonts.

Purpose:
    Centralize python-docx and low-level OOXML handling so analyzer, profiler,
    inspector, and fixer use the same interpretation of Word formatting.
MVP scope:
    A paragraph's representative run is the longest non-whitespace text run.
    Character-based indentation uses the documented approximation in
    ``constants.py``. This module never moves or deletes document content.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import Any

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from .constants import DEFAULT_FONT_SIZE_PT, chars_to_cm
from .models import PageRules, RoleFormatRule

ALIGNMENT_TO_NAME = {
    WD_ALIGN_PARAGRAPH.LEFT: "left",
    WD_ALIGN_PARAGRAPH.CENTER: "center",
    WD_ALIGN_PARAGRAPH.RIGHT: "right",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "both",
    WD_ALIGN_PARAGRAPH.DISTRIBUTE: "distribute",
}
NAME_TO_ALIGNMENT = {value: key for key, value in ALIGNMENT_TO_NAME.items()}


def _style_chain(style: Any) -> Iterator[Any]:
    seen: set[str] = set()
    current = style
    while current is not None:
        style_id = getattr(current, "style_id", repr(current))
        if style_id in seen:
            break
        seen.add(style_id)
        yield current
        current = getattr(current, "base_style", None)


def _effective_paragraph_property(paragraph: Any, name: str) -> Any:
    direct = getattr(paragraph.paragraph_format, name)
    if direct is not None:
        return direct
    for style in _style_chain(paragraph.style):
        value = getattr(style.paragraph_format, name)
        if value is not None:
            return value
    return None


def _rfonts_value(source: Any, attribute: str) -> str | None:
    element = getattr(source, "_element", None)
    if element is None:
        return None
    rpr = getattr(element, "rPr", None)
    if rpr is None or rpr.rFonts is None:
        return None
    return rpr.rFonts.get(qn(f"w:{attribute}"))


def _effective_run_property(run: Any, paragraph: Any, name: str) -> Any:
    value = getattr(run.font, name)
    if value is not None:
        return value
    run_style = getattr(run, "style", None)
    if run_style is not None:
        for style in _style_chain(run_style):
            value = getattr(style.font, name)
            if value is not None:
                return value
    for style in _style_chain(paragraph.style):
        value = getattr(style.font, name)
        if value is not None:
            return value
    return None


def _effective_run_font_name(
    run: Any, paragraph: Any, attribute: str,
) -> str | None:
    value = _rfonts_value(run, attribute)
    if value is not None:
        return value
    run_style = getattr(run, "style", None)
    if run_style is not None:
        for style in _style_chain(run_style):
            value = _rfonts_value(style, attribute)
            if value is not None:
                return value
    for style in _style_chain(paragraph.style):
        value = _rfonts_value(style, attribute)
        if value is not None:
            return value
    if attribute in {"ascii", "hAnsi"}:
        return _effective_run_property(run, paragraph, "name")
    return None


def _representative_run(paragraph: Any) -> Any | None:
    candidates = [run for run in paragraph.runs if run.text.strip()]
    if not candidates:
        return None
    return max(candidates, key=lambda run: len(run.text.strip()))


def _length_cm(value: Any) -> float | None:
    return None if value is None else round(value.cm, 4)


def _length_pt(value: Any) -> float | None:
    return None if value is None else round(value.pt, 4)


def set_run_fonts(
    run: Any,
    east_asia_font: str | None,
    ascii_font: str | None,
) -> None:
    """Set East Asian, ASCII, and high-ANSI fonts on one run."""

    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    if east_asia_font:
        rfonts.set(qn("w:eastAsia"), east_asia_font)
    if ascii_font:
        rfonts.set(qn("w:ascii"), ascii_font)
        rfonts.set(qn("w:hAnsi"), ascii_font)
        run.font.name = ascii_font


def get_paragraph_format(paragraph: Any) -> dict[str, Any]:
    """Return normalized effective formatting for one paragraph."""

    run = _representative_run(paragraph)
    font_size = (
        _effective_run_property(run, paragraph, "size") if run is not None else None
    )
    font_size_pt = (
        round(font_size.pt, 4) if font_size is not None else DEFAULT_FONT_SIZE_PT
    )

    alignment = _effective_paragraph_property(paragraph, "alignment")
    line_spacing = _effective_paragraph_property(paragraph, "line_spacing")
    if line_spacing is None:
        normalized_line_spacing: float | None = 1.0
        line_spacing_kind = "multiple"
    elif isinstance(line_spacing, (int, float)):
        normalized_line_spacing = round(float(line_spacing), 4)
        line_spacing_kind = "multiple"
    else:
        normalized_line_spacing = round(line_spacing.pt, 4)
        line_spacing_kind = "points"

    first_indent = _effective_paragraph_property(
        paragraph, "first_line_indent"
    )
    first_indent_cm = _length_cm(first_indent)
    positive_first_indent = max(first_indent_cm or 0.0, 0.0)
    hanging_indent = abs(min(first_indent_cm or 0.0, 0.0))

    if run is None:
        font_east_asia = None
        font_ascii = None
        bold = False
        italic = False
    else:
        font_east_asia = _effective_run_font_name(run, paragraph, "eastAsia")
        font_ascii = _effective_run_font_name(run, paragraph, "ascii")
        bold = bool(_effective_run_property(run, paragraph, "bold") or False)
        italic = bool(_effective_run_property(run, paragraph, "italic") or False)

    return {
        "font_east_asia": font_east_asia,
        "font_ascii": font_ascii,
        "font_size_pt": font_size_pt,
        "bold": bold,
        "italic": italic,
        "alignment": ALIGNMENT_TO_NAME.get(alignment, "left"),
        "line_spacing": normalized_line_spacing,
        "line_spacing_kind": line_spacing_kind,
        "left_indent_cm": _length_cm(
            _effective_paragraph_property(paragraph, "left_indent")
        )
        or 0.0,
        "first_line_indent_cm": round(positive_first_indent, 4),
        "hanging_indent_cm": round(hanging_indent, 4),
        "space_before_pt": _length_pt(
            _effective_paragraph_property(paragraph, "space_before")
        )
        or 0.0,
        "space_after_pt": _length_pt(
            _effective_paragraph_property(paragraph, "space_after")
        )
        or 0.0,
        "keep_with_next": bool(
            _effective_paragraph_property(paragraph, "keep_with_next") or False
        ),
        "representative_run_strategy": "longest_non_whitespace_run",
    }


def get_style_format(style: Any) -> dict[str, Any]:
    """Return the direct format stored on a Word style."""

    paragraph_format = getattr(style, "paragraph_format", None)
    font = getattr(style, "font", None)
    alignment = (
        getattr(paragraph_format, "alignment", None)
        if paragraph_format is not None
        else None
    )
    line_spacing = (
        getattr(paragraph_format, "line_spacing", None)
        if paragraph_format is not None
        else None
    )
    if isinstance(line_spacing, (int, float)):
        normalized_line_spacing: float | None = round(float(line_spacing), 4)
    elif line_spacing is not None:
        normalized_line_spacing = round(line_spacing.pt, 4)
    else:
        normalized_line_spacing = None

    return {
        "font_east_asia": _rfonts_value(style, "eastAsia"),
        "font_ascii": _rfonts_value(style, "ascii")
        or (font.name if font is not None else None),
        "font_size_pt": (
            round(font.size.pt, 4)
            if font is not None and font.size is not None
            else None
        ),
        "bold": font.bold if font is not None else None,
        "italic": font.italic if font is not None else None,
        "alignment": ALIGNMENT_TO_NAME.get(alignment),
        "left_indent_cm": (
            _length_cm(paragraph_format.left_indent)
            if paragraph_format is not None
            else None
        ),
        "first_line_indent_cm": (
            _length_cm(paragraph_format.first_line_indent)
            if paragraph_format is not None
            else None
        ),
        "line_spacing": normalized_line_spacing,
        "space_before_pt": (
            _length_pt(paragraph_format.space_before)
            if paragraph_format is not None
            else None
        ),
        "space_after_pt": (
            _length_pt(paragraph_format.space_after)
            if paragraph_format is not None
            else None
        ),
        "keep_with_next": (
            paragraph_format.keep_with_next
            if paragraph_format is not None
            else None
        ),
    }


def apply_role_format(
    paragraph: Any,
    rule: RoleFormatRule,
    fields: set[str] | None = None,
) -> None:
    """Apply selected role fields without changing text or object order."""

    paragraph_format = paragraph.paragraph_format

    def selected(field: str) -> bool:
        return fields is None or field in fields

    if selected("alignment") and rule.alignment is not None:
        paragraph.alignment = NAME_TO_ALIGNMENT[rule.alignment]
    if selected("line_spacing") and rule.line_spacing is not None:
        paragraph_format.line_spacing = rule.line_spacing
    if selected("left_indent_cm") and rule.left_indent_cm is not None:
        paragraph_format.left_indent = Cm(rule.left_indent_cm)
    if (
        selected("first_line_indent_chars")
        and rule.first_line_indent_chars is not None
    ):
        indent_cm = chars_to_cm(
            rule.first_line_indent_chars, rule.font_size_pt
        )
        paragraph_format.first_line_indent = Cm(indent_cm)
    if (
        selected("hanging_indent_chars")
        and rule.hanging_indent_chars is not None
    ):
        indent_cm = chars_to_cm(
            rule.hanging_indent_chars, rule.font_size_pt
        )
        paragraph_format.first_line_indent = Cm(-indent_cm)
    if selected("space_before_pt") and rule.space_before_pt is not None:
        paragraph_format.space_before = Pt(rule.space_before_pt)
    if selected("space_after_pt") and rule.space_after_pt is not None:
        paragraph_format.space_after = Pt(rule.space_after_pt)
    if selected("keep_with_next") and rule.keep_with_next is not None:
        paragraph_format.keep_with_next = rule.keep_with_next

    for run in paragraph.runs:
        east_asia_font = (
            rule.font_east_asia if selected("font_east_asia") else None
        )
        ascii_font = rule.font_ascii if selected("font_ascii") else None
        if east_asia_font is not None or ascii_font is not None:
            set_run_fonts(run, east_asia_font, ascii_font)
        if selected("font_size_pt") and rule.font_size_pt is not None:
            run.font.size = Pt(rule.font_size_pt)
        if selected("bold") and rule.bold is not None:
            run.bold = rule.bold
        if selected("italic") and rule.italic is not None:
            run.italic = rule.italic


def section_to_dict(section: Any) -> dict[str, float]:
    """Extract page geometry from one Word section."""

    return {
        "page_width_cm": round(section.page_width.cm, 4),
        "page_height_cm": round(section.page_height.cm, 4),
        "margin_top_cm": round(section.top_margin.cm, 4),
        "margin_bottom_cm": round(section.bottom_margin.cm, 4),
        "margin_left_cm": round(section.left_margin.cm, 4),
        "margin_right_cm": round(section.right_margin.cm, 4),
        "header_distance_cm": round(section.header_distance.cm, 4),
        "footer_distance_cm": round(section.footer_distance.cm, 4),
    }


def apply_page_rules(
    document: Any,
    page: PageRules | None,
    fields: set[str] | None = None,
) -> None:
    """Apply selected page settings to every section."""

    if page is None:
        return

    def selected(field: str) -> bool:
        return fields is None or field in fields

    paper_sizes = {
        "A4": (21.0, 29.7),
        "Letter": (21.59, 27.94),
        "Legal": (21.59, 35.56),
    }
    for section in document.sections:
        if page.paper_size is not None:
            width, height = paper_sizes[page.paper_size]
            if selected("page_width_cm"):
                section.page_width = Cm(width)
            if selected("page_height_cm"):
                section.page_height = Cm(height)
        if selected("page_width_cm") and page.page_width_cm is not None:
            section.page_width = Cm(page.page_width_cm)
        if selected("page_height_cm") and page.page_height_cm is not None:
            section.page_height = Cm(page.page_height_cm)
        if selected("margin_top_cm") and page.margin_top_cm is not None:
            section.top_margin = Cm(page.margin_top_cm)
        if selected("margin_bottom_cm") and page.margin_bottom_cm is not None:
            section.bottom_margin = Cm(page.margin_bottom_cm)
        if selected("margin_left_cm") and page.margin_left_cm is not None:
            section.left_margin = Cm(page.margin_left_cm)
        if selected("margin_right_cm") and page.margin_right_cm is not None:
            section.right_margin = Cm(page.margin_right_cm)
        if (
            selected("header_distance_cm")
            and page.header_distance_cm is not None
        ):
            section.header_distance = Cm(page.header_distance_cm)
        if (
            selected("footer_distance_cm")
            and page.footer_distance_cm is not None
        ):
            section.footer_distance = Cm(page.footer_distance_cm)
