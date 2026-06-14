"""Generate module-aware format rules directly from a DOCX template."""

from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Any

from .constants import cm_to_chars
from .document_analyzer import analyze_document
from .document_io import load_docx
from .exceptions import InputFileError
from .models import RoleName, Rules
from .ooxml_utils import section_to_dict

MODULE_ROLE_GROUPS: dict[str, set[RoleName]] = {
    "cover": {"title", "title_zh", "title_en", "cover_label", "cover_value"},
    "abstract_zh": {
        "abstract_heading_zh",
        "abstract_body_zh",
        "keywords_zh",
    },
    "abstract_en": {
        "abstract_heading_en",
        "abstract_body_en",
        "keywords_en",
    },
    "toc": {
        "table_of_contents",
        "toc_heading",
        "toc_entry_1",
        "toc_entry_2",
        "toc_entry_3",
    },
    "headings": {"heading_1", "heading_2", "heading_3"},
    "body": {"body"},
    "captions": {"figure_caption", "table_caption"},
    "references": {
        "reference_heading",
        "reference_entry",
        "reference_entry_zh",
        "reference_entry_en",
        "reference",
    },
    "acknowledgements": {"acknowledgements"},
    "appendix": {"appendix"},
    "tables": {"table_text"},
    "headers_footers": {"header", "footer"},
}

ROLE_RULE_FIELDS = {
    "font_east_asia",
    "font_ascii",
    "font_size_pt",
    "bold",
    "italic",
    "alignment",
    "line_spacing",
    "left_indent_cm",
    "space_before_pt",
    "space_after_pt",
    "keep_with_next",
}


def _analysis_rules() -> Rules:
    return Rules.model_validate({"version": "0.1", "roles": {"body": {}}})


def _rule_from_format(format_data: dict[str, Any]) -> dict[str, Any]:
    rule = {
        field: format_data[field]
        for field in ROLE_RULE_FIELDS
        if format_data.get(field) is not None
    }
    line_spacing = format_data.get("line_spacing")
    if (
        format_data.get("line_spacing_kind") != "multiple"
        or not isinstance(line_spacing, (int, float))
        or not 0.5 <= float(line_spacing) <= 4.0
    ):
        rule.pop("line_spacing", None)
    font_size = format_data.get("font_size_pt")
    first_indent = float(format_data.get("first_line_indent_cm") or 0)
    hanging_indent = float(format_data.get("hanging_indent_cm") or 0)
    if first_indent > 0.01:
        rule["first_line_indent_chars"] = round(
            cm_to_chars(first_indent, font_size), 3
        )
    elif hanging_indent > 0.01:
        rule["hanging_indent_chars"] = round(
            cm_to_chars(hanging_indent, font_size), 3
        )
    return rule


def _representative_format(formats: list[dict[str, Any]]) -> dict[str, Any]:
    signatures = [
        tuple(sorted(_rule_from_format(format_data).items()))
        for format_data in formats
    ]
    signature, _ = Counter(signatures).most_common(1)[0]
    return dict(signature)


def _selected_roles(modules: list[str] | tuple[str, ...] | None) -> set[RoleName] | None:
    if not modules:
        return None
    unknown = sorted(set(modules) - MODULE_ROLE_GROUPS.keys())
    if unknown:
        raise InputFileError(f"不支持的模块名称: {', '.join(unknown)}")
    return {
        role
        for module in modules
        for role in MODULE_ROLE_GROUPS[module]
    }


def build_template_rules_model(
    template_path: str | Path,
    *,
    modules: list[str] | tuple[str, ...] | None = None,
    include_page: bool = True,
) -> Rules:
    """Infer page and per-module paragraph rules from one template."""

    document, _ = load_docx(template_path)
    analyzed = analyze_document(document, _analysis_rules())
    selected_roles = _selected_roles(modules)
    formats_by_role: dict[RoleName, list[dict[str, Any]]] = {}
    for item in analyzed:
        if not item.text.strip() or item.role == "unknown":
            continue
        if selected_roles is not None and item.role not in selected_roles:
            continue
        formats_by_role.setdefault(item.role, []).append(item.format)

    roles: dict[str, dict[str, Any]] = {
        role: _representative_format(formats)
        for role, formats in formats_by_role.items()
        if formats
    }
    if not roles:
        roles["body"] = {}
    first_section = section_to_dict(document.sections[0])
    page = {
        field: first_section[field]
        for field in (
            "page_width_cm",
            "page_height_cm",
            "margin_top_cm",
            "margin_bottom_cm",
            "margin_left_cm",
            "margin_right_cm",
            "header_distance_cm",
            "footer_distance_cm",
        )
    }
    return Rules.model_validate(
        {
            "version": "0.1",
            "page": page if include_page else None,
            "roles": roles,
        }
    )


def build_template_rules(
    template_path: str | Path,
    *,
    modules: list[str] | tuple[str, ...] | None = None,
    include_page: bool = True,
) -> dict[str, Any]:
    """Return JSON-compatible module-aware rules inferred from a template."""

    return build_template_rules_model(
        template_path,
        modules=modules,
        include_page=include_page,
    ).model_dump(
        mode="json",
        exclude_none=True,
    )
