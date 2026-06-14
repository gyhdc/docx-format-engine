"""Template structure planning and explicit missing-section completion."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from .document_analyzer import analyze_document
from .document_io import (
    copy_docx_atomic,
    load_docx,
    prepare_output_path,
    save_docx_atomic,
)
from .models import Rules
from .ooxml_utils import apply_role_format
from .rule_loader import load_rules, resolve_role_rule

STRUCTURE_ROLES = {
    "abstract",
    "keywords",
    "acknowledgements",
    "appendix",
    "reference_heading",
}

STRUCTURE_ROLE_ALIASES = {
    "abstract_heading_zh": "abstract",
    "abstract_body_zh": "abstract",
    "abstract_heading_en": "abstract",
    "abstract_body_en": "abstract",
    "keywords_zh": "keywords",
    "keywords_en": "keywords",
}


def _structure_role(role: str) -> str:
    return STRUCTURE_ROLE_ALIASES.get(role, role)


def _analysis_rules() -> Rules:
    return Rules.model_validate({"version": "0.1", "roles": {"body": {}}})


def _outline(document: Any) -> list[dict[str, Any]]:
    analyzed = analyze_document(document, _analysis_rules())
    return [
        {
            "role": item.role,
            "text": item.text,
            "story": item.story,
            "location": item.location,
        }
        for item in analyzed
        if item.text.strip() and item.story == "body"
    ]


def plan_structure(
    template_path: str | Path,
    draft_path: str | Path,
) -> dict[str, Any]:
    """Compare semantic section roles without copying template content."""

    template, resolved_template = load_docx(template_path)
    draft, resolved_draft = load_docx(draft_path)
    template_outline = _outline(template)
    draft_outline = _outline(draft)
    template_roles = [
        _structure_role(item["role"])
        for item in template_outline
        if _structure_role(item["role"]) in STRUCTURE_ROLES
    ]
    draft_roles = {
        _structure_role(item["role"])
        for item in draft_outline
        if _structure_role(item["role"]) in STRUCTURE_ROLES
    }
    missing_roles: list[str] = []
    for role in template_roles:
        if role not in draft_roles and role not in missing_roles:
            missing_roles.append(role)
    return {
        "template": str(resolved_template),
        "draft": str(resolved_draft),
        "template_outline": template_outline,
        "draft_outline": draft_outline,
        "missing_roles": missing_roles,
        "not_automatically_copied": [
            "arbitrary_template_blocks",
            "template_body_text",
            "text_boxes_and_floating_objects",
        ],
    }


def complete_structure_to_path(
    docx_path: str | Path,
    rules_path: str | Path,
    output_path: str | Path,
) -> dict[str, Any]:
    """Insert only missing sections explicitly declared in rules."""

    document, source_path = load_docx(docx_path)
    rules = load_rules(rules_path)
    output = prepare_output_path(
        output_path, protected_input=source_path, suffix=".docx"
    )
    inserted_roles: list[str] = []

    for requirement in rules.structure.required_sections:
        analyzed = analyze_document(document, rules)
        if any(
            item.story == "body" and item.role == requirement.role
            for item in analyzed
        ):
            continue
        target = next(
            (
                document.paragraphs[item.paragraph_index]
                for item in analyzed
                if item.story == "body"
                and item.role == requirement.insert_before_role
            ),
            None,
        )
        if target is None:
            heading = document.add_paragraph(requirement.heading_text)
        else:
            heading = target.insert_paragraph_before(requirement.heading_text)
        role_rule = resolve_role_rule(rules, requirement.role)
        if role_rule is not None:
            apply_role_format(heading, role_rule)

        if requirement.placeholder_text:
            if target is None:
                placeholder = document.add_paragraph(
                    requirement.placeholder_text
                )
            else:
                placeholder = target.insert_paragraph_before(
                    requirement.placeholder_text
                )
            body_rule = resolve_role_rule(rules, "body")
            if body_rule is not None:
                apply_role_format(placeholder, body_rule)
        inserted_roles.append(requirement.role)

    if inserted_roles:
        save_docx_atomic(document, output)
    else:
        copy_docx_atomic(source_path, output)
    return {
        "source": str(source_path),
        "output": str(output),
        "inserted_roles": inserted_roles,
        "unchanged_existing_content": True,
        "notes": [
            "仅插入 rules.structure.required_sections 显式声明且当前缺失的章节。",
            "未复制模板正文、表格、文本框或浮动对象。",
        ],
    }
