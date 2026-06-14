"""Fine-grained semantic module detection and selective template alignment."""

from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Any

from .document_analyzer import analyze_document
from .document_io import load_docx
from .exceptions import InputFileError
from .format_fixer import fix_document_with_rules_to_path
from .models import Rules
from .template_rules import MODULE_ROLE_GROUPS, build_template_rules_model


def _analysis_rules() -> Rules:
    return Rules.model_validate(
        {
            "version": "0.1",
            "roles": {
                "body": {},
                "table_of_contents": {},
            },
        }
    )


def _module_for_role(role: str) -> str | None:
    for module, roles in MODULE_ROLE_GROUPS.items():
        if role in roles:
            return module
    return None


def detect_modules(document_path: str | Path) -> dict[str, Any]:
    """Report semantic module ranges without modifying the document."""

    document, source = load_docx(document_path)
    analyzed = analyze_document(document, _analysis_rules())
    modules: list[dict[str, Any]] = []
    active: dict[str, Any] | None = None

    for item in analyzed:
        module = _module_for_role(item.role)
        if module is None:
            active = None
            continue
        if (
            active is None
            or active["name"] != module
            or active["story"] != item.story
        ):
            active = {
                "name": module,
                "story": item.story,
                "start": item.location,
                "end": item.location,
                "start_paragraph": item.paragraph_index,
                "end_paragraph": item.paragraph_index,
                "roles": [],
                "paragraph_count": 0,
                "text_preview": [],
            }
            modules.append(active)
        active["end"] = item.location
        active["end_paragraph"] = item.paragraph_index
        active["paragraph_count"] += 1
        if item.role not in active["roles"]:
            active["roles"].append(item.role)
        if item.text.strip() and len(active["text_preview"]) < 3:
            active["text_preview"].append(item.text[:100])

    counts = Counter(module["name"] for module in modules)
    return {
        "document": str(source),
        "modules": modules,
        "module_counts": dict(sorted(counts.items())),
        "available_module_names": sorted(MODULE_ROLE_GROUPS),
    }


def compare_modules(
    template_path: str | Path,
    draft_path: str | Path,
) -> dict[str, Any]:
    """Compare detected template and draft modules without changing either file."""

    template = detect_modules(template_path)
    draft = detect_modules(draft_path)
    template_names = set(template["module_counts"])
    draft_names = set(draft["module_counts"])
    return {
        "template": template["document"],
        "draft": draft["document"],
        "template_module_counts": template["module_counts"],
        "draft_module_counts": draft["module_counts"],
        "missing_in_draft": sorted(template_names - draft_names),
        "extra_in_draft": sorted(draft_names - template_names),
        "shared_modules": sorted(template_names & draft_names),
    }


def align_modules_from_template(
    template_path: str | Path,
    draft_path: str | Path,
    output_path: str | Path,
    *,
    modules: list[str] | tuple[str, ...],
    include_page: bool = False,
) -> dict[str, Any]:
    """Apply only explicitly selected template modules to a draft."""

    selected = list(dict.fromkeys(modules))
    if not selected:
        raise InputFileError("至少指定一个待对齐模块。")
    try:
        rules = build_template_rules_model(
            template_path,
            modules=selected,
            include_page=include_page,
        )
    except ValueError as exc:
        raise InputFileError(str(exc)) from exc
    report = fix_document_with_rules_to_path(
        draft_path,
        rules,
        output_path,
        rules_label=f"template-modules:{Path(template_path)}",
        repair_references="references" in selected,
        repair_toc="toc" in selected,
    ).model_dump(mode="json")
    report["template"] = str(template_path)
    report["selected_modules"] = selected
    report["include_page"] = include_page
    report["template_rule_roles"] = sorted(rules.roles)
    return report
