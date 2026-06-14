"""Stable Python API for deterministic Word profiling, repair, and validation."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from .conversion import convert_to_docx as _convert_to_docx
from .document_analyzer import analyze_document
from .document_io import load_docx
from .format_fixer import fix_document_to_path, link_references_to_path
from .format_inspector import inspect_analyzed_document
from .layout_validator import validate_layout as _validate_layout
from .models import Rules
from .module_tools import (
    align_modules_from_template as _align_modules_from_template,
)
from .module_tools import compare_modules as _compare_modules
from .module_tools import detect_modules as _detect_modules
from .rule_loader import load_rules, validate_rules_file
from .optional_word_backend import refresh_fields_to_path
from .structure_inspector import inspect_structure as _inspect_structure
from .structure_operations import apply_structure_operations_to_path
from .structure_tools import (
    complete_structure_to_path,
)
from .structure_tools import (
    plan_structure as _plan_structure,
)
from .template_profiler import profile_template as _profile_template
from .template_rules import build_template_rules as _build_template_rules
from .visual_compare import compare_docx_visual
from .workflow import format_from_template as _format_from_template

__version__ = "0.6.1"


def rules_schema() -> dict[str, Any]:
    """Return the strict JSON Schema used for Agent-generated rules."""

    return Rules.model_json_schema()


def profile_template(template_path: str | Path) -> dict[str, Any]:
    """Extract a template profile suitable for Agent rule generation."""

    return _profile_template(template_path)


def build_template_rules(
    template_path: str | Path,
    *,
    modules: list[str] | tuple[str, ...] | None = None,
    include_page: bool = True,
) -> dict[str, Any]:
    """Infer module-aware rules directly from a DOCX template."""

    return _build_template_rules(
        template_path,
        modules=modules,
        include_page=include_page,
    )


def detect_modules(document_path: str | Path) -> dict[str, Any]:
    """Detect semantic module boundaries without modifying the document."""

    return _detect_modules(document_path)


def compare_modules(
    template_path: str | Path,
    draft_path: str | Path,
) -> dict[str, Any]:
    """Compare module presence before an Agent chooses an operation."""

    return _compare_modules(template_path, draft_path)


def align_modules_from_template(
    template_path: str | Path,
    draft_path: str | Path,
    output_path: str | Path,
    *,
    modules: list[str] | tuple[str, ...],
    include_page: bool = False,
) -> dict[str, Any]:
    """Apply only selected template modules to a draft."""

    return _align_modules_from_template(
        template_path,
        draft_path,
        output_path,
        modules=modules,
        include_page=include_page,
    )


def convert_to_docx(
    input_path: str | Path,
    output_path: str | Path,
) -> dict[str, str]:
    """Convert legacy .doc input or copy .docx input to a distinct output."""

    return _convert_to_docx(input_path, output_path)


def format_from_template(
    template_path: str | Path,
    draft_path: str | Path,
    output_path: str | Path,
) -> dict[str, Any]:
    """Run conversion, module-aware rule inference, and formatting in one call."""

    return _format_from_template(template_path, draft_path, output_path)


def validate_rules(rules_path: str | Path) -> bool:
    """Return ``True`` when a rules file passes strict validation."""

    return validate_rules_file(rules_path)


def inspect_document(
    docx_path: str | Path,
    rules_path: str | Path,
) -> dict[str, Any]:
    """Inspect a DOCX without modifying it and return JSON-compatible data."""

    document, source_path = load_docx(docx_path)
    rules = load_rules(rules_path)
    analyzed = analyze_document(document, rules)
    report = inspect_analyzed_document(
        document, analyzed, rules, source_path, rules_path
    )
    return report.model_dump(mode="json")


def fix_document(
    docx_path: str | Path,
    rules_path: str | Path,
    output_path: str | Path,
) -> dict[str, Any]:
    """Write a fixed DOCX to a distinct path and return the fix report."""

    return fix_document_to_path(
        docx_path, rules_path, output_path
    ).model_dump(mode="json")


def link_references(
    docx_path: str | Path,
    rules_path: str | Path,
    output_path: str | Path,
) -> dict[str, Any]:
    """Add reference navigation without applying unrelated format rules."""

    return link_references_to_path(
        docx_path, rules_path, output_path
    ).model_dump(mode="json")


def plan_structure(
    template_path: str | Path,
    draft_path: str | Path,
) -> dict[str, Any]:
    """Report semantic sections present in a template but missing in a draft."""

    return _plan_structure(template_path, draft_path)


def complete_structure(
    docx_path: str | Path,
    rules_path: str | Path,
    output_path: str | Path,
) -> dict[str, Any]:
    """Insert only missing sections explicitly declared in rules.json."""

    return complete_structure_to_path(docx_path, rules_path, output_path)


def inspect_structure(docx_path: str | Path) -> dict[str, Any]:
    """Return sections, fields, TOC boundaries, and stable paragraph locators."""

    return _inspect_structure(docx_path)


def apply_structure_operations(
    docx_path: str | Path,
    plan_path: str | Path,
    output_path: str | Path,
) -> dict[str, Any]:
    """Apply one strict, idempotent structural operation plan."""

    return apply_structure_operations_to_path(docx_path, plan_path, output_path)


def validate_layout(
    docx_path: str | Path,
    expectations_path: str | Path,
) -> dict[str, Any]:
    """Validate machine-readable structural layout expectations."""

    return _validate_layout(docx_path, expectations_path)


def refresh_fields(
    docx_path: str | Path,
    output_path: str | Path,
) -> dict[str, Any]:
    """Refresh Word fields and pagination through the optional Word backend."""

    return refresh_fields_to_path(docx_path, output_path)


def compare_visual(
    source_docx: str | Path,
    candidate_docx: str | Path,
    output_dir: str | Path,
    *,
    dpi: int = 120,
    changed_pixel_threshold: float = 0.005,
) -> dict[str, Any]:
    """Optionally render two DOCX files with Word and compare their pages."""

    return compare_docx_visual(
        source_docx,
        candidate_docx,
        output_dir,
        dpi=dpi,
        changed_pixel_threshold=changed_pixel_threshold,
    )


__all__ = [
    "align_modules_from_template",
    "apply_structure_operations",
    "build_template_rules",
    "compare_visual",
    "compare_modules",
    "complete_structure",
    "convert_to_docx",
    "detect_modules",
    "fix_document",
    "format_from_template",
    "inspect_document",
    "inspect_structure",
    "link_references",
    "plan_structure",
    "profile_template",
    "refresh_fields",
    "rules_schema",
    "validate_layout",
    "validate_rules",
]
