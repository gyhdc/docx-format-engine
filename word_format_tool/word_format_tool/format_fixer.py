"""Apply supported role/page rules and correlate before/after inspection issues.

Purpose:
    Create a new formatted DOCX and report which original issues disappeared.
MVP scope:
    Applies margins, paragraph properties, and run font properties. It never
    moves captions, objects, formulas, tables, or paragraphs, and it skips the
    ``unknown`` role entirely.
"""

from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path
from typing import Any

from .document_analyzer import analyze_document, resolve_analyzed_paragraph
from .document_io import (
    copy_docx_atomic,
    load_docx,
    prepare_output_path,
    save_docx_atomic,
    validate_preserved_content,
)
from .format_inspector import inspect_analyzed_document
from .models import FormatIssue, FormatReport, ReportSummary, Rules
from .ooxml_utils import apply_page_rules, apply_role_format
from .reference_links import (
    analyze_reference_document,
    apply_reference_navigation,
)
from .rule_loader import load_rules, resolve_role_rule
from .toc_repair import repair_toc_navigation


def _issue_signature(issue: FormatIssue) -> tuple[Any, ...]:
    return (
        issue.paragraph_index,
        issue.location,
        issue.role,
        issue.field,
        json.dumps(issue.expected, ensure_ascii=False, sort_keys=True),
    )


def _build_fix_report(
    before: FormatReport,
    after: FormatReport,
    *,
    source_path: Path,
    rules_path: str | Path,
    output: Path,
    total_paragraphs: int,
    notes: list[str],
) -> FormatReport:
    after_signatures = {_issue_signature(issue) for issue in after.issues}
    merged: list[FormatIssue] = []
    original_signatures: set[tuple[Any, ...]] = set()
    for issue in before.issues:
        signature = _issue_signature(issue)
        original_signatures.add(signature)
        issue.fixed = issue.fixable and signature not in after_signatures
        merged.append(issue)

    for issue in after.issues:
        if _issue_signature(issue) not in original_signatures:
            issue.id = f"ISSUE-{len(merged) + 1:04d}"
            merged.append(issue)

    fixed_count = sum(issue.fixed for issue in merged)
    fixable_count = sum(issue.fixable for issue in merged)
    return FormatReport(
        document=str(source_path),
        rules=str(rules_path),
        phase="fix",
        summary=ReportSummary(
            total_paragraphs=total_paragraphs,
            total_issues=len(merged),
            fixable_issues=fixable_count,
            fixed_issues=fixed_count,
            unfixed_issues=len(merged) - fixed_count,
        ),
        issues=merged,
        notes=[f"修复结果写入: {output}", *notes],
        coverage=after.coverage,
    )


def fix_document_with_rules_to_path(
    docx_path: str | Path,
    rules: Rules,
    output_path: str | Path,
    *,
    rules_label: str | Path = "<template-derived>",
    repair_references: bool = True,
    repair_toc: bool = True,
) -> FormatReport:
    """Fix supported issues with an already validated in-memory rules model."""

    document, source_path = load_docx(docx_path)
    output = prepare_output_path(
        output_path, protected_input=source_path, suffix=".docx"
    )

    before_analysis = analyze_document(document, rules)
    before = inspect_analyzed_document(
        document, before_analysis, rules, source_path, rules_label
    )

    fixable_issues = [issue for issue in before.issues if issue.fixable]
    if not fixable_issues:
        copy_docx_atomic(source_path, output)
    else:
        page_fields = {
            issue.field
            for issue in fixable_issues
            if issue.paragraph_index is None and issue.role == "document"
        }
        if page_fields:
            apply_page_rules(document, rules.page, page_fields)

        paragraph_fields: dict[str, set[str]] = defaultdict(set)
        for issue in fixable_issues:
            if issue.paragraph_index is not None:
                location = issue.location or f"body.paragraph[{issue.paragraph_index}]"
                paragraph_fields[location].add(issue.field)
        for item in before_analysis:
            fields = paragraph_fields.get(item.location)
            if item.role == "unknown" or not fields:
                continue
            role_rule = resolve_role_rule(rules, item.role)
            if role_rule is not None:
                apply_role_format(
                    resolve_analyzed_paragraph(document, item),
                    role_rule,
                    fields,
                )

        if repair_references:
            reference_map = analyze_reference_document(document, before_analysis)
            apply_reference_navigation(document, reference_map)
        if repair_toc:
            repair_toc_navigation(document, before_analysis)
        save_docx_atomic(
            document,
            output,
            validator=lambda candidate: validate_preserved_content(
                source_path, candidate
            ),
        )

    fixed_document, _ = load_docx(output)
    after_analysis = analyze_document(fixed_document, rules)
    after = inspect_analyzed_document(
        fixed_document, after_analysis, rules, output, rules_label
    )
    return _build_fix_report(
        before,
        after,
        source_path=source_path,
        rules_path=rules_label,
        output=output,
        total_paragraphs=len(before_analysis),
        notes=[
            "未改动文字内容、图片、表格结构、公式或段落顺序。",
            "表格单元格、页眉和页脚仅在存在显式角色规则时修改格式。",
            "图题和表题位置问题只检测，不自动移动。",
            "目录书签只在目录项能唯一匹配正文标题时补建，并设置打开时更新域。",
        ],
    )


def fix_document_to_path(
    docx_path: str | Path,
    rules_path: str | Path,
    output_path: str | Path,
) -> FormatReport:
    """Fix supported issues from a rules.json file."""

    return fix_document_with_rules_to_path(
        docx_path,
        load_rules(rules_path),
        output_path,
        rules_label=rules_path,
    )


def link_references_to_path(
    docx_path: str | Path,
    rules_path: str | Path,
    output_path: str | Path,
) -> FormatReport:
    """Add only reference navigation and superscript citation formatting."""

    document, source_path = load_docx(docx_path)
    rules = load_rules(rules_path)
    output = prepare_output_path(
        output_path, protected_input=source_path, suffix=".docx"
    )
    before_analysis = analyze_document(document, rules)
    before = inspect_analyzed_document(
        document, before_analysis, rules, source_path, rules_path
    )
    reference_map = analyze_reference_document(document, before_analysis)
    navigation_result = apply_reference_navigation(document, reference_map)
    if navigation_result.changed:
        save_docx_atomic(
            document,
            output,
            validator=lambda candidate: validate_preserved_content(
                source_path, candidate
            ),
        )
    else:
        copy_docx_atomic(source_path, output)

    linked_document, _ = load_docx(output)
    after_analysis = analyze_document(linked_document, rules)
    after = inspect_analyzed_document(
        linked_document, after_analysis, rules, output, rules_path
    )
    return _build_fix_report(
        before,
        after,
        source_path=source_path,
        rules_path=rules_path,
        output=output,
        total_paragraphs=len(before_analysis),
        notes=[
            "仅处理参考文献书签、正文上标和内部跳转。",
            "未应用页边距、字体、行距、缩进等普通格式规则。",
            "未改动可见文字、图片、表格、公式或段落顺序。",
        ],
    )
