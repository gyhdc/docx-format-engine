"""Serialize inspection/fix reports as UTF-8 JSON and readable Markdown.

Purpose:
    Keep machine-facing and human-facing report output consistent.
MVP scope:
    Writes local files only. Markdown uses a compact issue table plus detailed
    explanations for unresolved issues.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .document_io import prepare_output_path
from .exceptions import ReportWriteError
from .models import FormatReport


def _as_dict(report: FormatReport | dict[str, Any]) -> dict[str, Any]:
    if isinstance(report, FormatReport):
        return report.model_dump(mode="json")
    return report


def write_json_report(
    report: FormatReport | dict[str, Any], output_path: str | Path,
) -> Path:
    """Write a report as indented UTF-8 JSON."""

    output = prepare_output_path(output_path, suffix=".json")
    try:
        output.write_text(
            json.dumps(_as_dict(report), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except (OSError, TypeError, ValueError) as exc:
        raise ReportWriteError(f"无法写入 JSON 报告 {output}: {exc}") from exc
    return output


def _markdown_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        text = json.dumps(value, ensure_ascii=False, sort_keys=True)
    else:
        text = str(value)
    return text.replace("|", r"\|").replace("\n", "<br>")


def render_markdown_report(
    report: FormatReport | dict[str, Any],
) -> str:
    """Render report data to Markdown without writing a file."""

    data = _as_dict(report)
    summary = data["summary"]
    lines = [
        "# Word 格式检查报告",
        "",
        "## 总览",
        "",
        f"- 文档：{data['document']}",
        f"- 规则：{data['rules']}",
        f"- 阶段：{data.get('phase', 'inspection')}",
        f"- 段落数：{summary['total_paragraphs']}",
        f"- 问题数：{summary['total_issues']}",
        f"- 可修复：{summary['fixable_issues']}",
        f"- 已修复：{summary['fixed_issues']}",
        f"- 未修复：{summary['unfixed_issues']}",
        "",
        "## 覆盖范围",
        "",
    ]
    coverage = data.get("coverage") or {}
    if coverage:
        lines.extend(
            [
                f"- 文档区域：{_markdown_cell(coverage.get('story_counts', {}))}",
                f"- 角色计数：{_markdown_cell(coverage.get('role_counts', {}))}",
                (
                    "- 受保护封面段落："
                    f"{coverage.get('protected_front_matter', 0)}"
                ),
                (
                    "- 视觉比较："
                    f"{coverage.get('visual_comparison', 'not_run')}"
                ),
                "- 未覆盖区域："
                + (
                    "、".join(coverage.get("uncovered_areas") or [])
                    or "无"
                ),
                "",
            ]
        )
    else:
        lines.extend(["- 未提供覆盖摘要。", ""])
    lines.extend(
        [
        "## 问题列表",
        "",
        "| ID | 严重级别 | 位置 | 角色 | 字段 | 期望 | 实际 | 可修复 | 已修复 |",
        "|---|---|---|---|---|---|---|---|---|",
        ]
    )
    for issue in data["issues"]:
        lines.append(
            "| "
            + " | ".join(
                [
                    _markdown_cell(issue["id"]),
                    _markdown_cell(issue["severity"]),
                    _markdown_cell(
                        issue.get("location") or issue["paragraph_index"]
                    ),
                    _markdown_cell(issue["role"]),
                    _markdown_cell(issue["field"]),
                    _markdown_cell(issue["expected"]),
                    _markdown_cell(issue["actual"]),
                    "是" if issue["fixable"] else "否",
                    "是" if issue.get("fixed") else "否",
                ]
            )
            + " |"
        )

    unresolved = [issue for issue in data["issues"] if not issue.get("fixed")]
    lines.extend(["", "## 未修复问题说明", ""])
    if not unresolved:
        lines.append("无。")
    else:
        for issue in unresolved:
            lines.append(f"- **{issue['id']}**：{issue['message']}")
            if issue.get("approximation"):
                lines.append(f"  - 近似说明：{issue['approximation']}")

    notes = data.get("notes") or []
    if notes:
        lines.extend(["", "## 说明", ""])
        lines.extend(f"- {note}" for note in notes)
    return "\n".join(lines) + "\n"


def write_markdown_report(
    report: FormatReport | dict[str, Any], output_path: str | Path,
) -> Path:
    """Write a report as UTF-8 Markdown."""

    output = prepare_output_path(output_path, suffix=".md")
    try:
        output.write_text(render_markdown_report(report), encoding="utf-8")
    except OSError as exc:
        raise ReportWriteError(f"无法写入 Markdown 报告 {output}: {exc}") from exc
    return output
