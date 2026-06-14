"""Typer command-line interface exposed as the ``wordfmt`` executable.

Purpose:
    Provide seven deterministic commands for template profiling, rules
    validation, inspection, repair, dry-run preview, and sample extraction.
MVP scope:
    Local synchronous file operations only; no network, GUI, backend, or model
    calls are performed.
"""

from __future__ import annotations

from pathlib import Path
from typing import NoReturn

import typer
from rich.console import Console

from . import (
    align_modules_from_template,
    apply_structure_operations,
    build_template_rules,
    compare_modules,
    compare_visual,
    complete_structure,
    convert_to_docx,
    detect_modules,
    fix_document,
    format_from_template,
    inspect_document,
    inspect_structure,
    link_references,
    plan_structure,
    profile_template,
    refresh_fields,
    validate_layout,
    validate_rules,
)
from .exceptions import WordFormatToolError
from .report_writer import write_json_report, write_markdown_report
from .template_profiler import extract_samples

app = typer.Typer(
    name="wordfmt",
    help="Word 论文格式检查与安全修复工具。",
    no_args_is_help=True,
)
console = Console()


def _abort(exc: WordFormatToolError) -> NoReturn:
    console.print(f"[bold red]错误:[/bold red] {exc}")
    raise typer.Exit(code=1)


@app.command("profile-template")
def profile_template_command(
    template: Path = typer.Argument(..., help="模板 DOCX 路径"),
    output: Path = typer.Option(
        ..., "-o", "--output", help="template_profile.json 输出路径"
    ),
) -> None:
    """Extract page, style, and paragraph sample data from a template."""

    try:
        profile = profile_template(template)
        write_json_report(profile, output)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(f"[green]模板 profile 已写入:[/green] {output}")


@app.command("convert")
def convert_command(
    source: Path = typer.Argument(..., help="输入 .doc 或 .docx 路径"),
    output: Path = typer.Option(..., "-o", "--output", help="输出 DOCX 路径"),
) -> None:
    """Convert legacy Word input into DOCX using installed desktop software."""

    try:
        result = convert_to_docx(source, output)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]Word 转换完成:[/green] {result['output']}；"
        f"后端 {result['backend']}"
    )


@app.command("derive-rules")
def derive_rules_command(
    template: Path = typer.Argument(..., help="模板 DOCX 路径"),
    output: Path = typer.Option(
        ..., "-o", "--output", help="模块感知 rules.json 输出路径"
    ),
    modules: list[str] = typer.Option(
        [], "--module", help="只提取指定模块，可重复传入"
    ),
    include_page: bool = typer.Option(
        False, "--include-page", help="同时提取页面设置"
    ),
) -> None:
    """Infer page and module-specific rules directly from a template."""

    try:
        rules = build_template_rules(
            template,
            modules=modules or None,
            include_page=include_page,
        )
        write_json_report(rules, output)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]模板规则已生成:[/green] {output}；"
        f"{len(rules['roles'])} 个角色"
    )


@app.command("detect-modules")
def detect_modules_command(
    document: Path = typer.Argument(..., help="待识别 DOCX 路径"),
    output: Path = typer.Option(..., "-o", "--output", help="模块边界 JSON 输出路径"),
) -> None:
    """Detect semantic module boundaries without changing the document."""

    try:
        result = detect_modules(document)
        write_json_report(result, output)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]模块识别完成:[/green] {output}；"
        f"{len(result['modules'])} 个模块区段"
    )


@app.command("compare-modules")
def compare_modules_command(
    template: Path = typer.Argument(..., help="模板 DOCX 路径"),
    draft: Path = typer.Argument(..., help="初稿 DOCX 路径"),
    output: Path = typer.Option(..., "-o", "--output", help="模块差异 JSON 输出路径"),
) -> None:
    """Compare template and draft module presence for Agent decision-making."""

    try:
        result = compare_modules(template, draft)
        write_json_report(result, output)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]模块差异已写入:[/green] {output}；"
        f"初稿缺失 {len(result['missing_in_draft'])} 类模块"
    )


@app.command("align-modules")
def align_modules_command(
    template: Path = typer.Argument(..., help="模板 DOCX 路径"),
    draft: Path = typer.Argument(..., help="初稿 DOCX 路径"),
    output: Path = typer.Option(..., "-o", "--output", help="局部对齐 DOCX 输出路径"),
    modules: list[str] = typer.Option(
        ..., "--module", help="待对齐模块，可重复传入"
    ),
    include_page: bool = typer.Option(
        False, "--include-page", help="同时应用模板页面设置"
    ),
    report_path: Path | None = typer.Option(
        None, "--report", help="可选 JSON 报告输出路径"
    ),
) -> None:
    """Apply only explicitly selected template modules."""

    try:
        report = align_modules_from_template(
            template,
            draft,
            output,
            modules=modules,
            include_page=include_page,
        )
        if report_path is not None:
            write_json_report(report, report_path)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]模块局部对齐完成:[/green] {output}；"
        f"模块 {', '.join(report['selected_modules'])}"
    )


@app.command("format-template")
def format_template_command(
    template: Path = typer.Argument(..., help="模板 .doc 或 .docx 路径"),
    draft: Path = typer.Argument(..., help="初稿 .doc 或 .docx 路径"),
    output: Path = typer.Option(..., "-o", "--output", help="格式化 DOCX 输出路径"),
    report_path: Path | None = typer.Option(
        None, "--report", help="可选 JSON 报告输出路径"
    ),
    markdown: Path | None = typer.Option(
        None, "--md", help="可选 Markdown 报告输出路径"
    ),
) -> None:
    """Run conversion, module recognition, template inference, and repair."""

    try:
        report = format_from_template(template, draft, output)
        if report_path is not None:
            write_json_report(report, report_path)
        if markdown is not None:
            write_markdown_report(report, markdown)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]模板套版完成:[/green] {output}；"
        f"{report['summary']['fixed_issues']} 个问题已修复；"
        f"{len(report['template_rule_roles'])} 个模板角色"
    )


@app.command("validate-rules")
def validate_rules_command(
    rules: Path = typer.Argument(..., help="rules.json 路径"),
) -> None:
    """Strictly validate one rules file."""

    try:
        validate_rules(rules)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(f"[green]规则校验通过:[/green] {rules}")


@app.command("inspect")
def inspect_command(
    document: Path = typer.Argument(..., help="待检查 DOCX 路径"),
    rules: Path = typer.Option(..., "--rules", help="rules.json 路径"),
    output: Path = typer.Option(
        ..., "-o", "--output", help="JSON 报告输出路径"
    ),
    markdown: Path | None = typer.Option(
        None, "--md", help="可选 Markdown 报告输出路径"
    ),
) -> None:
    """Inspect a document without modifying it."""

    try:
        report = inspect_document(document, rules)
        write_json_report(report, output)
        if markdown is not None:
            write_markdown_report(report, markdown)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]检查完成:[/green] {report['summary']['total_issues']} 个问题"
    )


@app.command("fix")
def fix_command(
    document: Path = typer.Argument(..., help="待修复 DOCX 路径"),
    rules: Path = typer.Option(..., "--rules", help="rules.json 路径"),
    output: Path = typer.Option(
        ..., "-o", "--output", help="修复后 DOCX 输出路径"
    ),
    report_path: Path = typer.Option(
        ..., "--report", help="JSON 报告输出路径"
    ),
    markdown: Path = typer.Option(
        ..., "--md", help="Markdown 报告输出路径"
    ),
) -> None:
    """Repair supported formatting into a new DOCX."""

    try:
        report = fix_document(document, rules, output)
        write_json_report(report, report_path)
        write_markdown_report(report, markdown)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]支持范围内修复完成:[/green] "
        f"{report['summary']['fixed_issues']} 个问题；"
        f"输出 {output}"
    )


@app.command("link-references")
def link_references_command(
    document: Path = typer.Argument(..., help="待处理 DOCX 路径"),
    rules: Path = typer.Option(..., "--rules", help="rules.json 路径"),
    output: Path = typer.Option(
        ..., "-o", "--output", help="添加文献跳转后的 DOCX 输出路径"
    ),
    report_path: Path = typer.Option(
        ..., "--report", help="JSON 报告输出路径"
    ),
    markdown: Path = typer.Option(
        ..., "--md", help="Markdown 报告输出路径"
    ),
) -> None:
    """Add citation superscripts and reference cross-links only."""

    try:
        report = link_references(document, rules, output)
        write_json_report(report, report_path)
        write_markdown_report(report, markdown)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]文献跳转处理完成:[/green] "
        f"{report['summary']['fixed_issues']} 个问题；输出 {output}"
    )


@app.command("dry-run")
def dry_run_command(
    document: Path = typer.Argument(..., help="待预览 DOCX 路径"),
    rules: Path = typer.Option(..., "--rules", help="rules.json 路径"),
    markdown: Path = typer.Option(
        ..., "--md", help="Markdown 预览报告输出路径"
    ),
    report_path: Path | None = typer.Option(
        None, "--report", help="可选 JSON 预览报告输出路径"
    ),
) -> None:
    """Preview fixable issues without writing a DOCX."""

    try:
        report = inspect_document(document, rules)
        report["notes"].append(
            "dry-run：fixable=true 的问题是预计会修改的内容，文档未被修改。"
        )
        write_markdown_report(report, markdown)
        if report_path is not None:
            write_json_report(report, report_path)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]预览完成:[/green] "
        f"{report['summary']['fixable_issues']} 个预计可修复问题"
    )


@app.command("extract-samples")
def extract_samples_command(
    template: Path = typer.Argument(..., help="模板 DOCX 路径"),
    output: Path = typer.Option(
        ..., "-o", "--output", help="template_samples.json 输出路径"
    ),
) -> None:
    """Extract representative template paragraphs."""

    try:
        samples = extract_samples(template)
        write_json_report(samples, output)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(f"[green]段落样本已写入:[/green] {output}")


@app.command("plan-structure")
def plan_structure_command(
    template: Path = typer.Argument(..., help="模板 DOCX 路径"),
    draft: Path = typer.Argument(..., help="初稿 DOCX 路径"),
    output: Path = typer.Option(
        ..., "-o", "--output", help="结构差异 JSON 输出路径"
    ),
) -> None:
    """Report semantic sections present in the template but missing in the draft."""

    try:
        result = plan_structure(template, draft)
        write_json_report(result, output)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]结构计划已写入:[/green] {output}；"
        f"缺失 {len(result['missing_roles'])} 个角色"
    )


@app.command("complete-structure")
def complete_structure_command(
    document: Path = typer.Argument(..., help="待补建 DOCX 路径"),
    rules: Path = typer.Option(..., "--rules", help="rules.json 路径"),
    output: Path = typer.Option(
        ..., "-o", "--output", help="补建后 DOCX 输出路径"
    ),
    report_path: Path = typer.Option(
        ..., "--report", help="结构补建 JSON 报告路径"
    ),
) -> None:
    """Insert only missing sections explicitly declared in rules.json."""

    try:
        result = complete_structure(document, rules, output)
        write_json_report(result, report_path)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]显式结构补建完成:[/green] "
        f"{len(result['inserted_roles'])} 个章节；输出 {output}"
    )


@app.command("inspect-structure")
def inspect_structure_command(
    document: Path = typer.Argument(..., help="待检查 DOCX 路径"),
    output: Path = typer.Option(
        ..., "-o", "--output", help="结构事实 JSON 输出路径"
    ),
) -> None:
    """Inspect sections, TOC boundaries, fields, and paragraph locators."""

    try:
        report = inspect_structure(document)
        write_json_report(report, output)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]结构检查完成:[/green] {report['section_count']} 个分节；"
        f"报告 {output}"
    )


@app.command("apply-operations")
def apply_operations_command(
    document: Path = typer.Argument(..., help="待处理 DOCX 路径"),
    plan: Path = typer.Option(..., "--plan", help="结构操作计划 JSON"),
    output: Path = typer.Option(
        ..., "-o", "--output", help="处理后 DOCX 输出路径"
    ),
    report_path: Path = typer.Option(
        ..., "--report", help="操作结果 JSON 输出路径"
    ),
) -> None:
    """Apply a strict, idempotent structural operation plan."""

    try:
        report = apply_structure_operations(document, plan, output)
        write_json_report(report, report_path)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]结构操作完成:[/green] "
        f"{len(report['operations'])} 步；changed={report['changed']}；"
        f"输出 {output}"
    )


@app.command("validate-layout")
def validate_layout_command(
    document: Path = typer.Argument(..., help="待验证 DOCX 路径"),
    expectations: Path = typer.Option(
        ..., "--expect", help="布局期望 JSON"
    ),
    output: Path = typer.Option(
        ..., "-o", "--output", help="布局验证 JSON 输出路径"
    ),
) -> None:
    """Validate structural layout invariants without requiring image analysis."""

    try:
        report = validate_layout(document, expectations)
        write_json_report(report, output)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]布局验证完成:[/green] {report['status']}；报告 {output}"
    )


@app.command("refresh-fields")
def refresh_fields_command(
    document: Path = typer.Argument(..., help="待刷新 DOCX 路径"),
    output: Path = typer.Option(
        ..., "-o", "--output", help="刷新后 DOCX 输出路径"
    ),
    report_path: Path = typer.Option(
        ..., "--report", help="字段刷新 JSON 输出路径"
    ),
) -> None:
    """Refresh fields and real pagination through Microsoft Word."""

    try:
        report = refresh_fields(document, output)
        write_json_report(report, report_path)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]字段刷新完成:[/green] {report['page_count']} 页；"
        f"输出 {output}"
    )


@app.command("visual-compare")
def visual_compare_command(
    source: Path = typer.Argument(..., help="修复前 DOCX 路径"),
    candidate: Path = typer.Argument(..., help="修复后 DOCX 路径"),
    artifacts: Path = typer.Option(
        ..., "--artifacts", help="Word 导出 PDF 的目录"
    ),
    output: Path = typer.Option(
        ..., "-o", "--output", help="视觉比较 JSON 输出路径"
    ),
    dpi: int = typer.Option(120, "--dpi", min=72, max=300),
    threshold: float = typer.Option(
        0.005, "--threshold", min=0.0, max=1.0
    ),
) -> None:
    """Render two DOCX files through Word and compare their PDF pages."""

    try:
        result = compare_visual(
            source,
            candidate,
            artifacts,
            dpi=dpi,
            changed_pixel_threshold=threshold,
        )
        write_json_report(result, output)
    except WordFormatToolError as exc:
        _abort(exc)
    console.print(
        f"[green]视觉比较完成:[/green] {result['status']}；报告 {output}"
    )


if __name__ == "__main__":
    app()
