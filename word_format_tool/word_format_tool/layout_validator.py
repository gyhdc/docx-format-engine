"""Machine-readable validation for Word document structure invariants."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from docx import Document
from pydantic import ValidationError

from .exceptions import StructureOperationError
from .models import LayoutExpectations
from .paragraph_locator import locate_paragraph, normalize_paragraph_text
from .structure_inspector import inspect_structure


def _load_expectations(path: str | Path) -> LayoutExpectations:
    expectations_path = Path(path)
    if not expectations_path.is_file():
        raise StructureOperationError(
            f"布局期望不存在或不是文件: {expectations_path}"
        )
    try:
        payload = json.loads(expectations_path.read_text(encoding="utf-8-sig"))
        return LayoutExpectations.model_validate(payload)
    except (OSError, json.JSONDecodeError, ValidationError) as exc:
        raise StructureOperationError(
            f"布局期望无效 {expectations_path}: {exc}"
        ) from exc


def _check(
    checks: list[dict[str, Any]],
    *,
    code: str,
    passed: bool,
    expected: Any,
    actual: Any,
    location: str | None = None,
    suggestion: str | None = None,
) -> None:
    checks.append(
        {
            "code": code,
            "status": "passed" if passed else "failed",
            "expected": expected,
            "actual": actual,
            "location": location,
            "suggestion": None if passed else suggestion,
        }
    )


def validate_layout(
    document_path: str | Path,
    expectations_path: str | Path,
) -> dict[str, Any]:
    """Validate structural expectations without claiming Word pagination facts."""

    expectations = _load_expectations(expectations_path)
    report = inspect_structure(document_path)
    document = Document(document_path)
    checks: list[dict[str, Any]] = []

    if expectations.section_count_at_least is not None:
        _check(
            checks,
            code="section_count_at_least",
            passed=report["section_count"]
            >= expectations.section_count_at_least,
            expected=expectations.section_count_at_least,
            actual=report["section_count"],
            suggestion="在正文首段前插入下一页分节符。",
        )

    body_location = None
    body_paragraph: dict[str, Any] | None = None
    body_section: dict[str, Any] | None = None
    if expectations.body_start is not None:
        located = locate_paragraph(document, expectations.body_start)
        body_location = located.location
        body_paragraph = next(
            item
            for item in report["paragraphs"]
            if item["paragraph_index"] == located.paragraph_index
        )
        body_section = report["sections"][body_paragraph["section_index"]]
        body_starts_section = (
            body_paragraph["section_index"] > 0
            and body_paragraph["starts_section"]
        )
        _check(
            checks,
            code="body_starts_new_section",
            passed=body_starts_section,
            expected=True,
            actual=body_starts_section,
            location=body_location,
            suggestion="对该正文首标题执行 insert_section_before。",
        )

    if expectations.body_first_heading_equals is not None:
        first_heading = report["body_first_heading"]
        actual_text = first_heading["text"] if first_heading else None
        _check(
            checks,
            code="body_first_heading_equals",
            passed=(
                actual_text is not None
                and normalize_paragraph_text(actual_text)
                == normalize_paragraph_text(
                    expectations.body_first_heading_equals
                )
            ),
            expected=expectations.body_first_heading_equals,
            actual=actual_text,
            location=(
                f"body.paragraph[{first_heading['paragraph_index']}]"
                if first_heading
                else None
            ),
            suggestion="检查目录边界和一级标题识别结果。",
        )

    if expectations.toc_before_body is not None and body_paragraph is not None:
        toc_end = report["toc"]["end_paragraph"]
        actual = (
            toc_end is not None
            and toc_end < body_paragraph["paragraph_index"]
        )
        _check(
            checks,
            code="toc_before_body",
            passed=actual == expectations.toc_before_body,
            expected=expectations.toc_before_body,
            actual=actual,
            location=body_location,
            suggestion="检查目录项范围，确保正文标题不在目录区域内。",
        )

    if expectations.section_header_equals is not None and body_section is not None:
        actual_header = body_section["header_text"]
        _check(
            checks,
            code="section_header_equals",
            passed=actual_header == expectations.section_header_equals,
            expected=expectations.section_header_equals,
            actual=actual_header,
            location=f"section[{body_section['index']}].header",
            suggestion="对正文节执行 set_header，并解除前节链接。",
        )

    if expectations.section_has_page_field is not None and body_section is not None:
        has_page = any(
            "PAGE" in code.upper()
            for code in body_section["footer_field_codes"]
        )
        _check(
            checks,
            code="section_has_page_field",
            passed=has_page == expectations.section_has_page_field,
            expected=expectations.section_has_page_field,
            actual=has_page,
            location=f"section[{body_section['index']}].footer",
            suggestion="对正文节执行 set_page_number。",
        )

    if (
        expectations.section_page_number_starts_at is not None
        and body_section is not None
    ):
        actual_start = body_section["page_number_start"]
        _check(
            checks,
            code="section_page_number_starts_at",
            passed=actual_start
            == expectations.section_page_number_starts_at,
            expected=expectations.section_page_number_starts_at,
            actual=actual_start,
            location=f"section[{body_section['index']}].sectPr",
            suggestion="设置正文节 pgNumType 的 start 值。",
        )

    if expectations.update_fields_enabled is not None:
        actual_update = report["update_fields_enabled"]
        _check(
            checks,
            code="update_fields_enabled",
            passed=actual_update == expectations.update_fields_enabled,
            expected=expectations.update_fields_enabled,
            actual=actual_update,
            suggestion="执行 request_field_update 或 refresh-fields。",
        )

    status = "passed" if all(item["status"] == "passed" for item in checks) else "failed"
    return {
        "document": report["document"],
        "expectations": str(expectations_path),
        "status": status,
        "checks": checks,
        "word_pagination_available": False,
    }
