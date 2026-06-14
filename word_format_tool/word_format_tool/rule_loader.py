"""Load and strictly validate Agent-generated ``rules.json`` files.

Purpose:
    Provide one entry point for JSON decoding, regular-expression validation,
    Pydantic schema validation, and clear field-path error reporting.
MVP scope:
    Supports rules schema version ``0.1`` only.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from pydantic import ValidationError

from .exceptions import InputFileError, RuleValidationError
from .models import RoleFormatRule, RoleName, Rules


def _require_json_file(path: str | Path) -> Path:
    resolved = Path(path)
    if not resolved.is_file():
        raise InputFileError(f"规则文件不存在或不是文件: {resolved}")
    if resolved.suffix.lower() != ".json":
        raise InputFileError(f"规则文件必须是 .json: {resolved}")
    return resolved


def _format_validation_error(error: ValidationError) -> str:
    lines = ["rules.json 不符合 schema:"]
    for item in error.errors(include_url=False):
        location = ".".join(str(part) for part in item["loc"]) or "<root>"
        lines.append(f"- {location}: {item['msg']}")
    return "\n".join(lines)


def _validate_patterns(rules: Rules) -> None:
    errors: list[str] = []
    for field_name, patterns in rules.detection.model_dump().items():
        for index, pattern in enumerate(patterns):
            try:
                re.compile(pattern)
            except re.error as exc:
                errors.append(f"detection.{field_name}.{index}: {exc}")
    if errors:
        raise RuleValidationError(
            "rules.json 包含无效正则表达式:\n- " + "\n- ".join(errors)
        )


def load_rules(path: str | Path) -> Rules:
    """Load one rules file and return its validated schema model."""

    rules_path = _require_json_file(path)
    try:
        raw: Any = json.loads(rules_path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError) as exc:
        raise InputFileError(f"无法读取规则文件 {rules_path}: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise RuleValidationError(
            f"rules.json 不是有效 JSON: line {exc.lineno}, "
            f"column {exc.colno}: {exc.msg}"
        ) from exc

    try:
        rules = Rules.model_validate(raw)
    except ValidationError as exc:
        raise RuleValidationError(_format_validation_error(exc)) from exc

    _validate_patterns(rules)
    return rules


def validate_rules_file(path: str | Path) -> bool:
    """Validate a rules file, raising a domain error on failure."""

    load_rules(path)
    return True


def resolve_role_rule(
    rules: Rules,
    role: RoleName,
) -> RoleFormatRule | None:
    """Resolve new reference roles while preserving legacy rule files."""

    direct = rules.roles.get(role)
    if direct is not None:
        return direct
    if role in {"title_zh", "title_en"}:
        return rules.roles.get("title")
    if role in {"reference_entry_zh", "reference_entry_en"}:
        return rules.roles.get("reference_entry") or rules.roles.get("reference")
    if role == "reference_entry":
        return rules.roles.get("reference")
    if role == "reference_heading":
        return rules.roles.get("heading_1")
    if role in {"abstract_heading_zh", "abstract_heading_en"}:
        return rules.roles.get("abstract") or rules.roles.get("heading_1")
    if role in {"abstract_body_zh", "abstract_body_en"}:
        return rules.roles.get("abstract") or rules.roles.get("body")
    if role in {"keywords_zh", "keywords_en"}:
        return rules.roles.get("keywords") or rules.roles.get("body")
    if role in {"toc_heading", "toc_entry_1", "toc_entry_2", "toc_entry_3"}:
        return rules.roles.get("table_of_contents")
    return None
