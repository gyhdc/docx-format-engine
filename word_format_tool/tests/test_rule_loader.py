from __future__ import annotations

import json
from pathlib import Path

import pytest

from word_format_tool.exceptions import RuleValidationError
from word_format_tool.rule_loader import load_rules, validate_rules_file


def test_load_rules_accepts_valid_schema(rules_path: Path) -> None:
    rules = load_rules(rules_path)

    assert rules.version == "0.1"
    assert rules.roles["body"].font_size_pt == 12
    assert validate_rules_file(rules_path) is True


def test_load_rules_reports_invalid_field_location(
    tmp_path: Path, rules_data: dict,
) -> None:
    rules_data["roles"]["body"]["alignment"] = "diagonal"
    path = tmp_path / "invalid.json"
    path.write_text(json.dumps(rules_data), encoding="utf-8")

    with pytest.raises(RuleValidationError, match=r"roles\.body\.alignment"):
        load_rules(path)


def test_load_rules_rejects_unknown_fields(
    tmp_path: Path, rules_data: dict,
) -> None:
    rules_data["unexpected"] = True
    path = tmp_path / "invalid.json"
    path.write_text(json.dumps(rules_data), encoding="utf-8")

    with pytest.raises(RuleValidationError, match="unexpected"):
        load_rules(path)


def test_load_rules_does_not_coerce_numeric_strings(
    tmp_path: Path, rules_data: dict,
) -> None:
    rules_data["roles"]["body"]["font_size_pt"] = "12"
    path = tmp_path / "coercive.json"
    path.write_text(json.dumps(rules_data), encoding="utf-8")

    with pytest.raises(RuleValidationError, match=r"roles\.body\.font_size_pt"):
        load_rules(path)


def test_load_rules_rejects_conflicting_first_line_and_hanging_indents(
    tmp_path: Path, rules_data: dict,
) -> None:
    rules_data["roles"]["body"]["hanging_indent_chars"] = 2
    path = tmp_path / "conflicting-indents.json"
    path.write_text(
        json.dumps(rules_data, ensure_ascii=False), encoding="utf-8"
    )

    with pytest.raises(
        RuleValidationError,
        match=r"roles\.body.*first_line_indent_chars.*hanging_indent_chars",
    ):
        load_rules(path)
