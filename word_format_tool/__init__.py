"""Project-local facade for the deterministic Word formatting engine."""

from .word_format_tool import (
    compare_visual,
    complete_structure,
    fix_document,
    inspect_document,
    link_references,
    plan_structure,
    profile_template,
    rules_schema,
    validate_rules,
)

__all__ = [
    "compare_visual",
    "complete_structure",
    "fix_document",
    "inspect_document",
    "link_references",
    "plan_structure",
    "profile_template",
    "rules_schema",
    "validate_rules",
]
