"""High-level template-to-draft formatting workflow."""

from __future__ import annotations

import tempfile
from pathlib import Path

from .conversion import convert_to_docx
from .format_fixer import fix_document_with_rules_to_path
from .template_rules import build_template_rules_model


def _docx_input(path: str | Path, temporary_dir: Path) -> tuple[Path, dict | None]:
    source = Path(path)
    if source.suffix.lower() == ".docx":
        return source, None
    converted = temporary_dir / f"{source.stem}.docx"
    return converted, convert_to_docx(source, converted)


def format_from_template(
    template_path: str | Path,
    draft_path: str | Path,
    output_path: str | Path,
) -> dict:
    """Convert inputs when needed, infer module rules, and format in one call."""

    with tempfile.TemporaryDirectory(prefix="wordfmt-") as temporary:
        temporary_dir = Path(temporary)
        template_docx, template_conversion = _docx_input(
            template_path, temporary_dir
        )
        draft_docx, draft_conversion = _docx_input(draft_path, temporary_dir)
        rules = build_template_rules_model(template_docx)
        report = fix_document_with_rules_to_path(
            draft_docx,
            rules,
            output_path,
            rules_label=f"template:{Path(template_path)}",
        ).model_dump(mode="json")

    report["template"] = str(template_path)
    report["output"] = str(Path(output_path))
    report["module_counts"] = report["coverage"]["role_counts"]
    report["template_rule_roles"] = sorted(rules.roles)
    report["conversions"] = {
        "template": template_conversion,
        "draft": draft_conversion,
    }
    return report
