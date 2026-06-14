"""Extract page, style, and representative paragraph data from a template.

Purpose:
    Preserve both Word style definitions and actual sample formatting for an
    external Agent that will synthesize ``rules.json``.
MVP scope:
    Extracts body paragraphs and paragraph/character styles. It does not render
    pages or infer formatting from headers, footers, text boxes, or shapes.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from docx.enum.style import WD_STYLE_TYPE

from .document_analyzer import analyze_document
from .document_io import load_docx
from .layout_safety import analyze_layout_risk
from .models import Rules
from .ooxml_utils import get_style_format, section_to_dict

STYLE_TYPE_NAMES = {
    WD_STYLE_TYPE.PARAGRAPH: "paragraph",
    WD_STYLE_TYPE.CHARACTER: "character",
}


def _profile_rules() -> Rules:
    return Rules.model_validate({"version": "0.1", "roles": {"body": {}}})


def extract_samples(template_path: str | Path) -> list[dict[str, Any]]:
    """Extract non-empty paragraphs with role guesses and effective format."""

    document, _ = load_docx(template_path)
    samples: list[dict[str, Any]] = []
    for item in analyze_document(document, _profile_rules()):
        if not item.text.strip():
            continue
        samples.append(
            {
                "paragraph_index": item.paragraph_index,
                "text": item.text,
                "text_preview": item.text[:200],
                "style_name": item.style_name,
                "guessed_role": item.role,
                "format": item.format,
                "story": item.story,
                "location": item.location,
                "ignore_reason": item.ignore_reason,
            }
        )
    return samples


def profile_template(template_path: str | Path) -> dict[str, Any]:
    """Build a JSON-serializable profile for one Word template."""

    document, source_path = load_docx(template_path)
    styles: list[dict[str, Any]] = []
    for style in document.styles:
        if style.type not in STYLE_TYPE_NAMES:
            continue
        styles.append(
            {
                "style_id": style.style_id,
                "name": style.name,
                "type": STYLE_TYPE_NAMES[style.type],
                "base_style": (
                    style.base_style.name if style.base_style is not None else None
                ),
                "format": get_style_format(style),
            }
        )

    analyzed = analyze_document(document, _profile_rules())
    layout_risk = analyze_layout_risk(document, analyzed)
    return {
        "source_file": source_path.name,
        "page": {
            "sections": [
                {
                    "section_index": index,
                    **section_to_dict(section),
                }
                for index, section in enumerate(document.sections)
            ]
        },
        "styles": styles,
        "paragraph_samples": extract_samples(source_path),
        "structure_outline": [
            {
                "role": item.role,
                "text": item.text,
                "story": item.story,
                "location": item.location,
            }
            for item in analyzed
            if item.story == "body" and item.text.strip()
        ],
        "layout_risk": layout_risk.to_dict(),
        "notes": [
            "段落主格式取最长非空白 run。",
            "同时保留 style 定义与样本实际格式，供外部 Agent 判断。",
            "结构清单只提供事实，不自动复制模板正文或高风险对象。",
        ],
    }
