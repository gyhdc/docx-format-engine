"""Stable semantic paragraph selection for repeatable Agent operations."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from .document_analyzer import analyze_document
from .exceptions import StructureOperationError
from .models import AnalyzedParagraph, ParagraphSelector, Rules


@dataclass(frozen=True)
class LocatedParagraph:
    paragraph_index: int
    text: str
    role: str
    location: str
    paragraph: Any
    analyzed: AnalyzedParagraph


def _analysis_rules() -> Rules:
    return Rules.model_validate({"version": "0.1", "roles": {"body": {}}})


def normalize_paragraph_text(text: str) -> str:
    """Normalize visible whitespace without changing punctuation or case."""

    return re.sub(r"\s+", " ", text).strip()


def _after_boundary(
    analyzed: list[AnalyzedParagraph], after_role: str | None,
) -> int:
    if after_role is None:
        return -1
    if after_role == "table_of_contents":
        indexes = [
            item.paragraph_index
            for item in analyzed
            if item.story == "body"
            and (
                item.ignore_reason == "table_of_contents"
                or item.role
                in {
                    "table_of_contents",
                    "toc_heading",
                    "toc_entry_1",
                    "toc_entry_2",
                    "toc_entry_3",
                }
            )
        ]
    else:
        indexes = [
            item.paragraph_index
            for item in analyzed
            if item.story == "body" and item.role == after_role
        ]
    if not indexes:
        raise StructureOperationError(
            f"未找到定位边界角色 {after_role!r}，无法安全解析目标段落。"
        )
    return max(indexes)


def locate_paragraph(
    document: Any,
    selector: ParagraphSelector,
) -> LocatedParagraph:
    """Resolve one body paragraph or raise on missing and ambiguous matches."""

    analyzed = analyze_document(document, _analysis_rules())
    boundary = _after_boundary(analyzed, selector.after_role)
    expected_text = normalize_paragraph_text(selector.text)
    candidates = [
        item
        for item in analyzed
        if item.story == "body"
        and item.paragraph_index > boundary
        and normalize_paragraph_text(item.text) == expected_text
        and (selector.role is None or item.role == selector.role)
        and (
            selector.style_name is None
            or item.style_name.casefold() == selector.style_name.casefold()
        )
    ]
    if not candidates:
        raise StructureOperationError(
            f"未找到段落: text={selector.text!r}, role={selector.role!r}, "
            f"after_role={selector.after_role!r}。"
        )
    if selector.occurrence is None:
        if len(candidates) != 1:
            locations = ", ".join(item.location for item in candidates[:8])
            raise StructureOperationError(
                f"段落选择器匹配到 {len(candidates)} 个候选，请增加 role、"
                f"after_role 或 occurrence；候选: {locations}"
            )
        selected = candidates[0]
    else:
        offset = selector.occurrence - 1
        if offset >= len(candidates):
            raise StructureOperationError(
                f"段落选择器只有 {len(candidates)} 个候选，"
                f"无法选择 occurrence={selector.occurrence}。"
            )
        selected = candidates[offset]
    return LocatedParagraph(
        paragraph_index=selected.paragraph_index,
        text=selected.text,
        role=selected.role,
        location=selected.location,
        paragraph=document.paragraphs[selected.paragraph_index],
        analyzed=selected,
    )
