"""Conservative repair helpers for Word table-of-contents navigation."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .models import AnalyzedParagraph

PAGEREF_PATTERN = re.compile(
    r"(\bPAGEREF\s+)(?:\"([^\"]+)\"|([^\s\\]+))",
    flags=re.IGNORECASE,
)


@dataclass(frozen=True)
class TocRepairResult:
    repaired_targets: int
    unrepairable_targets: int
    update_fields_enabled: bool

    @property
    def changed(self) -> bool:
        return self.repaired_targets > 0 or self.update_fields_enabled


def _normalized_heading(text: str) -> str:
    first_column = re.split(r"\t|\.{3,}|…{2,}", text, maxsplit=1)[0]
    without_page = re.sub(r"\s+\d+\s*$", "", first_column)
    return re.sub(r"\s+", " ", without_page).strip().casefold()


def find_toc_heading_match(
    analyzed: list[AnalyzedParagraph],
    toc_item: AnalyzedParagraph,
) -> AnalyzedParagraph | None:
    """Return one exact recognized heading match, otherwise ``None``."""

    target = _normalized_heading(toc_item.text)
    if not target:
        return None
    matches = [
        item
        for item in analyzed
        if item.story == "body"
        and item.role in {"heading_1", "heading_2", "heading_3"}
        and _normalized_heading(item.text) == target
    ]
    return matches[0] if len(matches) == 1 else None


def _next_bookmark_id(document: Any) -> int:
    identifiers = document.element.xpath(
        ".//*[local-name()='bookmarkStart']/@*[local-name()='id']"
    )
    numeric = [int(value) for value in identifiers if str(value).isdigit()]
    return max(numeric, default=0) + 1


def _ensure_heading_bookmark(
    document: Any,
    heading: AnalyzedParagraph,
    bookmark_name: str,
    bookmark_id: int,
) -> None:
    paragraph = document.paragraphs[heading.paragraph_index]
    existing = paragraph._p.xpath(
        "./*[local-name()='bookmarkStart' "
        f"and @*[local-name()='name']='{bookmark_name}']"
    )
    if existing:
        return
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), bookmark_name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    insert_at = 1 if paragraph._p.pPr is not None else 0
    paragraph._p.insert(insert_at, start)
    paragraph._p.append(end)


def _enable_update_fields(document: Any) -> bool:
    settings = document.settings.element
    elements = settings.xpath("./*[local-name()='updateFields']")
    if elements:
        element = elements[0]
        current = element.get(qn("w:val"))
        if current in {"true", "1", "on"}:
            return False
    else:
        element = OxmlElement("w:updateFields")
        settings.append(element)
    element.set(qn("w:val"), "true")
    return True


def repair_toc_navigation(
    document: Any, analyzed: list[AnalyzedParagraph],
) -> TocRepairResult:
    """Repair only uniquely matched TOC `PAGEREF` targets."""

    repaired = 0
    unrepairable = 0
    bookmark_id = _next_bookmark_id(document)
    bookmark_names = set(
        document.element.xpath(
            ".//*[local-name()='bookmarkStart']/@*[local-name()='name']"
        )
    )
    for item in analyzed:
        if item.story != "body" or item.ignore_reason != "table_of_contents":
            continue
        paragraph = document.paragraphs[item.paragraph_index]
        instruction_nodes = paragraph._p.xpath(
            ".//*[local-name()='instrText']"
        )
        for instruction in instruction_nodes:
            text = instruction.text or ""
            match = PAGEREF_PATTERN.search(text)
            if match is None:
                continue
            current_target = match.group(2) or match.group(3)
            if current_target in bookmark_names:
                continue
            heading = find_toc_heading_match(analyzed, item)
            if heading is None:
                unrepairable += 1
                continue
            bookmark_name = f"_SDAU_TOC_{heading.paragraph_index:04d}"
            _ensure_heading_bookmark(
                document, heading, bookmark_name, bookmark_id
            )
            bookmark_names.add(bookmark_name)
            bookmark_id += 1
            instruction.text = PAGEREF_PATTERN.sub(
                lambda match, name=bookmark_name: f"{match.group(1)}{name}",
                text,
            )
            repaired += 1
    update_enabled = (
        _enable_update_fields(document)
        if repaired and unrepairable == 0
        else False
    )
    return TocRepairResult(
        repaired_targets=repaired,
        unrepairable_targets=unrepairable,
        update_fields_enabled=update_enabled,
    )
