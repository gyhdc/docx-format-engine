"""Detect OOXML structures that make page-geometry changes visually risky."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any

from docx.oxml.table import CT_Tbl

from .models import AnalyzedParagraph


@dataclass(frozen=True)
class LayoutRiskProfile:
    floating_anchor_count: int
    text_box_count: int
    vml_shape_count: int
    front_matter_table_count: int
    section_count: int

    @property
    def high_risk(self) -> bool:
        return any(
            (
                self.floating_anchor_count,
                self.text_box_count,
                self.vml_shape_count,
                self.front_matter_table_count,
            )
        )

    def to_dict(self) -> dict[str, Any]:
        return {**asdict(self), "high_risk": self.high_risk}

    def uncovered_areas(self) -> list[str]:
        areas: list[str] = []
        if self.floating_anchor_count:
            areas.append(f"floating_objects:{self.floating_anchor_count}")
        if self.text_box_count:
            areas.append(f"text_boxes:{self.text_box_count}")
        if self.vml_shape_count:
            areas.append(f"vml_shapes:{self.vml_shape_count}")
        if self.front_matter_table_count:
            areas.append(f"front_matter_tables:{self.front_matter_table_count}")
        return areas


def _front_matter_table_count(
    document: Any, analyzed: list[AnalyzedParagraph],
) -> int:
    body_items = [item for item in analyzed if item.story == "body"]
    boundary = next(
        (
            item
            for item in body_items
            if item.ignore_reason == "table_of_contents"
            or item.role
            in {
                "abstract",
                "heading_1",
                "acknowledgements",
                "appendix",
                "reference_heading",
            }
        ),
        None,
    )
    if boundary is not None:
        boundary_element = document.paragraphs[boundary.paragraph_index]._p
    else:
        first_body = next((item for item in body_items if item.text.strip()), None)
        boundary_element = (
            document.paragraphs[first_body.paragraph_index]._p
            if first_body is not None
            else None
        )

    count = 0
    for child in document.element.body.iterchildren():
        if child is boundary_element:
            break
        if isinstance(child, CT_Tbl):
            count += 1
    return count


def analyze_layout_risk(
    document: Any, analyzed: list[AnalyzedParagraph],
) -> LayoutRiskProfile:
    """Return a conservative layout-risk inventory."""

    roots = [document.element]
    seen_parts: set[Any] = set()
    for section in document.sections:
        for part in (section.header, section.footer):
            if part._element in seen_parts:
                continue
            seen_parts.add(part._element)
            roots.append(part._element)
    floating_anchor_count = sum(
        len(root.xpath(".//*[local-name()='anchor']")) for root in roots
    )
    text_box_count = sum(
        len(root.xpath(".//*[local-name()='txbxContent']")) for root in roots
    )
    vml_shape_count = sum(
        len(
            root.xpath(
                ".//*[local-name()='pict']"
                " | .//*[local-name()='shape' and namespace-uri()="
                "'urn:schemas-microsoft-com:vml']"
            )
        )
        for root in roots
    )
    return LayoutRiskProfile(
        floating_anchor_count=floating_anchor_count,
        text_box_count=text_box_count,
        vml_shape_count=vml_shape_count,
        front_matter_table_count=_front_matter_table_count(document, analyzed),
        section_count=len(document.sections),
    )
