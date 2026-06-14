"""Pydantic schemas for rules, analysis records, issues, and reports.

Purpose:
    Define the stable contract shared by the loader, analyzer, inspector,
    fixer, report writers, CLI, and external Agent callers.
MVP scope:
    Unknown fields are rejected. Role entries may be partial so an Agent can
    generate rules for only the document roles it intends to enforce.
"""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    field_validator,
    model_validator,
)

Alignment = Literal["left", "center", "right", "both", "distribute"]
ExpectedPosition = Literal["above_object", "below_object"]
RoleName = Literal[
    "title",
    "title_zh",
    "title_en",
    "abstract",
    "abstract_heading_zh",
    "abstract_body_zh",
    "abstract_heading_en",
    "abstract_body_en",
    "keywords",
    "keywords_zh",
    "keywords_en",
    "acknowledgements",
    "appendix",
    "heading_1",
    "heading_2",
    "heading_3",
    "body",
    "cover_label",
    "cover_value",
    "table_text",
    "header",
    "footer",
    "table_of_contents",
    "toc_heading",
    "toc_entry_1",
    "toc_entry_2",
    "toc_entry_3",
    "figure_caption",
    "table_caption",
    "reference_heading",
    "reference_entry",
    "reference_entry_zh",
    "reference_entry_en",
    "reference",
    "unknown",
]
StoryName = Literal["body", "table_cell", "header", "footer"]
Severity = Literal["info", "warning", "error"]
StructureRole = Literal[
    "abstract",
    "keywords",
    "acknowledgements",
    "appendix",
    "reference_heading",
]


class StrictModel(BaseModel):
    """Base schema that rejects misspelled or unsupported fields."""

    model_config = ConfigDict(extra="forbid", strict=True)


class PriorityRules(StrictModel):
    explicit_rules_over_template: bool = True
    preserve_content: bool = True
    preserve_images: bool = True
    preserve_tables: bool = True
    protect_front_matter: bool = True
    allow_unsafe_page_geometry: bool = False


class PageRules(StrictModel):
    paper_size: Literal["A4", "Letter", "Legal"] | None = None
    page_width_cm: float | None = Field(default=None, gt=0)
    page_height_cm: float | None = Field(default=None, gt=0)
    margin_top_cm: float | None = Field(default=None, ge=0)
    margin_bottom_cm: float | None = Field(default=None, ge=0)
    margin_left_cm: float | None = Field(default=None, ge=0)
    margin_right_cm: float | None = Field(default=None, ge=0)
    header_distance_cm: float | None = Field(default=None, ge=0)
    footer_distance_cm: float | None = Field(default=None, ge=0)


class RoleFormatRule(StrictModel):
    font_east_asia: str | None = None
    font_ascii: str | None = None
    font_size_pt: float | None = Field(default=None, gt=0)
    bold: bool | None = None
    italic: bool | None = None
    alignment: Alignment | None = None
    line_spacing: float | None = Field(default=None, gt=0)
    first_line_indent_chars: float | None = Field(default=None, ge=0)
    hanging_indent_chars: float | None = Field(default=None, ge=0)
    left_indent_cm: float | None = None
    space_before_pt: float | None = Field(default=None, ge=0)
    space_after_pt: float | None = Field(default=None, ge=0)
    keep_with_next: bool | None = None
    expected_position: ExpectedPosition | None = None

    @field_validator("font_east_asia", "font_ascii")
    @classmethod
    def font_names_must_not_be_blank(cls, value: str | None) -> str | None:
        if value is not None and not value.strip():
            raise ValueError("font name must not be blank")
        return value

    @model_validator(mode="after")
    def indentation_modes_must_not_conflict(self) -> RoleFormatRule:
        if (
            self.first_line_indent_chars is not None
            and self.hanging_indent_chars is not None
        ):
            raise ValueError(
                "first_line_indent_chars and hanging_indent_chars "
                "cannot be configured together"
            )
        return self


class DetectionRules(StrictModel):
    title_patterns: list[str] = Field(default_factory=list)
    title_zh_patterns: list[str] = Field(default_factory=list)
    title_en_patterns: list[str] = Field(default_factory=list)
    abstract_patterns: list[str] = Field(default_factory=list)
    keywords_patterns: list[str] = Field(default_factory=list)
    acknowledgements_patterns: list[str] = Field(default_factory=list)
    appendix_patterns: list[str] = Field(default_factory=list)
    heading_1_patterns: list[str] = Field(default_factory=list)
    heading_2_patterns: list[str] = Field(default_factory=list)
    heading_3_patterns: list[str] = Field(default_factory=list)
    figure_caption_patterns: list[str] = Field(default_factory=list)
    table_caption_patterns: list[str] = Field(default_factory=list)
    reference_patterns: list[str] = Field(default_factory=list)


class RequiredSection(StrictModel):
    role: StructureRole
    heading_text: str = Field(min_length=1)
    placeholder_text: str | None = None
    insert_before_role: StructureRole | None = None


class StructureRules(StrictModel):
    required_sections: list[RequiredSection] = Field(default_factory=list)


class Rules(StrictModel):
    version: Literal["0.1"]
    priority: PriorityRules = Field(default_factory=PriorityRules)
    page: PageRules | None = None
    roles: dict[RoleName, RoleFormatRule]
    detection: DetectionRules = Field(default_factory=DetectionRules)
    structure: StructureRules = Field(default_factory=StructureRules)

    @field_validator("roles")
    @classmethod
    def rules_must_define_a_role(
        cls, value: dict[RoleName, RoleFormatRule],
    ) -> dict[RoleName, RoleFormatRule]:
        if not value:
            raise ValueError("at least one role rule is required")
        if "unknown" in value:
            raise ValueError("unknown role cannot have a forced format rule")
        return value


class ParagraphSelector(StrictModel):
    """Stable paragraph selector resolved again before every operation."""

    text: str = Field(min_length=1)
    role: RoleName | None = None
    occurrence: int | None = Field(default=None, ge=1)
    after_role: RoleName | None = None
    style_name: str | None = None

    @field_validator("text", "style_name")
    @classmethod
    def selector_text_must_not_be_blank(cls, value: str | None) -> str | None:
        if value is not None and not value.strip():
            raise ValueError("selector text must not be blank")
        return value


class InsertSectionBeforeOperation(StrictModel):
    type: Literal["insert_section_before"]
    target: ParagraphSelector
    break_type: Literal["next_page"] = "next_page"


class SetHeaderOperation(StrictModel):
    type: Literal["set_header"]
    section_start: ParagraphSelector
    text: str
    alignment: Alignment = "center"
    bottom_border: bool = False


class SetPageNumberOperation(StrictModel):
    type: Literal["set_page_number"]
    section_start: ParagraphSelector
    start: int = Field(default=1, ge=0)
    alignment: Alignment = "center"


class SetFooterTextOperation(StrictModel):
    type: Literal["set_footer_text"]
    section_start: ParagraphSelector
    text: str
    alignment: Alignment = "center"


class SetParagraphPaginationOperation(StrictModel):
    type: Literal["set_paragraph_pagination"]
    target: ParagraphSelector
    page_break_before: bool | None = None
    keep_with_next: bool | None = None
    keep_together: bool | None = None
    widow_control: bool | None = None

    @model_validator(mode="after")
    def at_least_one_pagination_field(self) -> SetParagraphPaginationOperation:
        if all(
            value is None
            for value in (
                self.page_break_before,
                self.keep_with_next,
                self.keep_together,
                self.widow_control,
            )
        ):
            raise ValueError("at least one pagination field is required")
        return self


class RequestFieldUpdateOperation(StrictModel):
    type: Literal["request_field_update"]


StructureOperation = Annotated[
    InsertSectionBeforeOperation
    | SetHeaderOperation
    | SetPageNumberOperation
    | SetFooterTextOperation
    | SetParagraphPaginationOperation
    | RequestFieldUpdateOperation,
    Field(discriminator="type"),
]


class StructureOperationPlan(StrictModel):
    version: Literal["0.1"]
    operations: list[StructureOperation] = Field(min_length=1)


class LayoutExpectations(StrictModel):
    version: Literal["0.1"]
    section_count_at_least: int | None = Field(default=None, ge=1)
    body_start: ParagraphSelector | None = None
    body_first_heading_equals: str | None = None
    toc_before_body: bool | None = None
    section_header_equals: str | None = None
    section_has_page_field: bool | None = None
    section_page_number_starts_at: int | None = Field(default=None, ge=0)
    update_fields_enabled: bool | None = None

    @model_validator(mode="after")
    def section_checks_require_body_start(self) -> LayoutExpectations:
        section_checks = (
            self.section_header_equals is not None
            or self.section_has_page_field is not None
            or self.section_page_number_starts_at is not None
            or self.toc_before_body is not None
        )
        if section_checks and self.body_start is None:
            raise ValueError("section and TOC checks require body_start")
        return self


class TemplatePageSection(StrictModel):
    section_index: int = Field(ge=0)
    page_width_cm: float = Field(gt=0)
    page_height_cm: float = Field(gt=0)
    margin_top_cm: float = Field(ge=0)
    margin_bottom_cm: float = Field(ge=0)
    margin_left_cm: float = Field(ge=0)
    margin_right_cm: float = Field(ge=0)
    header_distance_cm: float = Field(ge=0)
    footer_distance_cm: float = Field(ge=0)


class TemplatePageProfile(StrictModel):
    sections: list[TemplatePageSection]


class TemplateStyleProfile(StrictModel):
    style_id: str
    name: str
    type: Literal["paragraph", "character"]
    base_style: str | None = None
    format: dict[str, Any]


class DocumentOutlineItem(StrictModel):
    role: RoleName
    text: str
    story: StoryName
    location: str


class TemplateParagraphSample(StrictModel):
    paragraph_index: int = Field(ge=0)
    text: str
    text_preview: str
    style_name: str
    guessed_role: RoleName
    format: dict[str, Any]
    story: StoryName
    location: str
    ignore_reason: str | None = None


class LayoutRiskSummary(StrictModel):
    floating_anchor_count: int = Field(ge=0)
    text_box_count: int = Field(ge=0)
    vml_shape_count: int = Field(ge=0)
    front_matter_table_count: int = Field(ge=0)
    section_count: int = Field(ge=0)
    high_risk: bool


class TemplateProfile(StrictModel):
    source_file: str
    page: TemplatePageProfile
    styles: list[TemplateStyleProfile]
    paragraph_samples: list[TemplateParagraphSample]
    structure_outline: list[DocumentOutlineItem]
    layout_risk: LayoutRiskSummary
    notes: list[str]


class StructurePlan(StrictModel):
    template: str
    draft: str
    template_outline: list[DocumentOutlineItem]
    draft_outline: list[DocumentOutlineItem]
    missing_roles: list[StructureRole]
    not_automatically_copied: list[str]


class StructureCompletionResult(StrictModel):
    source: str
    output: str
    inserted_roles: list[StructureRole]
    unchanged_existing_content: bool
    notes: list[str]


class VisualPageComparison(StrictModel):
    page: int = Field(ge=1)
    source_size: list[int] = Field(min_length=2, max_length=2)
    candidate_size: list[int] = Field(min_length=2, max_length=2)
    size_changed: bool
    changed_pixel_ratio: float = Field(ge=0, le=1)


class VisualComparisonReport(StrictModel):
    status: Literal["passed", "warning"]
    source_page_count: int = Field(ge=0)
    candidate_page_count: int = Field(ge=0)
    page_count_changed: bool
    page_size_changed: bool
    changed_pixel_threshold: float = Field(ge=0, le=1)
    pages: list[VisualPageComparison]
    source: str
    candidate: str
    dpi: int = Field(ge=1)
    source_docx: str
    candidate_docx: str
    rendered_pdfs: list[str] = Field(min_length=2, max_length=2)


class AnalyzedParagraph(StrictModel):
    paragraph_index: int = Field(ge=0)
    text: str
    role: RoleName
    style_name: str
    format: dict[str, Any]
    ignore_reason: str | None = None
    story: StoryName = "body"
    location: str = ""


class FormatIssue(StrictModel):
    id: str
    severity: Severity
    paragraph_index: int | None = None
    role: str
    text_preview: str
    field: str
    expected: Any
    actual: Any
    message: str
    fixable: bool
    fixed: bool = False
    approximation: str | None = None
    location: str | None = None


class ReportSummary(StrictModel):
    total_paragraphs: int = Field(ge=0)
    total_issues: int = Field(ge=0)
    fixable_issues: int = Field(ge=0)
    fixed_issues: int = Field(default=0, ge=0)
    unfixed_issues: int = Field(ge=0)


class CoverageSummary(StrictModel):
    story_counts: dict[str, int] = Field(default_factory=dict)
    role_counts: dict[str, int] = Field(default_factory=dict)
    protected_front_matter: int = Field(default=0, ge=0)
    layout_risk: dict[str, Any] = Field(default_factory=dict)
    uncovered_areas: list[str] = Field(default_factory=list)
    toc_unrepairable: int = Field(default=0, ge=0)
    visual_comparison: Literal["not_run", "passed", "warning", "unavailable"] = "not_run"


class FormatReport(StrictModel):
    document: str
    rules: str
    phase: Literal["inspection", "fix"]
    summary: ReportSummary
    issues: list[FormatIssue]
    notes: list[str] = Field(default_factory=list)
    coverage: CoverageSummary = Field(default_factory=CoverageSummary)
