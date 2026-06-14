# Reference Formatting and Cross-Link Design

## Goal

Extend `word_format_tool` so a thesis can:

- distinguish the reference heading from reference entries;
- format reference entries according to explicit rules;
- recognize numeric citations in body text;
- render recognized citations as superscript;
- link each visible citation number to its matching reference entry;
- link each reference label back to its first body citation;
- report missing, duplicate, discontinuous, unlinked, and unused references;
- preserve all visible document content and unrelated formatting.

The tool must not invent bibliographic metadata. It only links and formats
reference text already present in the document.

## Template Evidence

The supplied 2025 template establishes these behaviors:

- body citations are bracketed numeric superscripts, including `[1]` and
  adjacent forms such as `[3][4]`;
- the reference heading is a centered fourth-level-size heading;
- entries start with `[n]`;
- entries use a hanging indent and 1.25 line spacing.

Exact typography remains controlled by `rules.json`.

## Considered Approaches

### Rewrite complete paragraphs

Reconstruct every affected paragraph from plain text, then apply links and
formatting. This is simple but destroys mixed formatting, fields, drawings,
proofing metadata, and other run-level OOXML. It is rejected.

### Microsoft Word automation

Use Word COM to insert bookmarks and hyperlinks. Word has complete field
support, but this makes the core tool Windows-only, requires an installed Word
application, and is difficult to test deterministically. It is reserved for
optional field refresh, not core linking.

### In-place OOXML enhancement

Parse existing paragraphs, split only text runs that contain recognized
citations, clone their run properties, and insert native internal hyperlinks
and bookmarks. This preserves visible text and unrelated formatting while
remaining deterministic and testable. This is the selected approach.

## Document Model

Add these roles:

- `reference_heading`: the heading that starts a reference section;
- `reference_entry`: a numbered or section-contained bibliography paragraph.

Keep `reference` as a compatibility alias for existing rule files.

A reference section begins at a heading matching `参考文献`, `References`, or a
configured heading pattern. It ends at the next known terminal heading such as
`致谢`, `附录`, `Acknowledgements`, or `Appendix`. Numbered entry recognition
still works without a heading so partial documents remain supported.

## Citation Grammar

Recognize bracketed Arabic numeric citations in body paragraphs:

- single: `[1]`;
- adjacent: `[1][2]`;
- comma list: `[1,2]`, `[1，2]`;
- range: `[1-3]`, `[1–3]`;
- mixed: `[1,3-5]`;
- full-width brackets: `［1］`.

Only numbers that correspond to parsed reference entries are linkable. A group
containing an unknown number is reported. Visible text is never normalized or
renumbered automatically.

For list and range markers, every visible number receives its own hyperlink.
For a range, the two displayed endpoints are directly clickable; all expanded
numbers still participate in missing-reference validation.

## Navigation

Reference entry `n` receives a deterministic bookmark:

```text
_SDAU_REF_000n
```

Each citation occurrence receives:

```text
_SDAU_CITE_000n_001
```

The visible number in a body citation links to `_SDAU_REF_000n`. The reference
label links back to the first `_SDAU_CITE_000n_001` occurrence. Existing
third-party bookmarks and hyperlinks are preserved.

Running the fixer repeatedly must not duplicate bookmarks, hyperlinks, or
visible text.

## Formatting

`reference_heading` and `reference_entry` use separate role rules. When only
the legacy `reference` role exists:

- numbered entries use `reference`;
- the heading falls back to `heading_1`;
- no existing rule file becomes invalid.

Citation formatting is limited to superscript and link metadata. Font family,
font size, bold, italic, color, and surrounding punctuation are preserved.

## Inspection

The inspector reports:

- duplicate reference numbers;
- missing numbers inside the bibliography sequence;
- body citations without matching entries;
- reference entries never cited;
- citations that are not superscript;
- citations lacking a correct internal hyperlink;
- reference entries lacking the expected bookmark;
- reference labels lacking a return link when a citation exists.

Content problems are not auto-fixed when fixing would require inventing,
deleting, or renumbering bibliographic text.

## Save Safety

Content preservation compares logical visible text rather than individual
`w:t` node boundaries because safe hyperlink insertion may split runs.

The save validator additionally guarantees:

- every pre-existing bookmark remains present;
- every pre-existing internal hyperlink anchor remains present;
- only `_SDAU_REF_` and `_SDAU_CITE_` navigation objects may be added by this
  feature;
- tables, formulas, images, paragraph order, and visible text remain unchanged.

Output remains atomic and the source file cannot be overwritten.

## Integration

The normal `fix` workflow performs reference linking after ordinary format
repair and before atomic save. A dedicated `link-references` CLI command also
allows navigation enhancement without applying unrelated formatting rules.

Both commands produce a structured report.

## Acceptance Criteria

- Existing 29 tests remain green.
- New tests prove section-aware role classification.
- New tests prove single, adjacent, list, range, and full-width citations.
- New tests prove superscript and correct internal hyperlink anchors.
- New tests prove deterministic bookmarks and backlinks.
- A second run makes no navigation changes.
- Missing and duplicate references are reported without unsafe rewriting.
- Existing bookmarks and hyperlinks survive save validation.
- The supplied template pattern is reproduced in a synthetic end-to-end
  document and renders without visible text or layout corruption.
