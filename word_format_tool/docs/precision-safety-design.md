# Precision and Safety Hardening Design

## Goal

Improve template-driven draft repair without broadening the tool beyond
deterministic DOCX formatting. The fixer should change only fields that are
known to violate rules, preserve document content, and produce reports that
distinguish intentionally ignored structures from genuine recognition gaps.

## Scope

1. Apply only the fields represented by fixable inspection issues.
2. Save through a same-directory temporary DOCX, reopen and validate it, then
   atomically replace the requested output.
3. Validate preservation of paragraph text, table text, embedded media, math
   objects, and body block order before publishing the output.
4. Mark table-of-contents paragraphs as intentionally ignored for formatting,
   while reporting missing `PAGEREF` bookmark targets.
5. Reject role rules that configure both first-line and hanging indentation.
6. Allow blank paragraphs between a caption and its adjacent image or table.

## Non-Goals

- Formatting table-cell text, headers, footers, text boxes, or floating shapes.
- Updating Word fields or generating a table of contents.
- Adding missing captions, references, or thesis content.
- Moving images, tables, formulas, captions, or paragraphs.
- Inferring layout from rendered pages.

## Repair Model

Inspection remains the source of truth. Page fields are changed only when the
inspection report contains a fixable document-level issue for that field.
Paragraph fields are grouped by paragraph index and applied only when the same
field has a fixable issue. A role rule alone is not sufficient to rewrite a
field that already passed inspection.

Run-level font, size, bold, and italic changes remain paragraph-wide when that
specific field is erroneous. Other run properties are left untouched. This
preserves deliberate emphasis when the emphasis field itself is not governed
or is already correct.

## Safe Output

The output is built in a temporary file located beside the requested output so
the final replacement stays on the same filesystem. The temporary DOCX is
reopened, inspected, and compared with a source content fingerprint. Only after
all checks pass is it atomically moved into place. Any failure removes the
temporary file and leaves an existing destination unchanged.

The content fingerprint covers:

- body paragraph text in order;
- table cell text and table dimensions;
- embedded media names and hashes;
- math object count;
- top-level body block kind and order.

Formatting XML is intentionally excluded because formatting is the expected
change.

## Report Semantics

TOC paragraphs keep the `unknown` role for backward compatibility but include
an internal ignore reason. Inspection skips role warnings for intentionally
ignored paragraphs and records an aggregate note. `PAGEREF` targets are
compared with document bookmarks so field errors that Word would expose during
an update remain visible as non-fixable issues. Other unknown paragraphs
continue to produce warnings.

## Testing

Each behavior is introduced through a failing pytest test:

- fixing line spacing does not erase correct run emphasis;
- a failed preservation check does not replace an existing output;
- TOC paragraphs do not inflate unresolved issue counts;
- conflicting indentation rules fail schema validation;
- captions tolerate intervening blank paragraphs.

The complete unit suite, package build, CLI smoke tests, and the supplied real
template/draft workflow must pass before the optimized version is committed.
