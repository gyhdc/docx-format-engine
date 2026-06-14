"""Public exception hierarchy for predictable API and CLI failures.

Purpose:
    Convert filesystem, JSON, schema, DOCX, and output-safety failures into
    concise domain errors without swallowing their original cause.
MVP scope:
    Exceptions carry human-readable messages; callers can inspect ``__cause__``
    when lower-level diagnostics are needed.
"""

from __future__ import annotations


class WordFormatToolError(Exception):
    """Base class for all expected toolkit failures."""


class InputFileError(WordFormatToolError):
    """An input path is missing, unreadable, or has an unsupported type."""


class DocumentReadError(WordFormatToolError):
    """A DOCX package could not be opened or parsed."""


class RuleValidationError(WordFormatToolError):
    """A rules file is invalid JSON or does not match the rules schema."""


class UnsafeOutputPathError(WordFormatToolError):
    """An output path would overwrite a protected input file."""


class ReportWriteError(WordFormatToolError):
    """A report could not be serialized or written."""


class VisualValidationUnavailableError(WordFormatToolError):
    """Required Word/PDF visual validation backend is unavailable."""


class StructureOperationError(WordFormatToolError):
    """A structural selector or document operation could not be completed safely."""


class WordAutomationUnavailableError(WordFormatToolError):
    """Microsoft Word automation is unavailable or failed."""


class FieldRefreshIntegrityError(WordFormatToolError):
    """A field refresh changed protected document structure."""
