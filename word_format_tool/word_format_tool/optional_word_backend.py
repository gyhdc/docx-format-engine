"""Optional Microsoft Word backend for field refresh and real pagination."""

from __future__ import annotations

import os
import shutil
import tempfile
import json
from pathlib import Path
from typing import Any, Protocol

from .document_io import load_docx, prepare_output_path, require_docx_file
from .exceptions import (
    FieldRefreshIntegrityError,
    WordAutomationUnavailableError,
    WordFormatToolError,
)
from .structure_inspector import inspect_structure


WD_STATISTIC_PAGES = 2


class FieldRefreshBackend(Protocol):
    name: str

    def refresh(self, document_path: Path) -> dict[str, Any]:
        """Refresh one writable DOCX in place and return pagination facts."""


class WordComFieldRefreshBackend:
    """Refresh fields through an installed desktop Microsoft Word."""

    name = "microsoft-word"

    def refresh(self, document_path: Path) -> dict[str, Any]:
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise WordAutomationUnavailableError(
                "字段刷新需要 Windows、Microsoft Word 和可选依赖 pywin32。"
            ) from exc

        word = None
        document = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            document = word.Documents.Open(
                str(document_path.resolve()),
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            field_count = int(document.Fields.Count)
            document.Fields.Update()
            toc_count = int(document.TablesOfContents.Count)
            for index in range(1, toc_count + 1):
                document.TablesOfContents(index).Update()
            document.Repaginate()
            page_count = int(document.ComputeStatistics(WD_STATISTIC_PAGES))
            document.Save()
            return {
                "page_count": page_count,
                "field_count": field_count,
                "toc_count": toc_count,
            }
        except WordAutomationUnavailableError:
            raise
        except Exception as exc:
            raise WordAutomationUnavailableError(
                f"Microsoft Word 字段刷新失败: {exc}"
            ) from exc
        finally:
            if document is not None:
                document.Close(SaveChanges=True)
            if word is not None:
                word.Quit()


def _normalized_field_codes(codes: list[str]) -> list[str]:
    return [" ".join(code.split()).upper() for code in codes]


def _protected_structure_fingerprint(document_path: Path) -> dict[str, Any]:
    report = inspect_structure(document_path)
    body_heading = report.get("body_first_heading")
    heading_section = None
    if body_heading is not None:
        paragraph_index = body_heading["paragraph_index"]
        paragraph = next(
            (
                item
                for item in report["paragraphs"]
                if item["paragraph_index"] == paragraph_index
            ),
            None,
        )
        heading_section = (
            paragraph["section_index"] if paragraph is not None else None
        )
    return {
        "section_count": report["section_count"],
        "toc_present": report["toc"]["present"],
        "body_first_heading": (
            None
            if body_heading is None
            else {
                "text": body_heading["text"],
                "role": body_heading["role"],
                "section_index": heading_section,
            }
        ),
        "sections": [
            {
                "break_type": section["break_type"],
                "header_text": section["header_text"].strip(),
                "header_field_codes": _normalized_field_codes(
                    section["header_field_codes"]
                ),
                "footer_field_codes": _normalized_field_codes(
                    section["footer_field_codes"]
                ),
                "page_number_start": section["page_number_start"],
            }
            for section in report["sections"]
        ],
    }


def _structure_differences(
    before: dict[str, Any], after: dict[str, Any],
) -> dict[str, dict[str, Any]]:
    return {
        key: {"before": before.get(key), "after": after.get(key)}
        for key in before.keys() | after.keys()
        if before.get(key) != after.get(key)
    }


def refresh_fields_to_path(
    document_path: str | Path,
    output_path: str | Path,
    *,
    backend: FieldRefreshBackend | None = None,
) -> dict[str, Any]:
    """Refresh a temporary copy and atomically publish only a valid DOCX."""

    source = require_docx_file(document_path)
    output = prepare_output_path(
        output_path, protected_input=source, suffix=".docx"
    )
    selected_backend = backend or WordComFieldRefreshBackend()
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{output.stem}-refresh-",
        suffix=".tmp.docx",
        dir=output.parent,
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    before_fingerprint = _protected_structure_fingerprint(source)
    try:
        shutil.copyfile(source, temporary)
        facts = selected_backend.refresh(temporary)
        load_docx(temporary)
        after_fingerprint = _protected_structure_fingerprint(temporary)
        differences = _structure_differences(
            before_fingerprint, after_fingerprint
        )
        if differences:
            raise FieldRefreshIntegrityError(
                "字段刷新改变了受保护的结构指纹，输出未发布: "
                + json.dumps(differences, ensure_ascii=False)
            )
        os.replace(temporary, output)
    except WordFormatToolError:
        raise
    except Exception as exc:
        raise WordAutomationUnavailableError(
            f"字段刷新输出发布失败: {exc}"
        ) from exc
    finally:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass
    return {
        "source": str(source),
        "output": str(output),
        "backend": selected_backend.name,
        "structure_preserved": True,
        "section_count": before_fingerprint["section_count"],
        **facts,
    }
