"""Optional DOCX/PDF visual regression comparison."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Protocol

from .document_io import require_docx_file
from .exceptions import InputFileError, VisualValidationUnavailableError


@dataclass(frozen=True)
class RasterPage:
    width: int
    height: int
    pixels: bytes
    channels: int = 3


class DocxPdfRenderer(Protocol):
    def render(self, docx_path: Path, pdf_path: Path) -> None: ...


def _changed_pixel_ratio(source: RasterPage, candidate: RasterPage) -> float:
    if (
        source.width != candidate.width
        or source.height != candidate.height
        or source.channels != candidate.channels
    ):
        return 1.0
    pixel_count = source.width * source.height
    if pixel_count == 0:
        return 0.0
    changed = 0
    channels = source.channels
    for offset in range(0, len(source.pixels), channels):
        if (
            source.pixels[offset : offset + channels]
            != candidate.pixels[offset : offset + channels]
        ):
            changed += 1
    return round(changed / pixel_count, 8)


def compare_raster_pages(
    source_pages: Sequence[RasterPage],
    candidate_pages: Sequence[RasterPage],
    *,
    changed_pixel_threshold: float = 0.005,
) -> dict[str, Any]:
    """Compare already-rendered pages without requiring a PDF backend."""

    page_count_changed = len(source_pages) != len(candidate_pages)
    pages: list[dict[str, Any]] = []
    page_size_changed = False
    for index, (source, candidate) in enumerate(
        zip(source_pages, candidate_pages, strict=False)
    ):
        size_changed = (
            source.width != candidate.width
            or source.height != candidate.height
        )
        page_size_changed = page_size_changed or size_changed
        pages.append(
            {
                "page": index + 1,
                "source_size": [source.width, source.height],
                "candidate_size": [candidate.width, candidate.height],
                "size_changed": size_changed,
                "changed_pixel_ratio": _changed_pixel_ratio(source, candidate),
            }
        )
    warning = (
        page_count_changed
        or page_size_changed
        or any(
            page["changed_pixel_ratio"] > changed_pixel_threshold
            for page in pages
        )
    )
    return {
        "status": "warning" if warning else "passed",
        "source_page_count": len(source_pages),
        "candidate_page_count": len(candidate_pages),
        "page_count_changed": page_count_changed,
        "page_size_changed": page_size_changed,
        "changed_pixel_threshold": changed_pixel_threshold,
        "pages": pages,
    }


def load_pdf_raster_pages(
    pdf_path: str | Path,
    *,
    dpi: int = 120,
) -> list[RasterPage]:
    """Render PDF pages through optional PyMuPDF."""

    path = Path(pdf_path)
    if not path.is_file() or path.suffix.lower() != ".pdf":
        raise InputFileError(f"PDF 文件不存在或扩展名错误: {path}")
    try:
        import fitz
    except ImportError as exc:
        raise VisualValidationUnavailableError(
            "PDF 视觉比较需要可选依赖 PyMuPDF；"
            "请安装 word-format-tool[visual]。"
        ) from exc
    scale = dpi / 72
    matrix = fitz.Matrix(scale, scale)
    pages: list[RasterPage] = []
    try:
        with fitz.open(path) as document:
            for page in document:
                pixmap = page.get_pixmap(matrix=matrix, alpha=False)
                pages.append(
                    RasterPage(
                        width=pixmap.width,
                        height=pixmap.height,
                        pixels=bytes(pixmap.samples),
                        channels=pixmap.n,
                    )
                )
    except Exception as exc:
        raise VisualValidationUnavailableError(
            f"无法渲染 PDF {path}: {exc}"
        ) from exc
    return pages


def compare_pdf_layout(
    source_pdf: str | Path,
    candidate_pdf: str | Path,
    *,
    dpi: int = 120,
    changed_pixel_threshold: float = 0.005,
    page_loader: Callable[..., list[RasterPage]] = load_pdf_raster_pages,
) -> dict[str, Any]:
    """Render and compare two PDFs."""

    source_pages = page_loader(source_pdf, dpi=dpi)
    candidate_pages = page_loader(candidate_pdf, dpi=dpi)
    report = compare_raster_pages(
        source_pages,
        candidate_pages,
        changed_pixel_threshold=changed_pixel_threshold,
    )
    report.update(
        {
            "source": str(source_pdf),
            "candidate": str(candidate_pdf),
            "dpi": dpi,
        }
    )
    return report


class WordComPdfRenderer:
    """Export DOCX to PDF through an installed desktop Microsoft Word."""

    def render(self, docx_path: Path, pdf_path: Path) -> None:
        try:
            import win32com.client
        except ImportError as exc:
            raise VisualValidationUnavailableError(
                "DOCX 视觉比较需要 Microsoft Word 和可选依赖 pywin32；"
                "请安装 word-format-tool[visual]。"
            ) from exc

        word = None
        document = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            document = word.Documents.Open(
                str(docx_path.resolve()),
                ReadOnly=True,
                AddToRecentFiles=False,
            )
            document.ExportAsFixedFormat(
                str(pdf_path.resolve()),
                17,
                OpenAfterExport=False,
            )
        except Exception as exc:
            raise VisualValidationUnavailableError(
                f"Microsoft Word 导出 PDF 失败: {exc}"
            ) from exc
        finally:
            if document is not None:
                document.Close(False)
            if word is not None:
                word.Quit()


def compare_docx_visual(
    source_docx: str | Path,
    candidate_docx: str | Path,
    output_dir: str | Path,
    *,
    dpi: int = 120,
    changed_pixel_threshold: float = 0.005,
    renderer: DocxPdfRenderer | None = None,
    page_loader: Callable[..., list[RasterPage]] = load_pdf_raster_pages,
) -> dict[str, Any]:
    """Export two DOCX files to PDF and compare rendered pages."""

    source = require_docx_file(source_docx)
    candidate = require_docx_file(candidate_docx)
    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)
    source_pdf = output / f"{source.stem}-source.pdf"
    candidate_pdf = output / f"{candidate.stem}-candidate.pdf"
    active_renderer = renderer or WordComPdfRenderer()
    active_renderer.render(source, source_pdf)
    active_renderer.render(candidate, candidate_pdf)
    report = compare_pdf_layout(
        source_pdf,
        candidate_pdf,
        dpi=dpi,
        changed_pixel_threshold=changed_pixel_threshold,
        page_loader=page_loader,
    )
    report.update(
        {
            "source_docx": str(source),
            "candidate_docx": str(candidate),
            "rendered_pdfs": [str(source_pdf), str(candidate_pdf)],
        }
    )
    return report
