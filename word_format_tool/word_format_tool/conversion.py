"""Built-in conversion of legacy Word documents to DOCX."""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

from .document_io import copy_docx_atomic, prepare_output_path
from .exceptions import DocumentReadError, InputFileError

WORD_DOCX_FORMAT = 16


def _convert_with_word(source: Path, output: Path) -> None:
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DocumentReadError("当前 Python 未安装 pywin32，无法调用 Microsoft Word。") from exc

    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(str(source.resolve()), ReadOnly=True)
        document.SaveAs2(str(output.resolve()), FileFormat=WORD_DOCX_FORMAT)
    except Exception as exc:
        raise DocumentReadError(f"Microsoft Word 转换失败: {source}: {exc}") from exc
    finally:
        if document is not None:
            document.Close(False)
        if word is not None:
            word.Quit()


def _convert_with_libreoffice(source: Path, output: Path) -> None:
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if executable is None:
        raise DocumentReadError("未找到 Microsoft Word 或 LibreOffice 转换后端。")
    completed = subprocess.run(
        [
            executable,
            "--headless",
            "--convert-to",
            "docx",
            "--outdir",
            str(output.parent),
            str(source),
        ],
        check=False,
        capture_output=True,
        text=True,
        timeout=120,
    )
    generated = output.parent / f"{source.stem}.docx"
    if completed.returncode != 0 or not generated.is_file():
        message = completed.stderr.strip() or completed.stdout.strip()
        raise DocumentReadError(f"LibreOffice 转换失败: {message}")
    if generated.resolve() != output.resolve():
        generated.replace(output)


def convert_to_docx(
    input_path: str | Path,
    output_path: str | Path,
) -> dict[str, str]:
    """Convert a .doc or copy a .docx into a distinct DOCX output path."""

    source = Path(input_path)
    if not source.is_file():
        raise InputFileError(f"Word 文件不存在或不是文件: {source}")
    if source.suffix.lower() not in {".doc", ".docx"}:
        raise InputFileError(f"仅支持 .doc 或 .docx 文件: {source}")
    output = prepare_output_path(
        output_path,
        protected_input=source,
        suffix=".docx",
    )
    if source.suffix.lower() == ".docx":
        copy_docx_atomic(source, output)
        backend = "copy"
    else:
        try:
            _convert_with_word(source, output)
            backend = "microsoft-word"
        except DocumentReadError as word_error:
            try:
                _convert_with_libreoffice(source, output)
                backend = "libreoffice"
            except DocumentReadError as libreoffice_error:
                raise DocumentReadError(
                    f"{word_error}；备用后端也失败: {libreoffice_error}"
                ) from libreoffice_error
    return {
        "source": str(source),
        "output": str(output),
        "backend": backend,
    }
