from __future__ import annotations

from pathlib import Path

import pytest

from word_format_tool import convert_to_docx
from word_format_tool.exceptions import InputFileError


def test_convert_to_docx_rejects_unsupported_input(tmp_path: Path) -> None:
    source = tmp_path / "source.txt"
    source.write_text("not a Word document", encoding="utf-8")

    with pytest.raises(InputFileError, match="仅支持"):
        convert_to_docx(source, tmp_path / "output.docx")
