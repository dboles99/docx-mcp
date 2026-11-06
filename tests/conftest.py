"""Every fixture here is an invented document built in code."""
from __future__ import annotations

import zipfile
from collections.abc import Callable
from pathlib import Path

import pytest

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
REVISION = 'w:id="1" w:author="ed" w:date="2026-01-01T00:00:00Z"'


def para(*runs: str) -> str:
    inner = "".join(f"<w:r><w:t>{run}</w:t></w:r>" for run in runs)
    return f"<w:p>{inner}</w:p>"


def write_docx(
    path: Path,
    body: str,
    *,
    comments: int | None = None,
    prolog: str = '<?xml version="1.0" encoding="UTF-8"?>',
) -> str:
    document = (
        f'{prolog}<w:document xmlns:w="{W}" xmlns:mc="{MC}">'
        f"<w:body>{body}</w:body></w:document>"
    )
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        )
        archive.writestr("word/document.xml", document)
        if comments is not None:
            items = "".join(f'<w:comment w:id="{i}"/>' for i in range(comments))
            archive.writestr("word/comments.xml", f'<w:comments xmlns:w="{W}">{items}</w:comments>')
    return str(path)


@pytest.fixture
def docx(tmp_path: Path) -> Callable[..., str]:
    counter = iter(range(1000))

    def build(body: str, **kwargs) -> str:
        return write_docx(tmp_path / f"d{next(counter)}.docx", body, **kwargs)

    return build
