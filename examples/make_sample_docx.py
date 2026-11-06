"""Build a small, invented DOCX with tracked changes, using only the standard library.

    python examples/make_sample_docx.py sample.docx
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-'
    'officedocument.wordprocessingml.document.main+xml"/>'
    "</Types>"
)
ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)
REVISION = 'w:id="1" w:author="Example Author" w:date="2026-01-01T00:00:00Z"'
BODY = (
    "<w:p><w:r><w:t>TODO: check the figures before this goes out.</w:t></w:r></w:p>"
    "<w:p><w:r><w:t>The pilot covered twelve sites </w:t></w:r>"
    f"<w:del {REVISION}><w:r><w:delText>and failed </w:delText></w:r></w:del>"
    f"<w:ins {REVISION}><w:r><w:t>and met its target </w:t></w:r></w:ins>"
    "<w:r><w:t>in the first year.</w:t></w:r></w:p>"
    "<w:p><w:r><w:t>Cyber </w:t></w:r><w:r><w:t>security spending rose.</w:t></w:r></w:p>"
    "<w:p><w:r><w:t>Cybersecurity spending fell in one region.</w:t></w:r></w:p>"
)


def write_sample(path: Path) -> Path:
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}"><w:body>{BODY}<w:sectPr/></w:body></w:document>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/document.xml", document)
    return path


if __name__ == "__main__":
    print(write_sample(Path(sys.argv[1] if len(sys.argv) > 1 else "sample.docx")))
