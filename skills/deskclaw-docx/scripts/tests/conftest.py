"""Shared fixtures for deskclaw-docx tests."""

import json
import sys
import tempfile
from pathlib import Path

import pytest

_SCRIPTS_DIR = Path(__file__).resolve().parent.parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

from core import (
    create_document,
    add_heading,
    add_paragraph,
    add_table,
)


@pytest.fixture()
def tmp_docx(tmp_path):
    """An empty docx file."""
    f = str(tmp_path / "test.docx")
    create_document(f, title="Test", author="Tester")
    return f


@pytest.fixture()
def docx_with_paragraphs(tmp_path):
    """A docx with several paragraphs."""
    f = str(tmp_path / "paras.docx")
    create_document(f)
    add_heading(f, "Title Heading", level=1)
    add_paragraph(f, "First paragraph of the document.")
    add_paragraph(f, "Second paragraph with more text.")
    add_paragraph(f, "Third paragraph for testing.")
    return f


@pytest.fixture()
def docx_with_table(tmp_path):
    """A docx with a 3x3 table."""
    f = str(tmp_path / "table.docx")
    create_document(f)
    add_table(f, rows=3, cols=3, data=[
        ["A1", "B1", "C1"],
        ["A2", "B2", "C2"],
        ["A3", "B3", "C3"],
    ])
    return f


@pytest.fixture()
def tmp_image(tmp_path):
    """A minimal 1x1 PNG image for testing image-related features."""
    import struct, zlib
    img_path = tmp_path / "test.png"

    def _minimal_png():
        sig = b'\x89PNG\r\n\x1a\n'
        ihdr_data = struct.pack('>IIBBBBB', 1, 1, 8, 2, 0, 0, 0)
        ihdr_crc = zlib.crc32(b'IHDR' + ihdr_data) & 0xffffffff
        ihdr = struct.pack('>I', 13) + b'IHDR' + ihdr_data + struct.pack('>I', ihdr_crc)
        raw = b'\x00\x00\x00\x00'
        compressed = zlib.compress(raw)
        idat_crc = zlib.crc32(b'IDAT' + compressed) & 0xffffffff
        idat = struct.pack('>I', len(compressed)) + b'IDAT' + compressed + struct.pack('>I', idat_crc)
        iend_crc = zlib.crc32(b'IEND') & 0xffffffff
        iend = struct.pack('>I', 0) + b'IEND' + struct.pack('>I', iend_crc)
        return sig + ihdr + idat + iend

    img_path.write_bytes(_minimal_png())
    return str(img_path)


@pytest.fixture()
def word_tool_path():
    """Absolute path to word_tool.py CLI script."""
    return str(_SCRIPTS_DIR / "word_tool.py")
