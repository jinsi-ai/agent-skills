"""CLI integration tests: run word_tool.py via subprocess."""

import json
import subprocess
import sys
from pathlib import Path

import pytest

_SCRIPTS_DIR = Path(__file__).resolve().parent.parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))


def _run(word_tool_path, *args):
    result = subprocess.run(
        [sys.executable, word_tool_path, *args],
        capture_output=True, text=True, timeout=30,
    )
    return result


class TestCLIDocumentAndPageSetup:
    def test_create_and_set_page_size(self, tmp_path, word_tool_path):
        f = str(tmp_path / "cli_test.docx")
        r = _run(word_tool_path, "create_document", f, "--title", "CLI Test")
        assert r.returncode == 0
        assert "Created" in r.stdout

        r = _run(word_tool_path, "set_page_size", f, "--paper", "Letter")
        assert r.returncode == 0
        assert "Set page size" in r.stdout

        r = _run(word_tool_path, "get_page_settings", f)
        assert r.returncode == 0
        settings = json.loads(r.stdout)
        assert abs(settings["page_width_inches"] - 8.5) < 0.1


class TestCLIParagraph:
    def test_add_and_align_paragraph(self, tmp_path, word_tool_path):
        f = str(tmp_path / "cli_para.docx")
        _run(word_tool_path, "create_document", f)
        _run(word_tool_path, "add_paragraph", f, "Test paragraph")

        r = _run(word_tool_path, "set_paragraph_alignment", f,
                 "--paragraph-index", "0", "--alignment", "center")
        assert r.returncode == 0
        assert "alignment" in r.stdout


class TestCLITable:
    def test_add_table_and_set_cell(self, tmp_path, word_tool_path):
        f = str(tmp_path / "cli_table.docx")
        _run(word_tool_path, "create_document", f)
        data = json.dumps([["H1", "H2"], ["V1", "V2"]])
        _run(word_tool_path, "add_table", f, "--rows", "2", "--cols", "2", "--data", data)

        r = _run(word_tool_path, "set_table_cell", f,
                 "--table-index", "0", "--row", "0", "--col", "1", "--text", "UPDATED")
        assert r.returncode == 0
        assert "Set cell" in r.stdout

        r = _run(word_tool_path, "get_table_info", f, "--table-index", "0")
        assert r.returncode == 0
        info = json.loads(r.stdout)
        assert info["rows"] == 2


class TestCLIHyperlink:
    def test_add_hyperlink(self, tmp_path, word_tool_path):
        f = str(tmp_path / "cli_link.docx")
        _run(word_tool_path, "create_document", f)

        r = _run(word_tool_path, "add_hyperlink", f, "Google", "https://google.com")
        assert r.returncode == 0
        assert "Added hyperlink" in r.stdout

        r = _run(word_tool_path, "get_hyperlinks", f)
        assert r.returncode == 0
        links = json.loads(r.stdout)
        assert len(links) >= 1
        assert links[0]["url"] == "https://google.com"


class TestCLIComments:
    def test_add_and_get_comments(self, tmp_path, word_tool_path):
        f = str(tmp_path / "cli_comment.docx")
        _run(word_tool_path, "create_document", f)
        _run(word_tool_path, "add_paragraph", f, "Content paragraph")

        r = _run(word_tool_path, "add_comment", f,
                 "--paragraph-index", "0", "--text", "Review this", "--author", "Tester")
        assert r.returncode == 0
        assert "Added comment" in r.stdout

        r = _run(word_tool_path, "get_all_comments", f)
        assert r.returncode == 0
        comments = json.loads(r.stdout)
        assert len(comments) >= 1
        assert comments[0]["text"] == "Review this"
