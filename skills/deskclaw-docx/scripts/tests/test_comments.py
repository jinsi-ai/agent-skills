"""Tests for comments.py: add/delete/query comments."""

import json
import sys
import zipfile
from pathlib import Path

import pytest

_SCRIPTS_DIR = Path(__file__).resolve().parent.parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

from core import (
    add_comment,
    delete_comment,
    get_all_comments,
    get_comments_by_author,
    get_comments_for_paragraph,
)


class TestComments:
    def test_add_comment(self, docx_with_paragraphs):
        result = add_comment(
            docx_with_paragraphs,
            paragraph_index=1,
            comment_text="Please verify this data",
            author="Reviewer",
        )
        assert "Added comment" in result

        comments = json.loads(get_all_comments(docx_with_paragraphs))
        assert len(comments) >= 1
        assert comments[0]["text"] == "Please verify this data"
        assert comments[0]["author"] == "Reviewer"

    def test_add_comment_creates_comments_xml(self, tmp_docx):
        """Fresh doc has no comments.xml; add_comment should create it."""
        from core import add_paragraph
        add_paragraph(tmp_docx, "Test paragraph")
        result = add_comment(tmp_docx, paragraph_index=0, comment_text="Auto-created", author="Bot")
        assert "Added comment" in result

        with zipfile.ZipFile(tmp_docx, "r") as z:
            assert "word/comments.xml" in z.namelist()

    def test_delete_comment(self, docx_with_paragraphs):
        add_comment(docx_with_paragraphs, paragraph_index=1, comment_text="To delete", author="A")
        comments = json.loads(get_all_comments(docx_with_paragraphs))
        assert len(comments) >= 1
        cid = comments[0]["id"]

        result = delete_comment(docx_with_paragraphs, comment_id=cid)
        assert "Deleted" in result

        comments_after = json.loads(get_all_comments(docx_with_paragraphs))
        ids_after = [c["id"] for c in comments_after]
        assert cid not in ids_after

    def test_get_comments_by_author(self, docx_with_paragraphs):
        add_comment(docx_with_paragraphs, paragraph_index=1, comment_text="C1", author="Alice")
        add_comment(docx_with_paragraphs, paragraph_index=2, comment_text="C2", author="Bob")

        alice = json.loads(get_comments_by_author(docx_with_paragraphs, author="Alice"))
        assert len(alice) >= 1
        assert all(c["author"] == "Alice" for c in alice)

    def test_get_comments_for_paragraph(self, docx_with_paragraphs):
        add_comment(docx_with_paragraphs, paragraph_index=2, comment_text="Paragraph-level", author="X")
        result = json.loads(get_comments_for_paragraph(docx_with_paragraphs, paragraph_index=2))
        assert len(result) >= 1
        assert result[0]["text"] == "Paragraph-level"
