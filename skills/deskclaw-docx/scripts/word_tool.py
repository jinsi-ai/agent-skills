#!/usr/bin/env python3
"""CLI for Word document operations (deskclaw-docx skill)."""

import argparse
import json
import sys
from pathlib import Path

# Run from skill scripts dir; add parent so "core" is importable
_SCRIPT_DIR = Path(__file__).resolve().parent
if str(_SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPT_DIR))

from core import (
    create_document,
    get_document_info,
    get_document_text,
    get_document_outline,
    list_available_documents,
    copy_document,
    convert_to_pdf,
    merge_documents,
    add_header,
    add_footer,
    add_page_number,
    set_different_first_page_header,
    add_header_image,
    add_table_of_contents,
    update_table_of_contents,
    set_page_size,
    set_page_margins,
    set_page_orientation,
    get_page_settings,
    add_watermark,
    add_watermark_image,
    add_footnote,
    add_endnote,
    add_heading,
    add_paragraph,
    add_table,
    add_picture,
    add_page_break,
    insert_header_near_text,
    insert_line_or_paragraph_near_text,
    insert_numbered_list_near_text,
    get_table_info,
    set_table_cell,
    batch_set_table_cells,
    add_table_row,
    add_table_rows,
    delete_table_row,
    delete_table_rows,
    add_table_column,
    delete_table_column,
    add_hyperlink,
    add_hyperlink_to_text,
    get_hyperlinks,
    format_text,
    search_and_replace,
    delete_paragraph,
    create_custom_style,
    format_table,
    set_table_cell_shading,
    apply_table_alternating_rows,
    highlight_table_header,
    format_table_cell_text,
    set_table_cell_padding,
    set_table_cell_alignment,
    set_table_alignment_all,
    merge_table_cells,
    merge_table_cells_horizontal,
    merge_table_cells_vertical,
    set_table_column_width,
    set_table_column_widths,
    set_table_width,
    auto_fit_table_columns,
    set_paragraph_alignment,
    set_paragraph_indent,
    set_paragraph_spacing,
    format_all_paragraphs,
    get_all_comments,
    get_comments_by_author,
    get_comments_for_paragraph,
    add_comment,
    delete_comment,
    add_password_protection,
    add_restricted_editing,
)


def _bool_arg(s):
    return s.lower() in ("true", "1", "yes")


def main():
    parser = argparse.ArgumentParser(description="Word document CLI (deskclaw-docx)")
    sub = parser.add_subparsers(dest="command", required=True)

    # document
    p = sub.add_parser("create_document")
    p.add_argument("filename")
    p.add_argument("--title", default=None)
    p.add_argument("--author", default=None)

    p = sub.add_parser("get_document_info")
    p.add_argument("filename")

    p = sub.add_parser("get_document_text")
    p.add_argument("filename")

    p = sub.add_parser("get_document_outline")
    p.add_argument("filename")

    p = sub.add_parser("list_available_documents")
    p.add_argument("--directory", default=".")

    p = sub.add_parser("copy_document")
    p.add_argument("source_filename")
    p.add_argument("--destination", default=None, dest="destination_filename")

    p = sub.add_parser("convert_to_pdf")
    p.add_argument("filename")
    p.add_argument("--output", default=None, dest="output_filename")

    p = sub.add_parser("merge_documents")
    p.add_argument("filename")
    p.add_argument("--sources", nargs="+", required=True, dest="source_files")

    # header/footer
    p = sub.add_parser("add_header")
    p.add_argument("filename")
    p.add_argument("text")
    p.add_argument("--font-name", default=None)
    p.add_argument("--font-size", type=int, default=None)
    p.add_argument("--bold", type=_bool_arg, default=None)
    p.add_argument("--alignment", default="center", choices=["left", "center", "right"])

    p = sub.add_parser("add_footer")
    p.add_argument("filename")
    p.add_argument("text")
    p.add_argument("--font-name", default=None)
    p.add_argument("--font-size", type=int, default=None)
    p.add_argument("--bold", type=_bool_arg, default=None)
    p.add_argument("--alignment", default="center", choices=["left", "center", "right"])

    p = sub.add_parser("add_page_number")
    p.add_argument("filename")
    p.add_argument("--position", default="footer", choices=["header", "footer"])
    p.add_argument("--alignment", default="center", choices=["left", "center", "right"])
    p.add_argument("--format", default="第 {page} 页", dest="format_text")

    p = sub.add_parser("set_different_first_page_header")
    p.add_argument("filename")
    p.add_argument("--enabled", type=_bool_arg, default=True)

    p = sub.add_parser("add_header_image")
    p.add_argument("filename")
    p.add_argument("image_path")
    p.add_argument("--width", type=float, default=None)
    p.add_argument("--alignment", default="center", choices=["left", "center", "right"])

    # page setup
    p = sub.add_parser("set_page_size")
    p.add_argument("filename")
    p.add_argument("--paper", default="A4", help="A4, A3, Letter, Legal, B5, or custom")
    p.add_argument("--width", type=float, default=None, help="Width in inches (for custom)")
    p.add_argument("--height", type=float, default=None, help="Height in inches (for custom)")

    p = sub.add_parser("set_page_margins")
    p.add_argument("filename")
    p.add_argument("--top", type=float, default=None, help="Top margin in inches")
    p.add_argument("--bottom", type=float, default=None, help="Bottom margin in inches")
    p.add_argument("--left", type=float, default=None, help="Left margin in inches")
    p.add_argument("--right", type=float, default=None, help="Right margin in inches")

    p = sub.add_parser("set_page_orientation")
    p.add_argument("filename")
    p.add_argument("--orientation", default="portrait", choices=["portrait", "landscape"])

    p = sub.add_parser("get_page_settings")
    p.add_argument("filename")

    # watermark
    p = sub.add_parser("add_watermark")
    p.add_argument("filename")
    p.add_argument("text")
    p.add_argument("--font-name", default="Arial")
    p.add_argument("--font-size", type=int, default=48)
    p.add_argument("--color", default="C0C0C0")
    p.add_argument("--diagonal", type=_bool_arg, default=True)

    p = sub.add_parser("add_watermark_image")
    p.add_argument("filename")
    p.add_argument("image_path")
    p.add_argument("--width", type=float, default=None)

    # footnotes/endnotes
    p = sub.add_parser("add_footnote")
    p.add_argument("filename")
    p.add_argument("--paragraph-index", type=int, required=True)
    p.add_argument("--text", required=True, dest="footnote_text")

    p = sub.add_parser("add_endnote")
    p.add_argument("filename")
    p.add_argument("--paragraph-index", type=int, required=True)
    p.add_argument("--text", required=True, dest="endnote_text")

    # table of contents
    p = sub.add_parser("add_table_of_contents")
    p.add_argument("filename")
    p.add_argument("--title", default="目录")
    p.add_argument("--max-level", type=int, default=3, dest="max_level")
    p.add_argument("--position", default="start", help="'start' or paragraph index")

    p = sub.add_parser("update_table_of_contents")
    p.add_argument("filename")

    # content
    p = sub.add_parser("add_heading")
    p.add_argument("filename")
    p.add_argument("text")
    p.add_argument("--level", type=int, default=1)
    p.add_argument("--font-name", default=None)
    p.add_argument("--font-size", type=int, default=None)
    p.add_argument("--bold", type=_bool_arg, default=None)
    p.add_argument("--italic", type=_bool_arg, default=None)
    p.add_argument("--border-bottom", type=_bool_arg, default=False)

    p = sub.add_parser("add_paragraph")
    p.add_argument("filename")
    p.add_argument("text")
    p.add_argument("--style", default=None)
    p.add_argument("--font-name", default=None)
    p.add_argument("--font-size", type=int, default=None)
    p.add_argument("--bold", type=_bool_arg, default=None)
    p.add_argument("--italic", type=_bool_arg, default=None)
    p.add_argument("--color", default=None)

    p = sub.add_parser("add_table")
    p.add_argument("filename")
    p.add_argument("--rows", type=int, required=True)
    p.add_argument("--cols", type=int, required=True)
    p.add_argument("--data", default=None, help="JSON array of rows")

    p = sub.add_parser("add_picture")
    p.add_argument("filename")
    p.add_argument("image_path")
    p.add_argument("--width", type=float, default=None)

    p = sub.add_parser("add_page_break")
    p.add_argument("filename")

    p = sub.add_parser("insert_header_near_text")
    p.add_argument("filename")
    p.add_argument("--target-text", default=None)
    p.add_argument("--header-title", default=None)
    p.add_argument("--position", default="after")
    p.add_argument("--header-style", default="Heading 1")
    p.add_argument("--target-paragraph-index", type=int, default=None)

    p = sub.add_parser("insert_line_or_paragraph_near_text")
    p.add_argument("filename")
    p.add_argument("--target-text", default=None)
    p.add_argument("--line-text", default=None)
    p.add_argument("--position", default="after")
    p.add_argument("--line-style", default=None)
    p.add_argument("--target-paragraph-index", type=int, default=None)

    p = sub.add_parser("insert_numbered_list_near_text")
    p.add_argument("filename")
    p.add_argument("--target-text", default=None)
    p.add_argument("--list-items", required=True, type=json.loads)
    p.add_argument("--position", default="after")
    p.add_argument("--target-paragraph-index", type=int, default=None)
    p.add_argument("--bullet-type", default="bullet")

    # table cell operations
    p = sub.add_parser("get_table_info")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, default=None)
    p.add_argument("--show-merged", type=_bool_arg, default=True, help="Show merged cell info")

    p = sub.add_parser("set_table_cell")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--row", type=int, required=True, dest="row_index")
    p.add_argument("--col", type=int, required=True, dest="col_index")
    p.add_argument("--text", required=True)
    p.add_argument("--visual", type=_bool_arg, default=False, dest="use_visual_index", help="Use visual column index (skip merged cells)")

    p = sub.add_parser("batch_set_table_cells")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--cells", required=True, type=json.loads, dest="cells_data", help='JSON array: [{"row":0,"col":1,"text":"value"},...]')
    p.add_argument("--visual", type=_bool_arg, default=False, dest="use_visual_index", help="Use visual column index (skip merged cells)")

    # table row/column operations
    p = sub.add_parser("add_table_row")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--position", default="end", help="'end', 'start', or row index")
    p.add_argument("--copy-style-from", type=int, default=None)

    p = sub.add_parser("add_table_rows")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--count", type=int, required=True)
    p.add_argument("--position", default="end", help="'end', 'start', or row index")

    p = sub.add_parser("delete_table_row")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--row", type=int, required=True, dest="row_index")

    p = sub.add_parser("delete_table_rows")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--rows", type=json.loads, required=True, dest="row_indices", help="JSON array of row indices")

    p = sub.add_parser("add_table_column")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--position", default="end", help="'end', 'start', or column index")
    p.add_argument("--width", type=float, default=None, help="Column width in inches")

    p = sub.add_parser("delete_table_column")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--col", type=int, required=True, dest="col_index")

    # hyperlinks
    p = sub.add_parser("add_hyperlink")
    p.add_argument("filename")
    p.add_argument("text")
    p.add_argument("url")
    p.add_argument("--paragraph-index", type=int, default=None)
    p.add_argument("--color", default="0000FF")
    p.add_argument("--underline", type=_bool_arg, default=True)

    p = sub.add_parser("add_hyperlink_to_text")
    p.add_argument("filename")
    p.add_argument("--find", required=True, dest="find_text")
    p.add_argument("--url", required=True)
    p.add_argument("--color", default="0000FF")
    p.add_argument("--underline", type=_bool_arg, default=True)

    p = sub.add_parser("get_hyperlinks")
    p.add_argument("filename")

    # formatting
    p = sub.add_parser("format_text")
    p.add_argument("filename")
    p.add_argument("--paragraph-index", type=int, required=True)
    p.add_argument("--start-pos", type=int, required=True)
    p.add_argument("--end-pos", type=int, required=True)
    p.add_argument("--bold", type=_bool_arg, default=None)
    p.add_argument("--italic", type=_bool_arg, default=None)
    p.add_argument("--underline", type=_bool_arg, default=None)
    p.add_argument("--color", default=None)
    p.add_argument("--font-size", type=int, default=None)
    p.add_argument("--font-name", default=None)

    p = sub.add_parser("search_and_replace")
    p.add_argument("filename")
    p.add_argument("--find", required=True, dest="find_text")
    p.add_argument("--replace", required=True, dest="replace_text")

    p = sub.add_parser("delete_paragraph")
    p.add_argument("filename")
    p.add_argument("--paragraph-index", type=int, required=True)

    p = sub.add_parser("create_custom_style")
    p.add_argument("filename")
    p.add_argument("--style-name", required=True)
    p.add_argument("--bold", type=_bool_arg, default=None)
    p.add_argument("--italic", type=_bool_arg, default=None)
    p.add_argument("--font-size", type=int, default=None)
    p.add_argument("--font-name", default=None)
    p.add_argument("--color", default=None)
    p.add_argument("--base-style", default=None)

    p = sub.add_parser("format_table")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--has-header-row", type=_bool_arg, default=None)
    p.add_argument("--border-style", default=None)
    p.add_argument("--shading", default=None)

    p = sub.add_parser("set_table_cell_shading")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--row-index", type=int, required=True)
    p.add_argument("--col-index", type=int, required=True)
    p.add_argument("--fill-color", required=True)
    p.add_argument("--pattern", default="clear")

    p = sub.add_parser("apply_table_alternating_rows")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--color1", default="FFFFFF")
    p.add_argument("--color2", default="F2F2F2")

    p = sub.add_parser("highlight_table_header")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--header-color", default="4472C4")
    p.add_argument("--text-color", default="FFFFFF")

    p = sub.add_parser("format_table_cell_text")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--row-index", type=int, required=True)
    p.add_argument("--col-index", type=int, required=True)
    p.add_argument("--text-content", default=None)
    p.add_argument("--bold", type=_bool_arg, default=None)
    p.add_argument("--italic", type=_bool_arg, default=None)
    p.add_argument("--underline", type=_bool_arg, default=None)
    p.add_argument("--color", default=None)
    p.add_argument("--font-size", type=int, default=None)
    p.add_argument("--font-name", default=None)

    p = sub.add_parser("set_table_cell_padding")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--row-index", type=int, required=True)
    p.add_argument("--col-index", type=int, required=True)
    p.add_argument("--top", type=int, default=None)
    p.add_argument("--bottom", type=int, default=None)
    p.add_argument("--left", type=int, default=None)
    p.add_argument("--right", type=int, default=None)
    p.add_argument("--unit", default="points")

    p = sub.add_parser("set_table_cell_alignment")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--row-index", type=int, required=True)
    p.add_argument("--col-index", type=int, required=True)
    p.add_argument("--horizontal", default="left")
    p.add_argument("--vertical", default="top")

    p = sub.add_parser("set_table_alignment_all")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--horizontal", default="left")
    p.add_argument("--vertical", default="top")

    p = sub.add_parser("merge_table_cells")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--start-row", type=int, required=True)
    p.add_argument("--start-col", type=int, required=True)
    p.add_argument("--end-row", type=int, required=True)
    p.add_argument("--end-col", type=int, required=True)

    p = sub.add_parser("merge_table_cells_horizontal")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--row-index", type=int, required=True)
    p.add_argument("--start-col", type=int, required=True)
    p.add_argument("--end-col", type=int, required=True)

    p = sub.add_parser("merge_table_cells_vertical")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--col-index", type=int, required=True)
    p.add_argument("--start-row", type=int, required=True)
    p.add_argument("--end-row", type=int, required=True)

    p = sub.add_parser("set_table_column_width")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--col-index", type=int, required=True)
    p.add_argument("--width", type=float, required=True)
    p.add_argument("--width-type", default="points")

    p = sub.add_parser("set_table_column_widths")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--widths", type=json.loads, required=True)
    p.add_argument("--width-type", default="points")

    p = sub.add_parser("set_table_width")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)
    p.add_argument("--width", type=float, required=True)
    p.add_argument("--width-type", default="points")

    p = sub.add_parser("auto_fit_table_columns")
    p.add_argument("filename")
    p.add_argument("--table-index", type=int, required=True)

    # paragraph formatting
    p = sub.add_parser("set_paragraph_alignment")
    p.add_argument("filename")
    p.add_argument("--paragraph-index", type=int, required=True)
    p.add_argument("--alignment", required=True, choices=["left", "center", "right", "justify"])

    p = sub.add_parser("set_paragraph_indent")
    p.add_argument("filename")
    p.add_argument("--paragraph-index", type=int, required=True)
    p.add_argument("--left", type=float, default=None, help="Left indent in points")
    p.add_argument("--right", type=float, default=None, help="Right indent in points")
    p.add_argument("--first-line", type=float, default=None, dest="first_line", help="First line indent in points")
    p.add_argument("--hanging", type=float, default=None, help="Hanging indent in points")

    p = sub.add_parser("set_paragraph_spacing")
    p.add_argument("filename")
    p.add_argument("--paragraph-index", type=int, required=True)
    p.add_argument("--before", type=float, default=None, help="Space before in points")
    p.add_argument("--after", type=float, default=None, help="Space after in points")
    p.add_argument("--line-spacing", type=float, default=None, dest="line_spacing")
    p.add_argument("--line-spacing-rule", default=None, dest="line_spacing_rule", choices=["single", "one_point_five", "double", "exactly", "at_least", "multiple"])

    p = sub.add_parser("format_all_paragraphs")
    p.add_argument("filename")
    p.add_argument("--alignment", default=None, choices=["left", "center", "right", "justify"])
    p.add_argument("--left-indent", type=float, default=None, dest="left_indent")
    p.add_argument("--first-line-indent", type=float, default=None, dest="first_line_indent")
    p.add_argument("--space-before", type=float, default=None, dest="space_before")
    p.add_argument("--space-after", type=float, default=None, dest="space_after")
    p.add_argument("--line-spacing", type=float, default=None, dest="line_spacing")

    # comments
    p = sub.add_parser("get_all_comments")
    p.add_argument("filename")

    p = sub.add_parser("get_comments_by_author")
    p.add_argument("filename")
    p.add_argument("--author", required=True)

    p = sub.add_parser("get_comments_for_paragraph")
    p.add_argument("filename")
    p.add_argument("--paragraph-index", type=int, required=True)

    p = sub.add_parser("add_comment")
    p.add_argument("filename")
    p.add_argument("--paragraph-index", type=int, required=True)
    p.add_argument("--text", required=True, dest="comment_text")
    p.add_argument("--author", default="AI Assistant")

    p = sub.add_parser("delete_comment")
    p.add_argument("filename")
    p.add_argument("--comment-id", required=True, dest="comment_id")

    # protection
    p = sub.add_parser("add_password_protection")
    p.add_argument("filename")
    p.add_argument("--password", required=True)
    p.add_argument("--output", default=None, dest="output_filename")

    p = sub.add_parser("add_restricted_editing")
    p.add_argument("filename")
    p.add_argument("--password", default=None)
    p.add_argument("--output", default=None, dest="output_filename")

    args = parser.parse_args()
    cmd = args.command
    kwargs = {k: v for k, v in vars(args).items() if k != "command" and v is not None}

    try:
        if cmd == "create_document":
            out = create_document(args.filename, title=kwargs.get("title"), author=kwargs.get("author"))
        elif cmd == "get_document_info":
            out = get_document_info(args.filename)
        elif cmd == "get_document_text":
            out = get_document_text(args.filename)
        elif cmd == "get_document_outline":
            out = get_document_outline(args.filename)
        elif cmd == "list_available_documents":
            out = list_available_documents(kwargs.get("directory", "."))
        elif cmd == "copy_document":
            out = copy_document(args.source_filename, kwargs.get("destination_filename"))
        elif cmd == "convert_to_pdf":
            out = convert_to_pdf(args.filename, kwargs.get("output_filename"))
        elif cmd == "merge_documents":
            out = merge_documents(args.filename, args.source_files)
        elif cmd == "add_header":
            out = add_header(
                args.filename, args.text,
                font_name=kwargs.get("font_name"),
                font_size=kwargs.get("font_size"),
                bold=kwargs.get("bold"),
                alignment=kwargs.get("alignment", "center"),
            )
        elif cmd == "add_footer":
            out = add_footer(
                args.filename, args.text,
                font_name=kwargs.get("font_name"),
                font_size=kwargs.get("font_size"),
                bold=kwargs.get("bold"),
                alignment=kwargs.get("alignment", "center"),
            )
        elif cmd == "add_page_number":
            out = add_page_number(
                args.filename,
                position=kwargs.get("position", "footer"),
                alignment=kwargs.get("alignment", "center"),
                format_text=kwargs.get("format_text", "第 {page} 页"),
            )
        elif cmd == "set_different_first_page_header":
            out = set_different_first_page_header(args.filename, enabled=kwargs.get("enabled", True))
        elif cmd == "add_header_image":
            out = add_header_image(
                args.filename, args.image_path,
                width=kwargs.get("width"),
                alignment=kwargs.get("alignment", "center"),
            )
        elif cmd == "set_page_size":
            out = set_page_size(
                args.filename,
                width=kwargs.get("width"),
                height=kwargs.get("height"),
                paper=kwargs.get("paper", "A4"),
            )
        elif cmd == "set_page_margins":
            out = set_page_margins(
                args.filename,
                top=kwargs.get("top"),
                bottom=kwargs.get("bottom"),
                left=kwargs.get("left"),
                right=kwargs.get("right"),
            )
        elif cmd == "set_page_orientation":
            out = set_page_orientation(args.filename, orientation=kwargs.get("orientation", "portrait"))
        elif cmd == "get_page_settings":
            out = get_page_settings(args.filename)
        elif cmd == "add_watermark":
            out = add_watermark(
                args.filename, args.text,
                font_name=kwargs.get("font_name", "Arial"),
                font_size=kwargs.get("font_size", 48),
                color=kwargs.get("color", "C0C0C0"),
                diagonal=kwargs.get("diagonal", True),
            )
        elif cmd == "add_watermark_image":
            out = add_watermark_image(args.filename, args.image_path, width=kwargs.get("width"))
        elif cmd == "add_footnote":
            out = add_footnote(args.filename, args.paragraph_index, args.footnote_text)
        elif cmd == "add_endnote":
            out = add_endnote(args.filename, args.paragraph_index, args.endnote_text)
        elif cmd == "add_table_of_contents":
            pos = kwargs.get("position", "start")
            if pos != "start":
                try:
                    pos = int(pos)
                except ValueError:
                    pos = "start"
            out = add_table_of_contents(
                args.filename,
                title=kwargs.get("title", "目录"),
                max_level=kwargs.get("max_level", 3),
                position=pos,
            )
        elif cmd == "update_table_of_contents":
            out = update_table_of_contents(args.filename)
        elif cmd == "add_heading":
            out = add_heading(
                args.filename, args.text,
                level=kwargs.get("level", 1),
                font_name=kwargs.get("font_name"),
                font_size=kwargs.get("font_size"),
                bold=kwargs.get("bold"),
                italic=kwargs.get("italic"),
                border_bottom=kwargs.get("border_bottom", False),
            )
        elif cmd == "add_paragraph":
            out = add_paragraph(
                args.filename, args.text,
                style=kwargs.get("style"),
                font_name=kwargs.get("font_name"),
                font_size=kwargs.get("font_size"),
                bold=kwargs.get("bold"),
                italic=kwargs.get("italic"),
                color=kwargs.get("color"),
            )
        elif cmd == "add_table":
            data = json.loads(args.data) if args.data else None
            out = add_table(args.filename, args.rows, args.cols, data=data)
        elif cmd == "add_picture":
            out = add_picture(args.filename, args.image_path, width=kwargs.get("width"))
        elif cmd == "add_page_break":
            out = add_page_break(args.filename)
        elif cmd == "insert_header_near_text":
            out = insert_header_near_text(
                args.filename,
                target_text=kwargs.get("target_text"),
                header_title=kwargs.get("header_title"),
                position=kwargs.get("position", "after"),
                header_style=kwargs.get("header_style", "Heading 1"),
                target_paragraph_index=kwargs.get("target_paragraph_index"),
            )
        elif cmd == "insert_line_or_paragraph_near_text":
            out = insert_line_or_paragraph_near_text(
                args.filename,
                target_text=kwargs.get("target_text"),
                line_text=kwargs.get("line_text"),
                position=kwargs.get("position", "after"),
                line_style=kwargs.get("line_style"),
                target_paragraph_index=kwargs.get("target_paragraph_index"),
            )
        elif cmd == "insert_numbered_list_near_text":
            out = insert_numbered_list_near_text(
                args.filename,
                target_text=kwargs.get("target_text"),
                list_items=args.list_items,
                position=kwargs.get("position", "after"),
                target_paragraph_index=kwargs.get("target_paragraph_index"),
                bullet_type=kwargs.get("bullet_type", "bullet"),
            )
        elif cmd == "get_table_info":
            out = get_table_info(args.filename, table_index=kwargs.get("table_index"), show_merged=kwargs.get("show_merged", True))
        elif cmd == "set_table_cell":
            out = set_table_cell(args.filename, args.table_index, args.row_index, args.col_index, args.text, use_visual_index=kwargs.get("use_visual_index", False))
        elif cmd == "batch_set_table_cells":
            out = batch_set_table_cells(args.filename, args.table_index, args.cells_data, use_visual_index=kwargs.get("use_visual_index", False))
        elif cmd == "add_table_row":
            out = add_table_row(args.filename, args.table_index, position=kwargs.get("position", "end"), copy_style_from=kwargs.get("copy_style_from"))
        elif cmd == "add_table_rows":
            out = add_table_rows(args.filename, args.table_index, args.count, position=kwargs.get("position", "end"))
        elif cmd == "delete_table_row":
            out = delete_table_row(args.filename, args.table_index, args.row_index)
        elif cmd == "delete_table_rows":
            out = delete_table_rows(args.filename, args.table_index, args.row_indices)
        elif cmd == "add_table_column":
            out = add_table_column(args.filename, args.table_index, position=kwargs.get("position", "end"), width=kwargs.get("width"))
        elif cmd == "delete_table_column":
            out = delete_table_column(args.filename, args.table_index, args.col_index)
        elif cmd == "add_hyperlink":
            out = add_hyperlink(
                args.filename, args.text, args.url,
                paragraph_index=kwargs.get("paragraph_index"),
                color=kwargs.get("color", "0000FF"),
                underline=kwargs.get("underline", True),
            )
        elif cmd == "add_hyperlink_to_text":
            out = add_hyperlink_to_text(
                args.filename, args.find_text, args.url,
                color=kwargs.get("color", "0000FF"),
                underline=kwargs.get("underline", True),
            )
        elif cmd == "get_hyperlinks":
            out = get_hyperlinks(args.filename)
        elif cmd == "format_text":
            out = format_text(
                args.filename,
                args.paragraph_index,
                args.start_pos,
                args.end_pos,
                bold=kwargs.get("bold"),
                italic=kwargs.get("italic"),
                underline=kwargs.get("underline"),
                color=kwargs.get("color"),
                font_size=kwargs.get("font_size"),
                font_name=kwargs.get("font_name"),
            )
        elif cmd == "search_and_replace":
            out = search_and_replace(args.filename, args.find_text, args.replace_text)
        elif cmd == "delete_paragraph":
            out = delete_paragraph(args.filename, args.paragraph_index)
        elif cmd == "create_custom_style":
            out = create_custom_style(
                args.filename,
                args.style_name,
                bold=kwargs.get("bold"),
                italic=kwargs.get("italic"),
                font_size=kwargs.get("font_size"),
                font_name=kwargs.get("font_name"),
                color=kwargs.get("color"),
                base_style=kwargs.get("base_style"),
            )
        elif cmd == "format_table":
            out = format_table(
                args.filename,
                args.table_index,
                has_header_row=kwargs.get("has_header_row"),
                border_style=kwargs.get("border_style"),
                shading=kwargs.get("shading"),
            )
        elif cmd == "set_table_cell_shading":
            out = set_table_cell_shading(
                args.filename, args.table_index, args.row_index, args.col_index,
                args.fill_color, pattern=kwargs.get("pattern", "clear"),
            )
        elif cmd == "apply_table_alternating_rows":
            out = apply_table_alternating_rows(
                args.filename, args.table_index,
                color1=kwargs.get("color1", "FFFFFF"),
                color2=kwargs.get("color2", "F2F2F2"),
            )
        elif cmd == "highlight_table_header":
            out = highlight_table_header(
                args.filename, args.table_index,
                header_color=kwargs.get("header_color", "4472C4"),
                text_color=kwargs.get("text_color", "FFFFFF"),
            )
        elif cmd == "format_table_cell_text":
            out = format_table_cell_text(
                args.filename, args.table_index, args.row_index, args.col_index,
                text_content=kwargs.get("text_content"),
                bold=kwargs.get("bold"),
                italic=kwargs.get("italic"),
                underline=kwargs.get("underline"),
                color=kwargs.get("color"),
                font_size=kwargs.get("font_size"),
                font_name=kwargs.get("font_name"),
            )
        elif cmd == "set_table_cell_padding":
            out = set_table_cell_padding(
                args.filename, args.table_index, args.row_index, args.col_index,
                top=kwargs.get("top"), bottom=kwargs.get("bottom"),
                left=kwargs.get("left"), right=kwargs.get("right"),
                unit=kwargs.get("unit", "points"),
            )
        elif cmd == "set_table_cell_alignment":
            out = set_table_cell_alignment(
                args.filename, args.table_index, args.row_index, args.col_index,
                horizontal=kwargs.get("horizontal", "left"),
                vertical=kwargs.get("vertical", "top"),
            )
        elif cmd == "set_table_alignment_all":
            out = set_table_alignment_all(
                args.filename, args.table_index,
                horizontal=kwargs.get("horizontal", "left"),
                vertical=kwargs.get("vertical", "top"),
            )
        elif cmd == "merge_table_cells":
            out = merge_table_cells(
                args.filename, args.table_index,
                args.start_row, args.start_col, args.end_row, args.end_col,
            )
        elif cmd == "merge_table_cells_horizontal":
            out = merge_table_cells_horizontal(
                args.filename, args.table_index, args.row_index,
                args.start_col, args.end_col,
            )
        elif cmd == "merge_table_cells_vertical":
            out = merge_table_cells_vertical(
                args.filename, args.table_index, args.col_index,
                args.start_row, args.end_row,
            )
        elif cmd == "set_table_column_width":
            out = set_table_column_width(
                args.filename, args.table_index, args.col_index,
                args.width, width_type=kwargs.get("width_type", "points"),
            )
        elif cmd == "set_table_column_widths":
            out = set_table_column_widths(
                args.filename, args.table_index,
                args.widths, width_type=kwargs.get("width_type", "points"),
            )
        elif cmd == "set_table_width":
            out = set_table_width(
                args.filename, args.table_index,
                args.width, width_type=kwargs.get("width_type", "points"),
            )
        elif cmd == "auto_fit_table_columns":
            out = auto_fit_table_columns(args.filename, args.table_index)
        elif cmd == "set_paragraph_alignment":
            out = set_paragraph_alignment(args.filename, args.paragraph_index, args.alignment)
        elif cmd == "set_paragraph_indent":
            out = set_paragraph_indent(
                args.filename, args.paragraph_index,
                left=kwargs.get("left"),
                right=kwargs.get("right"),
                first_line=kwargs.get("first_line"),
                hanging=kwargs.get("hanging"),
            )
        elif cmd == "set_paragraph_spacing":
            out = set_paragraph_spacing(
                args.filename, args.paragraph_index,
                before=kwargs.get("before"),
                after=kwargs.get("after"),
                line_spacing=kwargs.get("line_spacing"),
                line_spacing_rule=kwargs.get("line_spacing_rule"),
            )
        elif cmd == "format_all_paragraphs":
            out = format_all_paragraphs(
                args.filename,
                alignment=kwargs.get("alignment"),
                left_indent=kwargs.get("left_indent"),
                first_line_indent=kwargs.get("first_line_indent"),
                space_before=kwargs.get("space_before"),
                space_after=kwargs.get("space_after"),
                line_spacing=kwargs.get("line_spacing"),
            )
        elif cmd == "get_all_comments":
            out = get_all_comments(args.filename)
        elif cmd == "get_comments_by_author":
            out = get_comments_by_author(args.filename, args.author)
        elif cmd == "get_comments_for_paragraph":
            out = get_comments_for_paragraph(args.filename, args.paragraph_index)
        elif cmd == "add_comment":
            out = add_comment(args.filename, args.paragraph_index, args.comment_text, author=kwargs.get("author", "AI Assistant"))
        elif cmd == "delete_comment":
            out = delete_comment(args.filename, args.comment_id)
        elif cmd == "add_password_protection":
            out = add_password_protection(
                args.filename, args.password,
                output_filename=kwargs.get("output_filename"),
            )
        elif cmd == "add_restricted_editing":
            out = add_restricted_editing(
                args.filename,
                password=kwargs.get("password"),
                output_filename=kwargs.get("output_filename"),
            )
        else:
            out = f"Unknown command: {cmd}"
        print(out)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
