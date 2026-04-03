#!/usr/bin/env python3
"""CLI for PowerPoint operations (deskclaw-pptx skill)."""

import argparse
import json
import sys
from pathlib import Path

_SCRIPT_DIR = Path(__file__).resolve().parent
if str(_SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPT_DIR))

from core import (
    create_presentation,
    create_presentation_from_template,
    get_presentation_info,
    set_core_properties,
    get_template_file_info,
    add_slide,
    get_slide_info,
    extract_slide_text,
    extract_presentation_text,
    populate_placeholder,
    add_bullet_points,
    add_table,
    add_shape,
    format_table_cell,
    add_chart,
)


def _bool_arg(s):
    return s.lower() in ("true", "1", "yes")


def main():
    parser = argparse.ArgumentParser(description="PowerPoint CLI (deskclaw-pptx)")
    sub = parser.add_subparsers(dest="command", required=True)

    # presentation
    p = sub.add_parser("create_presentation")
    p.add_argument("file_path")

    p = sub.add_parser("create_presentation_from_template")
    p.add_argument("template_path")
    p.add_argument("file_path")

    p = sub.add_parser("get_presentation_info")
    p.add_argument("file_path")

    p = sub.add_parser("set_core_properties")
    p.add_argument("file_path")
    p.add_argument("--title", default=None)
    p.add_argument("--subject", default=None)
    p.add_argument("--author", default=None)
    p.add_argument("--keywords", default=None)
    p.add_argument("--comments", default=None)

    p = sub.add_parser("get_template_file_info")
    p.add_argument("template_path")

    # content
    p = sub.add_parser("add_slide")
    p.add_argument("file_path")
    p.add_argument("--layout-index", type=int, default=0)
    p.add_argument("--title", default=None)

    p = sub.add_parser("get_slide_info")
    p.add_argument("file_path")
    p.add_argument("--slide-index", type=int, required=True)

    p = sub.add_parser("extract_slide_text")
    p.add_argument("file_path")
    p.add_argument("--slide-index", type=int, required=True)

    p = sub.add_parser("extract_presentation_text")
    p.add_argument("file_path")
    p.add_argument("--no-slide-info", action="store_true", help="omit slide headers in combined text")

    p = sub.add_parser("populate_placeholder")
    p.add_argument("file_path")
    p.add_argument("--slide-index", type=int, required=True)
    p.add_argument("--placeholder-idx", type=int, required=True)
    p.add_argument("--text", required=True)

    p = sub.add_parser("add_bullet_points")
    p.add_argument("file_path")
    p.add_argument("--slide-index", type=int, required=True)
    p.add_argument("--items", required=True, type=json.loads)
    p.add_argument("--left", type=float, default=1.0)
    p.add_argument("--top", type=float, default=1.5)
    p.add_argument("--width", type=float, default=8.0)
    p.add_argument("--height", type=float, default=4.0)

    # structural
    p = sub.add_parser("add_table")
    p.add_argument("file_path")
    p.add_argument("--slide-index", type=int, required=True)
    p.add_argument("--rows", type=int, required=True)
    p.add_argument("--cols", type=int, required=True)
    p.add_argument("--data", default=None, type=lambda x: json.loads(x) if x else None)
    p.add_argument("--left", type=float, default=1.0)
    p.add_argument("--top", type=float, default=2.0)
    p.add_argument("--width", type=float, default=8.0)
    p.add_argument("--height", type=float, default=3.0)

    p = sub.add_parser("add_shape")
    p.add_argument("file_path")
    p.add_argument("--slide-index", type=int, required=True)
    p.add_argument("--shape-type", required=True)
    p.add_argument("--left", type=float, default=1.0)
    p.add_argument("--top", type=float, default=1.0)
    p.add_argument("--width", type=float, default=1.0)
    p.add_argument("--height", type=float, default=1.0)
    p.add_argument("--text", default=None)

    p = sub.add_parser("format_table_cell")
    p.add_argument("file_path")
    p.add_argument("--slide-index", type=int, required=True)
    p.add_argument("--shape-index", type=int, required=True)
    p.add_argument("--row-index", type=int, required=True)
    p.add_argument("--col-index", type=int, required=True)
    p.add_argument("--text", default=None)
    p.add_argument("--font-size", type=int, default=None)
    p.add_argument("--bold", type=_bool_arg, default=None)

    p = sub.add_parser("add_chart")
    p.add_argument("file_path")
    p.add_argument("--slide-index", type=int, required=True)
    p.add_argument("--chart-type", required=True, choices=["column", "bar", "line", "pie"])
    p.add_argument("--categories", required=True, type=json.loads)
    p.add_argument("--series-names", required=True, type=json.loads)
    p.add_argument("--series-values", required=True, type=json.loads)
    p.add_argument("--left", type=float, default=1.0)
    p.add_argument("--top", type=float, default=2.0)
    p.add_argument("--width", type=float, default=8.0)
    p.add_argument("--height", type=float, default=4.0)
    p.add_argument("--title", default=None)

    args = parser.parse_args()
    cmd = args.command

    try:
        if cmd == "create_presentation":
            out = create_presentation(args.file_path)
        elif cmd == "create_presentation_from_template":
            out = create_presentation_from_template(args.template_path, args.file_path)
        elif cmd == "get_presentation_info":
            out = get_presentation_info(args.file_path)
        elif cmd == "set_core_properties":
            out = set_core_properties(
                args.file_path,
                title=args.title,
                subject=args.subject,
                author=args.author,
                keywords=args.keywords,
                comments=args.comments,
            )
        elif cmd == "get_template_file_info":
            out = get_template_file_info(args.template_path)
        elif cmd == "add_slide":
            out = add_slide(args.file_path, layout_index=args.layout_index, title=args.title)
        elif cmd == "get_slide_info":
            out = get_slide_info(args.file_path, args.slide_index)
        elif cmd == "extract_slide_text":
            out = extract_slide_text(args.file_path, args.slide_index)
        elif cmd == "extract_presentation_text":
            out = extract_presentation_text(args.file_path, include_slide_info=not args.no_slide_info)
        elif cmd == "populate_placeholder":
            out = populate_placeholder(args.file_path, args.slide_index, args.placeholder_idx, args.text)
        elif cmd == "add_bullet_points":
            out = add_bullet_points(
                args.file_path,
                args.slide_index,
                args.items,
                left=args.left,
                top=args.top,
                width=args.width,
                height=args.height,
            )
        elif cmd == "add_table":
            out = add_table(
                args.file_path,
                args.slide_index,
                args.rows,
                args.cols,
                data=args.data,
                left=args.left,
                top=args.top,
                width=args.width,
                height=args.height,
            )
        elif cmd == "add_shape":
            out = add_shape(
                args.file_path,
                args.slide_index,
                args.shape_type,
                left=args.left,
                top=args.top,
                width=args.width,
                height=args.height,
                text=args.text,
            )
        elif cmd == "format_table_cell":
            out = format_table_cell(
                args.file_path,
                args.slide_index,
                args.shape_index,
                args.row_index,
                args.col_index,
                text=args.text,
                font_size=args.font_size,
                bold=args.bold,
            )
        elif cmd == "add_chart":
            out = add_chart(
                args.file_path,
                args.slide_index,
                args.chart_type,
                args.categories,
                args.series_names,
                args.series_values,
                left=args.left,
                top=args.top,
                width=args.width,
                height=args.height,
                title=args.title,
            )
        else:
            out = f"Unknown command: {cmd}"
        print(out)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
