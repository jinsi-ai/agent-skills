#!/usr/bin/env python3
"""CLI for Excel operations (deskclaw-xlsx skill)."""

import argparse
import json
import sys
from pathlib import Path

_SCRIPT_DIR = Path(__file__).resolve().parent
if str(_SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPT_DIR))

from core import (
    create_workbook,
    create_worksheet,
    get_workbook_metadata,
    read_data,
    write_data,
    apply_formula,
    validate_formula,
    format_range,
    validate_range,
    copy_worksheet,
    delete_worksheet,
    rename_worksheet,
    merge_cells,
    unmerge_cells,
    get_merged_cells,
    copy_range,
    delete_range,
    insert_rows,
    insert_columns,
    delete_rows,
    delete_columns,
    create_chart,
    create_table,
    create_pivot_table,
)


def _bool_arg(s):
    return str(s).lower() in ("true", "1", "yes")


def main():
    parser = argparse.ArgumentParser(description="Excel CLI (deskclaw-xlsx)")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("create_workbook")
    p.add_argument("file_path")

    p = sub.add_parser("create_worksheet")
    p.add_argument("file_path")
    p.add_argument("sheet_name")

    p = sub.add_parser("get_workbook_metadata")
    p.add_argument("file_path")
    p.add_argument("--include-ranges", action="store_true")

    p = sub.add_parser("read_data")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("--start-cell", default="A1")
    p.add_argument("--end-cell", default=None)
    p.add_argument("--preview-only", action="store_true")

    p = sub.add_parser("write_data")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("--data", required=True, type=json.loads)
    p.add_argument("--start-cell", default="A1")

    p = sub.add_parser("apply_formula")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("cell")
    p.add_argument("formula")

    p = sub.add_parser("validate_formula")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("cell")
    p.add_argument("formula")

    p = sub.add_parser("format_range")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("start_cell")
    p.add_argument("--end-cell", default=None)
    p.add_argument("--bold", type=_bool_arg, default=False)
    p.add_argument("--italic", type=_bool_arg, default=False)
    p.add_argument("--underline", type=_bool_arg, default=False)
    p.add_argument("--font-size", type=int, default=None)
    p.add_argument("--font-color", default=None)
    p.add_argument("--bg-color", default=None)
    p.add_argument("--border-style", default=None)
    p.add_argument("--border-color", default=None)
    p.add_argument("--number-format", default=None)
    p.add_argument("--alignment", default=None)
    p.add_argument("--wrap-text", type=_bool_arg, default=False)
    p.add_argument("--merge-cells", type=_bool_arg, default=False)
    p.add_argument("--protection", type=json.loads, default=None)
    p.add_argument("--conditional-format", type=json.loads, default=None)

    p = sub.add_parser("validate_range")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("start_cell")
    p.add_argument("--end-cell", default=None)

    p = sub.add_parser("copy_worksheet")
    p.add_argument("file_path")
    p.add_argument("source_sheet")
    p.add_argument("target_sheet")

    p = sub.add_parser("delete_worksheet")
    p.add_argument("file_path")
    p.add_argument("sheet_name")

    p = sub.add_parser("rename_worksheet")
    p.add_argument("file_path")
    p.add_argument("old_name")
    p.add_argument("new_name")

    p = sub.add_parser("merge_cells")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("start_cell")
    p.add_argument("end_cell")

    p = sub.add_parser("unmerge_cells")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("start_cell")
    p.add_argument("end_cell")

    p = sub.add_parser("get_merged_cells")
    p.add_argument("file_path")
    p.add_argument("sheet_name")

    p = sub.add_parser("copy_range")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("source_start")
    p.add_argument("source_end")
    p.add_argument("target_start")
    p.add_argument("--target-sheet", default=None)

    p = sub.add_parser("delete_range")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("start_cell")
    p.add_argument("end_cell")
    p.add_argument("--shift-direction", default="up")

    p = sub.add_parser("insert_rows")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("start_row", type=int)
    p.add_argument("--count", type=int, default=1)

    p = sub.add_parser("insert_columns")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("start_col", type=int)
    p.add_argument("--count", type=int, default=1)

    p = sub.add_parser("delete_rows")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("start_row", type=int)
    p.add_argument("--count", type=int, default=1)

    p = sub.add_parser("delete_columns")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("start_col", type=int)
    p.add_argument("--count", type=int, default=1)

    p = sub.add_parser("create_chart")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("data_range")
    p.add_argument("chart_type")
    p.add_argument("target_cell")
    p.add_argument("--title", default="")
    p.add_argument("--x-axis", default="")
    p.add_argument("--y-axis", default="")

    p = sub.add_parser("create_table")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("data_range")
    p.add_argument("--table-name", default=None)
    p.add_argument("--table-style", default="TableStyleMedium9")

    p = sub.add_parser("create_pivot_table")
    p.add_argument("file_path")
    p.add_argument("sheet_name")
    p.add_argument("data_range")
    p.add_argument("--rows", required=True, type=json.loads)
    p.add_argument("--values", required=True, type=json.loads)
    p.add_argument("--columns", default=None, type=json.loads)
    p.add_argument("--agg-func", default="mean")

    args = parser.parse_args()
    cmd = args.command

    try:
        if cmd == "create_workbook":
            out = create_workbook(args.file_path)
        elif cmd == "create_worksheet":
            out = create_worksheet(args.file_path, args.sheet_name)
        elif cmd == "get_workbook_metadata":
            out = get_workbook_metadata(args.file_path, include_ranges=args.include_ranges)
        elif cmd == "read_data":
            out = read_data(args.file_path, args.sheet_name, start_cell=args.start_cell, end_cell=args.end_cell, preview_only=args.preview_only)
        elif cmd == "write_data":
            out = write_data(args.file_path, args.sheet_name, args.data, start_cell=args.start_cell)
        elif cmd == "apply_formula":
            out = apply_formula(args.file_path, args.sheet_name, args.cell, args.formula)
        elif cmd == "validate_formula":
            out = validate_formula(args.file_path, args.sheet_name, args.cell, args.formula)
        elif cmd == "format_range":
            out = format_range(args.file_path, args.sheet_name, args.start_cell, end_cell=args.end_cell, bold=args.bold, italic=args.italic, underline=args.underline, font_size=args.font_size, font_color=args.font_color, bg_color=args.bg_color, border_style=args.border_style, border_color=args.border_color, number_format=args.number_format, alignment=args.alignment, wrap_text=args.wrap_text, merge_cells=args.merge_cells, protection=args.protection, conditional_format=args.conditional_format)
        elif cmd == "validate_range":
            out = validate_range(args.file_path, args.sheet_name, args.start_cell, end_cell=args.end_cell)
        elif cmd == "copy_worksheet":
            out = copy_worksheet(args.file_path, args.source_sheet, args.target_sheet)
        elif cmd == "delete_worksheet":
            out = delete_worksheet(args.file_path, args.sheet_name)
        elif cmd == "rename_worksheet":
            out = rename_worksheet(args.file_path, args.old_name, args.new_name)
        elif cmd == "merge_cells":
            out = merge_cells(args.file_path, args.sheet_name, args.start_cell, args.end_cell)
        elif cmd == "unmerge_cells":
            out = unmerge_cells(args.file_path, args.sheet_name, args.start_cell, args.end_cell)
        elif cmd == "get_merged_cells":
            out = get_merged_cells(args.file_path, args.sheet_name)
        elif cmd == "copy_range":
            out = copy_range(args.file_path, args.sheet_name, args.source_start, args.source_end, args.target_start, target_sheet=args.target_sheet)
        elif cmd == "delete_range":
            out = delete_range(args.file_path, args.sheet_name, args.start_cell, args.end_cell, shift_direction=args.shift_direction)
        elif cmd == "insert_rows":
            out = insert_rows(args.file_path, args.sheet_name, args.start_row, count=args.count)
        elif cmd == "insert_columns":
            out = insert_columns(args.file_path, args.sheet_name, args.start_col, count=args.count)
        elif cmd == "delete_rows":
            out = delete_rows(args.file_path, args.sheet_name, args.start_row, count=args.count)
        elif cmd == "delete_columns":
            out = delete_columns(args.file_path, args.sheet_name, args.start_col, count=args.count)
        elif cmd == "create_chart":
            out = create_chart(args.file_path, args.sheet_name, args.data_range, args.chart_type, args.target_cell, title=args.title, x_axis=args.x_axis, y_axis=args.y_axis)
        elif cmd == "create_table":
            out = create_table(args.file_path, args.sheet_name, args.data_range, table_name=args.table_name, table_style=args.table_style)
        elif cmd == "create_pivot_table":
            out = create_pivot_table(args.file_path, args.sheet_name, args.data_range, args.rows, args.values, columns=args.columns, agg_func=args.agg_func)
        else:
            out = f"Unknown command: {cmd}"
        print(out)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
