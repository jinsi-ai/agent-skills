---
name: xlsx
slug: deskclaw-xlsx
version: 1.0.2
description: "Use this skill any time a spreadsheet file is the primary input or output. This means any task where the user wants to: open, read, edit, or fix an existing .xlsx, .xlsm, .csv, or .tsv file (e.g., adding columns, computing formulas, formatting, charting, cleaning messy data); create a new spreadsheet from scratch or from other data sources; or convert between tabular file formats. Trigger especially when the user references a spreadsheet file by name or path. The deliverable must be a spreadsheet file."
---

# Excel 表格操作 (deskclaw-xlsx)

本技能通过 `scripts/excel_tool.py` 提供 .xlsx 的创建、读写、公式、格式化、图表、透视汇总和工作表管理能力。基于 [excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) 的 openpyxl 能力模型，采用无状态 CLI（每次命令对给定文件 load -> 操作 -> save）。

## 输出规则（必读）

- **只保存到本地**：所有生成或编辑的文件保存到用户指定路径或当前工作目录，操作完成后告知用户完整的文件路径。
- **禁止向对话发送文件内容**：不要将文件内容（文本、二进制、base64 等任何形式）粘贴或发送到用户的聊天/会话中。大文件会导致整个 session 卡死。
- **如需预览**：只在对话中展示简短的摘要信息（如文档标题、页数、表格行列数等元信息），不要展示完整内容。

## 首次使用（前置条件）

1. 安装 `openpyxl`（必须）

```bash
pip install openpyxl
```

2. 调用方式
- 若当前目录是技能根目录：`python scripts/excel_tool.py <command> ...`
- 否则使用绝对路径：`python /path/to/skills/deskclaw-xlsx/scripts/excel_tool.py <command> ...`

## 命令列表（24）

### workbook
- `create_workbook file_path`
- `create_worksheet file_path sheet_name`
- `get_workbook_metadata file_path [--include-ranges]`

### data
- `read_data file_path sheet_name [--start-cell A1] [--end-cell B10] [--preview-only]`
- `write_data file_path sheet_name --data '[[1,2],[3,4]]' [--start-cell A1]`

### formula
- `apply_formula file_path sheet_name cell formula`
- `validate_formula file_path sheet_name cell formula`

### formatting
- `format_range file_path sheet_name start_cell [--end-cell B10] [--bold true] [--italic true] [--underline true] [--font-size 12] [--font-color FF0000] [--bg-color FFFF00] [--border-style thin] [--border-color 000000] [--number-format "0.00%"] [--alignment center] [--wrap-text true] [--merge-cells true]`
- `validate_range file_path sheet_name start_cell [--end-cell B10]`

### sheet
- `copy_worksheet file_path source_sheet target_sheet`
- `delete_worksheet file_path sheet_name`
- `rename_worksheet file_path old_name new_name`
- `merge_cells file_path sheet_name start_cell end_cell`
- `unmerge_cells file_path sheet_name start_cell end_cell`
- `get_merged_cells file_path sheet_name`
- `copy_range file_path sheet_name source_start source_end target_start [--target-sheet Sheet2]`
- `delete_range file_path sheet_name start_cell end_cell [--shift-direction up|left]`
- `insert_rows file_path sheet_name start_row [--count 1]`
- `insert_columns file_path sheet_name start_col [--count 1]`
- `delete_rows file_path sheet_name start_row [--count 1]`
- `delete_columns file_path sheet_name start_col [--count 1]`

### chart
- `create_chart file_path sheet_name data_range chart_type target_cell [--title ...] [--x-axis ...] [--y-axis ...]`
  - `chart_type`: `line` / `bar` / `column` / `pie` / `scatter`

### table
- `create_table file_path sheet_name data_range [--table-name T1] [--table-style TableStyleMedium9]`
- `create_pivot_table file_path sheet_name data_range --rows '["Region"]' --values '["Sales"]' [--columns '["Year"]'] [--agg-func mean|sum|count|max|min]`

## 示例工作流

```bash
# 1) 新建工作簿
python scripts/excel_tool.py create_workbook /tmp/demo.xlsx

# 2) 新建工作表并写入数据
python scripts/excel_tool.py create_worksheet /tmp/demo.xlsx Sales
python scripts/excel_tool.py write_data /tmp/demo.xlsx Sales --data '[["Region","Year","Sales"],["East",2024,100],["West",2024,120],["East",2025,150]]' --start-cell A1

# 3) 公式与格式化
python scripts/excel_tool.py apply_formula /tmp/demo.xlsx Sales D2 '=C2*1.1'
python scripts/excel_tool.py format_range /tmp/demo.xlsx Sales A1:D1 --bold true --bg-color D9E1F2

# 4) 创建图表与表格
python scripts/excel_tool.py create_chart /tmp/demo.xlsx Sales A1:C4 column F2 --title "Sales Trend"
python scripts/excel_tool.py create_table /tmp/demo.xlsx Sales A1:D4 --table-name SalesTable

# 5) 透视汇总与读取验证
python scripts/excel_tool.py create_pivot_table /tmp/demo.xlsx Sales A1:C4 --rows '["Region"]' --values '["Sales"]' --agg-func sum
python scripts/excel_tool.py read_data /tmp/demo.xlsx Sales --start-cell A1 --end-cell D4
python scripts/excel_tool.py get_workbook_metadata /tmp/demo.xlsx --include-ranges
```

## 注意事项

- 单元格地址使用 Excel 记法：`A1`、`B2:D10`
- `--data`、`--rows`、`--values`、`--columns` 参数均需合法 JSON 字符串
- 颜色建议使用十六进制（如 `FF0000`）
- `file_path` 建议使用绝对路径，避免工作目录歧义
- `create_pivot_table` 在本技能中生成的是“透视汇总表”工作表（便于无状态 CLI 使用）
