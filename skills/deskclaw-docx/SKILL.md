---
name: docx
slug: deskclaw-docx
version: 1.0.2
description: "Use this skill whenever the user wants to create, read, edit, or manipulate Word documents (.docx files). Triggers include: any mention of 'Word doc', 'word document', '.docx', or requests to produce professional documents with formatting like tables of contents, headings, page numbers, or letterheads. Also use when extracting or reorganizing content from .docx files, inserting or replacing images in documents, performing find-and-replace in Word files, working with tracked changes or comments, or converting content into a polished Word document. If the user asks for a 'report', 'memo', 'letter', 'template', or similar deliverable as a Word or .docx file, use this skill. Do NOT use for PDFs, spreadsheets, Google Docs, or general coding tasks unrelated to document generation."
---

# Word 文档操作 (deskclaw-docx)

本技能通过 `scripts/word_tool.py` 提供 Word 文档的创建、读取、编辑和格式化能力。基于 python-docx，无需手动编写 JavaScript 或操作 XML。

## 输出规则（必读）

- **只保存到本地**：所有生成或编辑的文件保存到用户指定路径或当前工作目录，操作完成后告知用户完整的文件路径。
- **禁止向对话发送文件内容**：不要将文件内容（文本、二进制、base64 等任何形式）粘贴或发送到用户的聊天/会话中。大文件会导致整个 session 卡死。
- **如需预览**：只在对话中展示简短的摘要信息（如文档标题、页数、表格行列数等元信息），不要展示完整内容。

## 首次使用（前置条件）

**不能“拿到就用”**，需先满足以下条件：

1. **安装 python-docx**（必须）：  
   ```bash
   pip install python-docx
   ```  
   未安装时运行脚本会报错并提示上述命令。

2. **调用脚本的方式**：  
   - 若当前工作目录是技能根目录 `.../skills/deskclaw-docx/`，可直接写：  
     `python scripts/word_tool.py <命令> ...`  
   - 否则请使用脚本的**绝对路径**，例如：  
     `python /path/to/openclaw/skills/deskclaw-docx/scripts/word_tool.py <命令> ...`  
   agent 执行时若不在技能目录下，应使用该技能所在路径下的 `scripts/word_tool.py` 的绝对路径。

## 依赖

- **python-docx**：`pip install python-docx`（必须）
- **docx2pdf**（可选）：仅当需要转 PDF 时 `pip install docx2pdf`
- **msoffcrypto-tool**（可选）：仅当需要密码保护时 `pip install msoffcrypto-tool`

## 工具概览

| 类别 | 命令 | 用途 |
|------|------|------|
| 文档管理 | `create_document` | 创建新 .docx |
| 文档管理 | `get_document_info` | 文档属性和统计 |
| 文档管理 | `get_document_text` | 提取全文 |
| 文档管理 | `get_document_outline` | 文档大纲 |
| 文档管理 | `list_available_documents` | 列出目录下 .docx |
| 文档管理 | `copy_document` | 复制文档 |
| 文档管理 | `convert_to_pdf` | 转为 PDF |
| 文档管理 | `merge_documents` | 合并多个文档 |
| 页面设置 | `set_page_size` | 纸张大小（A4/A3/Letter等） |
| 页面设置 | `set_page_margins` | 页边距 |
| 页面设置 | `set_page_orientation` | 横向/纵向 |
| 页面设置 | `get_page_settings` | 获取当前页面设置 |
| 页眉页脚 | `add_header` | 添加页眉文本 |
| 页眉页脚 | `add_footer` | 添加页脚文本 |
| 页眉页脚 | `add_page_number` | 添加页码 |
| 页眉页脚 | `set_different_first_page_header` | 首页不同页眉 |
| 页眉页脚 | `add_header_image` | 页眉添加图片/Logo |
| 水印 | `add_watermark` | 添加文字水印 |
| 水印 | `add_watermark_image` | 添加图片水印 |
| 目录 | `add_table_of_contents` | 添加目录 |
| 目录 | `update_table_of_contents` | 标记目录更新 |
| 脚注尾注 | `add_footnote` | 添加脚注 |
| 脚注尾注 | `add_endnote` | 添加尾注 |
| 内容创建 | `add_heading` | 添加标题 |
| 内容创建 | `add_paragraph` | 添加段落 |
| 内容创建 | `add_table` | 添加表格 |
| 内容创建 | `add_picture` | 添加图片 |
| 内容创建 | `add_page_break` | 分页符 |
| 高级内容 | `insert_header_near_text` | 在指定文本附近插入标题 |
| 高级内容 | `insert_line_or_paragraph_near_text` | 在指定文本附近插入段落 |
| 高级内容 | `insert_numbered_list_near_text` | 在指定文本附近插入列表 |
| 超链接 | `add_hyperlink` | 添加超链接 |
| 超链接 | `add_hyperlink_to_text` | 将文本转为超链接 |
| 超链接 | `get_hyperlinks` | 获取所有超链接 |
| 段落格式 | `set_paragraph_alignment` | 段落对齐 |
| 段落格式 | `set_paragraph_indent` | 段落缩进 |
| 段落格式 | `set_paragraph_spacing` | 段落间距 |
| 段落格式 | `format_all_paragraphs` | 批量格式化所有段落 |
| 文本格式 | `format_text` | 格式化指定文本片段 |
| 文本格式 | `search_and_replace` | 全文查找替换 |
| 文本格式 | `delete_paragraph` | 删除段落 |
| 文本格式 | `create_custom_style` | 创建自定义样式 |
| 表格内容 | `get_table_info` | 获取表格结构和内容 |
| 表格内容 | `set_table_cell` | 设置单个单元格内容 |
| 表格内容 | `batch_set_table_cells` | 批量设置多个单元格 |
| 表格行列 | `add_table_row` | 添加行 |
| 表格行列 | `add_table_rows` | 批量添加行 |
| 表格行列 | `delete_table_row` | 删除行 |
| 表格行列 | `delete_table_rows` | 批量删除行 |
| 表格行列 | `add_table_column` | 添加列 |
| 表格行列 | `delete_table_column` | 删除列 |
| 表格格式 | `format_table` | 表格边框与样式 |
| 表格格式 | `set_table_cell_shading` | 单元格底色 |
| 表格格式 | `apply_table_alternating_rows` | 交替行颜色 |
| 表格格式 | `highlight_table_header` | 高亮表头 |
| 表格格式 | `format_table_cell_text` | 单元格文本格式 |
| 表格格式 | `set_table_cell_padding` | 单元格内边距 |
| 表格格式 | `set_table_cell_alignment` | 单元格对齐 |
| 表格格式 | `set_table_alignment_all` | 整表对齐 |
| 表格格式 | `merge_table_cells` | 合并单元格（矩形） |
| 表格格式 | `merge_table_cells_horizontal` | 水平合并 |
| 表格格式 | `merge_table_cells_vertical` | 垂直合并 |
| 表格格式 | `set_table_column_width` | 单列宽 |
| 表格格式 | `set_table_column_widths` | 各列宽 |
| 表格格式 | `set_table_width` | 表格总宽 |
| 表格格式 | `auto_fit_table_columns` | 自动适应列宽 |
| 批注 | `get_all_comments` | 提取所有批注 |
| 批注 | `get_comments_by_author` | 按作者筛选批注 |
| 批注 | `get_comments_for_paragraph` | 某段落的批注 |
| 批注 | `add_comment` | 添加批注 |
| 批注 | `delete_comment` | 删除批注 |
| 保护 | `add_password_protection` | 密码保护 |
| 保护 | `add_restricted_editing` | 限制编辑 |

## 调用方式

在技能根目录（即 `deskclaw-docx/` 所在目录）下执行，或使用脚本绝对路径。`filename` 使用绝对路径或相对当前工作目录的路径。

### 创建文档

```bash
python scripts/word_tool.py create_document report.docx --title "年度报告" --author "张三"
```

### 页面设置

```bash
# 设置纸张大小（支持 A4, A3, Letter, Legal, B5, custom）
python scripts/word_tool.py set_page_size report.docx --paper A4

# 自定义纸张大小（单位：英寸）
python scripts/word_tool.py set_page_size report.docx --paper custom --width 8.5 --height 11

# 设置页边距（单位：英寸）
python scripts/word_tool.py set_page_margins report.docx --top 1 --bottom 1 --left 1.25 --right 1.25

# 设置横向/纵向
python scripts/word_tool.py set_page_orientation report.docx --orientation landscape
python scripts/word_tool.py set_page_orientation report.docx --orientation portrait

# 查看当前页面设置
python scripts/word_tool.py get_page_settings report.docx
```

### 页眉页脚与目录

```bash
# 添加页眉
python scripts/word_tool.py add_header report.docx "公司名称" --alignment center --bold true

# 添加页脚
python scripts/word_tool.py add_footer report.docx "机密文件" --alignment left

# 添加页码（默认在页脚居中）
python scripts/word_tool.py add_page_number report.docx --position footer --alignment center --format "第 {page} 页"

# 页眉添加 Logo
python scripts/word_tool.py add_header_image report.docx logo.png --width 1.5 --alignment left

# 首页不同页眉
python scripts/word_tool.py set_different_first_page_header report.docx --enabled true

# 添加目录
python scripts/word_tool.py add_table_of_contents report.docx --title "目录" --max-level 3 --position start
```

### 水印

```bash
# 添加文字水印（如"机密"、"草稿"）
python scripts/word_tool.py add_watermark report.docx "机密" --font-size 72 --color C0C0C0

# 添加图片水印
python scripts/word_tool.py add_watermark_image report.docx watermark.png --width 3
```

### 脚注与尾注

```bash
# 在指定段落添加脚注（显示在页面底部）
python scripts/word_tool.py add_footnote report.docx --paragraph-index 2 --text "数据来源：国家统计局"

# 在指定段落添加尾注（显示在文档末尾）
python scripts/word_tool.py add_endnote report.docx --paragraph-index 5 --text "详见附录A"
```

### 添加内容

```bash
python scripts/word_tool.py add_heading report.docx "第一章 概述" --level 1 --font-name Arial --bold
python scripts/word_tool.py add_paragraph report.docx "正文内容。" --font-name "宋体" --font-size 24 --color 333333
python scripts/word_tool.py add_table report.docx --rows 4 --cols 3 --data '[["姓名","部门","职位"],["张三","技术部","工程师"]]'
python scripts/word_tool.py add_picture report.docx image.png --width 2
python scripts/word_tool.py add_page_break report.docx
```

### 在指定位置插入（近文本）

```bash
python scripts/word_tool.py insert_header_near_text report.docx --target-text "概述" --header-title "详细说明" --position after
python scripts/word_tool.py insert_line_or_paragraph_near_text report.docx --target-text "步骤如下" --line-text "第一步内容" --position after
python scripts/word_tool.py insert_numbered_list_near_text report.docx --list-items '["第一步","第二步","第三步"]' --bullet-type number
```

### 超链接

**重要**：超链接必须使用 `add_hyperlink` 命令创建，不能用 `add_paragraph` 写纯文本！

```bash
# 添加新超链接（创建新段落，链接文字会显示为蓝色带下划线）
python scripts/word_tool.py add_hyperlink report.docx "点击访问OpenAI" "https://openai.com"

# 在现有段落末尾追加超链接（先写前缀文字，再追加链接）
python scripts/word_tool.py add_paragraph report.docx "了解更多请访问："
python scripts/word_tool.py add_hyperlink report.docx "OpenAI官网" "https://openai.com" --paragraph-index 5

# 将文档中已有的文本转为超链接
python scripts/word_tool.py add_hyperlink_to_text report.docx --find "公司官网" --url "https://example.com"

# 获取所有超链接
python scripts/word_tool.py get_hyperlinks report.docx
```

**错误示范**（不会生成可点击链接）：
```bash
# ❌ 错误：这只是纯文本，不是真正的超链接
python scripts/word_tool.py add_paragraph report.docx "访问 https://openai.com"
```

**正确示范**：
```bash
# ✅ 正确：使用 add_hyperlink 创建可点击的链接
python scripts/word_tool.py add_hyperlink report.docx "访问OpenAI" "https://openai.com"
```

### 段落格式化

```bash
# 设置段落对齐
python scripts/word_tool.py set_paragraph_alignment report.docx --paragraph-index 0 --alignment center

# 设置段落缩进（单位：磅）
python scripts/word_tool.py set_paragraph_indent report.docx --paragraph-index 0 --left 36 --first-line 24

# 设置段落间距
python scripts/word_tool.py set_paragraph_spacing report.docx --paragraph-index 0 --before 12 --after 12 --line-spacing 1.5

# 批量格式化所有段落
python scripts/word_tool.py format_all_paragraphs report.docx --alignment justify --first-line-indent 24 --line-spacing 1.5
```

### 文本格式化

```bash
python scripts/word_tool.py format_text report.docx --paragraph-index 2 --start-pos 0 --end-pos 5 --bold true --color FF0000 --font-size 28
python scripts/word_tool.py search_and_replace report.docx --find "旧内容" --replace "新内容"
python scripts/word_tool.py delete_paragraph report.docx --paragraph-index 1
python scripts/word_tool.py create_custom_style report.docx --style-name "MyHeading" --bold true --font-size 28
```

### 表格内容操作

```bash
# 查看文档中所有表格的概要
python scripts/word_tool.py get_table_info report.docx

# 查看指定表格的详细内容（包含合并单元格信息）
python scripts/word_tool.py get_table_info report.docx --table-index 0 --show-merged true

# 设置单个单元格内容
python scripts/word_tool.py set_table_cell report.docx --table-index 0 --row 1 --col 2 --text "新内容"

# 使用视觉列索引（处理合并单元格时更直观）
python scripts/word_tool.py set_table_cell report.docx --table-index 0 --row 1 --col 2 --text "新内容" --visual true

# 批量设置多个单元格
python scripts/word_tool.py batch_set_table_cells report.docx --table-index 0 --cells '[{"row":0,"col":1,"text":"企业名称"},{"row":1,"col":0,"text":"张三"}]'

# 批量设置（使用视觉索引）
python scripts/word_tool.py batch_set_table_cells report.docx --table-index 0 --cells '[...]' --visual true
```

### 表格行列操作

```bash
# 添加行（末尾）
python scripts/word_tool.py add_table_row report.docx --table-index 0

# 添加行（开头或指定位置）
python scripts/word_tool.py add_table_row report.docx --table-index 0 --position start
python scripts/word_tool.py add_table_row report.docx --table-index 0 --position 2

# 批量添加行
python scripts/word_tool.py add_table_rows report.docx --table-index 0 --count 5 --position end

# 删除行
python scripts/word_tool.py delete_table_row report.docx --table-index 0 --row 3

# 批量删除行
python scripts/word_tool.py delete_table_rows report.docx --table-index 0 --rows '[3,5,7]'

# 添加列
python scripts/word_tool.py add_table_column report.docx --table-index 0 --position end --width 1.5

# 删除列
python scripts/word_tool.py delete_table_column report.docx --table-index 0 --col 2
```

### 表格格式化

```bash
python scripts/word_tool.py format_table report.docx --table-index 0 --has-header-row true
python scripts/word_tool.py highlight_table_header report.docx --table-index 0 --header-color 4472C4 --text-color FFFFFF
python scripts/word_tool.py apply_table_alternating_rows report.docx --table-index 0 --color1 FFFFFF --color2 F2F2F2
python scripts/word_tool.py set_table_cell_shading report.docx --table-index 0 --row-index 0 --col-index 0 --fill-color D5E8F0
python scripts/word_tool.py format_table_cell_text report.docx --table-index 0 --row-index 0 --col-index 0 --bold true --color FFFFFF
python scripts/word_tool.py set_table_cell_padding report.docx --table-index 0 --row-index 0 --col-index 0 --top 10 --bottom 10 --left 10 --right 10
python scripts/word_tool.py merge_table_cells report.docx --table-index 0 --start-row 0 --start-col 0 --end-row 0 --end-col 2
python scripts/word_tool.py set_table_column_widths report.docx --table-index 0 --widths '[200,150,150]' --width-type points
python scripts/word_tool.py auto_fit_table_columns report.docx --table-index 0
```

### 读取与分析

```bash
python scripts/word_tool.py get_document_info report.docx
python scripts/word_tool.py get_document_text report.docx
python scripts/word_tool.py get_document_outline report.docx
python scripts/word_tool.py list_available_documents --directory .
```

### 批注与保护

```bash
# 获取批注
python scripts/word_tool.py get_all_comments report.docx
python scripts/word_tool.py get_comments_by_author report.docx --author "张三"
python scripts/word_tool.py get_comments_for_paragraph report.docx --paragraph-index 2

# 添加批注
python scripts/word_tool.py add_comment report.docx --paragraph-index 2 --text "请核实此数据" --author "审核员"

# 删除批注
python scripts/word_tool.py delete_comment report.docx --comment-id 1

# 密码保护
python scripts/word_tool.py add_password_protection report.docx --password "secret" --output report_protected.docx
```

### 合并与复制

```bash
python scripts/word_tool.py copy_document report.docx --destination report_backup.docx
python scripts/word_tool.py merge_documents merged.docx --sources part1.docx part2.docx part3.docx
python scripts/word_tool.py convert_to_pdf report.docx --output report.pdf
```

## 典型工作流

### 创建完整报告

1. `create_document` 创建文档
2. `add_header` / `add_footer` / `add_page_number` 设置页眉页脚
3. `add_heading` 添加标题
4. `add_paragraph` 添加正文
5. `add_table` 添加表格
6. `format_table`、`highlight_table_header` 美化表格
7. `add_picture`、`add_page_break` 按需添加
8. 重复 3–7 完成各章节
9. `add_table_of_contents` 添加目录

### 编辑现有文档

1. `get_document_text` 或 `get_document_outline` 了解结构
2. `search_and_replace` 批量替换
3. `format_text` 局部格式
4. `set_paragraph_alignment` / `set_paragraph_indent` 调整段落格式
5. `insert_header_near_text` / `insert_line_or_paragraph_near_text` 在指定位置插入

### 填写表格模板

1. `get_table_info --show-merged true` 查看表格结构（包含合并单元格信息）
2. `set_table_cell --visual true` 或 `batch_set_table_cells --visual true` 填写单元格（使用视觉索引处理合并单元格）
3. `add_table_row` / `delete_table_row` 按需调整行数
4. `highlight_table_header` / `format_table` 美化表格

### 分析文档

1. `get_document_info` 元信息
2. `get_document_text` 全文
3. `get_all_comments` 批注
4. `get_document_outline` 大纲
5. `get_hyperlinks` 超链接

## 注意事项

- **颜色**：十六进制不带 `#`（如 `FF0000`），或标准颜色名
- **字号**：`--font-size` 为半磅，24 表示 12pt，28 表示 14pt
- **索引**：段落、表格、行、列索引均从 0 开始
- **路径**：建议使用绝对路径；从技能根执行时 `scripts/word_tool.py` 即指本技能的 scripts
- **布尔参数**：`--bold true` / `--bold false`，或 `true`/`1`/`yes`
- **JSON 参数**：`--data`、`--widths`、`--list-items` 等需传合法 JSON 字符串（注意 shell 引号）
- **合并单元格**：使用 `--visual true` 参数可按视觉列位置填写，避免合并单元格导致的索引混乱
- **目录更新**：`add_table_of_contents` 插入的目录需要在 Word/WPS 中右键"更新域"才能显示实际内容
