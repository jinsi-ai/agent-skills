---
name: pptx
slug: deskclaw-pptx
version: 1.0.2
description: "Use this skill whenever the user wants to create, read, edit, or manipulate PowerPoint presentations (.pptx files). Triggers include: any mention of 'PowerPoint', 'PPT', '.pptx', 'slides', 'presentation', or requests to produce or edit slide decks, pitch decks, or presentation files. Also use when extracting text from slides, adding slides/tables/charts/shapes, or populating placeholders. If the user asks for a 'deck', 'slides', 'pitch', or similar deliverable as a .pptx file, use this skill. Do NOT use for Word documents, spreadsheets, or general coding tasks unrelated to presentations."
---

# PowerPoint 演示文稿操作 (deskclaw-pptx)

本技能通过 `scripts/ppt_tool.py` 提供 .pptx 的创建、编辑和内容提取能力。基于 [python-pptx](https://python-pptx.readthedocs.io/)，与 [Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server) 能力对齐，采用无状态 CLI（每次命令对给定文件 load → 操作 → save）。

## 输出规则（必读）

- **只保存到本地**：所有生成或编辑的文件保存到用户指定路径或当前工作目录，操作完成后告知用户完整的文件路径。
- **禁止向对话发送文件内容**：不要将文件内容（文本、二进制、base64 等任何形式）粘贴或发送到用户的聊天/会话中。大文件会导致整个 session 卡死。
- **如需预览**：只在对话中展示简短的摘要信息（如文档标题、页数、表格行列数等元信息），不要展示完整内容。

## 首次使用（前置条件）

**需先安装依赖**：

1. **安装 python-pptx**（必须）：  
   ```bash
   pip install python-pptx
   ```  
   未安装时运行脚本会报错并提示上述命令。

2. **调用方式**：  
   - 当前工作目录为技能根目录 `.../skills/deskclaw-pptx/` 时：  
     `python scripts/ppt_tool.py <命令> <file_path> ...`  
   - 否则使用脚本**绝对路径**：  
     `python /path/to/openclaw/skills/deskclaw-pptx/scripts/ppt_tool.py <命令> ...`

## 依赖

- **python-pptx**：`pip install python-pptx`（必须）

## 工具概览

| 类别 | 命令 | 用途 |
|------|------|------|
| 演示 | create_presentation | 创建空白演示并保存到 file_path |
| 演示 | create_presentation_from_template | 从模板创建并保存到 file_path |
| 演示 | get_presentation_info | 幻灯片数、布局数、核心属性 |
| 演示 | set_core_properties | 设置 title/subject/author/keywords/comments |
| 演示 | get_template_file_info | 只读，返回模板的布局与占位符信息 |
| 内容 | add_slide | 按 layout_index 添加幻灯片，可选 title |
| 内容 | get_slide_info | 只读，指定 slide_index 返回占位符与形状概要 |
| 内容 | extract_slide_text | 只读，提取单页文本 |
| 内容 | extract_presentation_text | 只读，提取全部幻灯片文本 |
| 内容 | populate_placeholder | 按 slide_index + placeholder_idx 填入文本 |
| 内容 | add_bullet_points | 在指定幻灯片添加项目符号列表 |
| 结构 | add_table | 在指定幻灯片添加表格 |
| 结构 | add_shape | 在指定幻灯片添加形状（类型、位置、尺寸） |
| 结构 | format_table_cell | 格式化表格单元格（文本、字号、粗体） |
| 结构 | add_chart | 在指定幻灯片添加图表（column/bar/line/pie） |

## 调用方式

所有命令以 `file_path`（.pptx）为位置参数或通过选项传入；索引均从 0 开始。

### 创建与属性

```bash
python scripts/ppt_tool.py create_presentation deck.pptx
python scripts/ppt_tool.py create_presentation_from_template template.pptx deck.pptx
python scripts/ppt_tool.py set_core_properties deck.pptx --title "季度汇报" --author "张三"
python scripts/ppt_tool.py get_presentation_info deck.pptx
python scripts/ppt_tool.py get_template_file_info template.pptx
```

### 幻灯片与占位符

```bash
python scripts/ppt_tool.py add_slide deck.pptx --layout-index 0 --title "封面"
python scripts/ppt_tool.py add_slide deck.pptx --layout-index 1 --title "目录"
python scripts/ppt_tool.py populate_placeholder deck.pptx --slide-index 1 --placeholder-idx 1 --text "第一项\n第二项"
python scripts/ppt_tool.py add_bullet_points deck.pptx --slide-index 2 --items '["要点一","要点二","要点三"]'
```

### 表格与形状

```bash
python scripts/ppt_tool.py add_table deck.pptx --slide-index 2 --rows 3 --cols 4 --data '[["A","B","C","D"],["1","2","3","4"],["5","6","7","8"]]' --left 1 --top 2 --width 8 --height 2
python scripts/ppt_tool.py add_shape deck.pptx --slide-index 2 --shape-type rectangle --left 1 --top 1 --width 2 --height 0.8 --text "标签"
python scripts/ppt_tool.py format_table_cell deck.pptx --slide-index 2 --shape-index 1 --row-index 0 --col-index 0 --bold true --font-size 14
```

### 图表

```bash
python scripts/ppt_tool.py add_chart deck.pptx --slide-index 3 --chart-type column --categories '["Q1","Q2","Q3","Q4"]' --series-names '["2024"]' --series-values '[[100,120,140,160]]' --title "季度趋势"
```

### 提取文本

```bash
python scripts/ppt_tool.py get_slide_info deck.pptx --slide-index 0
python scripts/ppt_tool.py extract_slide_text deck.pptx --slide-index 0
python scripts/ppt_tool.py extract_presentation_text deck.pptx
python scripts/ppt_tool.py extract_presentation_text deck.pptx --no-slide-info
```

## 典型工作流

### 创建完整演示

1. `create_presentation deck.pptx` 或 `create_presentation_from_template template.pptx deck.pptx`
2. `set_core_properties deck.pptx --title "标题" --author "作者"`
3. `add_slide deck.pptx --layout-index 0 --title "封面"`
4. `add_slide deck.pptx --layout-index 1 --title "目录"`
5. `populate_placeholder` 或 `add_bullet_points` 填充内容
6. `add_table` / `add_chart` / `add_shape` 按需添加
7. 重复 4–6 完成多页

### 分析现有演示

1. `get_presentation_info deck.pptx` 查看元信息与幻灯片数
2. `extract_presentation_text deck.pptx` 提取全部文本
3. `get_slide_info deck.pptx --slide-index 0` 查看单页占位符与形状

## 注意事项

- **slide_index / placeholder_idx / layout_index**：均从 0 开始。layout_index 对应 `prs.slide_layouts[index]`（常见：0 标题页，1 标题与内容，6 空白）。
- **add_table 的 shape_index**：同一页上表格的索引（先添加的为 0），用于 `format_table_cell` 的 `--shape-index`。
- **add_chart**：categories、series_names、series_values 为 JSON；series_values 为二维数组，如 `[[1,2,3,4]]` 表示一个系列。
- **形状类型**：add_shape 支持 rectangle, rounded_rectangle, oval, diamond, triangle, arrow, star, flowchart_process, flowchart_decision 等。
- **模板路径**：get_template_file_info / create_presentation_from_template 会在当前目录、./templates、./assets、./resources 中查找模板文件名；也可传绝对路径。

## 后续可扩展能力

以下能力在 [Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server) 中存在，当前 CLI 未实现，可按需后续补充：

- 模板序列生成（create_presentation_from_templates）、应用单页模板（apply_slide_template）
- 专业主题与图片效果（apply_professional_design、apply_picture_effects）
- 连接符、母版管理、切换效果（add_connector、manage_slide_masters、manage_slide_transitions）
- 图片管理（manage_image）、超链接（manage_hyperlinks）
