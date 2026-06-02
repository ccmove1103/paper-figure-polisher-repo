# paper-figure-polisher

`paper-figure-polisher` is a Codex skill for polishing `.docx` manuscripts, normalizing tables and figures, and converting spreadsheet data into journal-style three-line tables.

It is designed for manuscript cleanup workflows where table structure, caption format, figure pagination, and Word-native formatting need to be standardized with minimal manual editing.

## Features

- Convert `.csv` and `.xlsx` tables into editable Word three-line tables
- Infer medical grouping headings such as `人口学与住院信息`, `既往史与合并症`, `基线实验室与辅助检查`, `随访指标`, and `结局事件`
- Normalize existing `.docx` tables without rebuilding them from scratch
- Detect grouped heading rows in Word tables and apply heading-plus-indented-children hierarchy
- Clean up non-standard captions such as `Figure1`, `figure 1 :`, or `Table 1 :`
- Auto-generate missing `Table n.` and `Figure n.` captions in sequence order
- Keep one figure block per page and normalize figure captions and notes
- Isolate truly wide tables into landscape sections while keeping surrounding prose in portrait layout
- Repeat table header rows across page breaks and reduce row splitting where Word supports it

## Install

This repository contains the skill at:

`skills/paper-figure-polisher`

To install manually, copy that directory into your Codex skills directory:

```text
~/.codex/skills/paper-figure-polisher
```

On Windows, this is typically:

```text
C:\Users\<your-user>\.codex\skills\paper-figure-polisher
```

After copying the folder, restart Codex so the skill is discovered.

## How to Install from GitHub

If you do not use Git, you can still install this skill manually from GitHub.

### Option 1: Download as ZIP

1. Open this repository on GitHub.
2. Click the green `Code` button.
3. Click `Download ZIP`.
4. Extract the ZIP file on your computer.
5. Open the extracted folder and locate:

   `skills/paper-figure-polisher`

6. Copy that folder into your Codex skills directory:

   ```text
   ~/.codex/skills/paper-figure-polisher
   ```

   On Windows, this is usually:

   ```text
   C:\Users\<your-user>\.codex\skills\paper-figure-polisher
   ```

7. Restart Codex.

### Option 2: Download Only the Skill Folder

If you only need this skill, open:

`skills/paper-figure-polisher`

and copy the full folder contents into your local Codex skills directory.

### Verify Installation

A correct installation should contain files like:

- `SKILL.md`
- `agents/openai.yaml`
- `assets/`
- `references/`
- `scripts/`
- `tests/`

## Supported Input Types

The skill currently supports:

- `.docx` manuscripts
- `.csv` spreadsheets
- `.xlsx` spreadsheets
- existing tables embedded in `.docx`
- existing figures embedded in `.docx`

The default workflow is optimized for manuscript finishing rather than raw statistical analysis or chart generation from scratch.

## Main Scripts

### `scripts/spreadsheet_to_three_line_table.py`

Use this when spreadsheet data needs to become an editable Word table.

Capabilities:

- convert `.csv` or `.xlsx` to `.docx`
- apply three-line table borders
- normalize fonts, alignment, captions, and notes
- infer medical grouping headings from row labels
- read fixed grouping rules from `assets/medical-grouping-rules.json`

### `scripts/normalize_docx_tables.py`

Use this when a `.docx` already contains tables and they need formatting repair.

Capabilities:

- normalize existing Word tables into the same three-line style
- preserve content and row order
- clean or auto-generate `Table n.` captions
- preserve merged-cell structure
- improve multi-page table appearance
- isolate wide tables into landscape sections

### `scripts/paginate_docx_figures.py`

Use this when a `.docx` contains figures that need pagination and caption cleanup.

Capabilities:

- insert page breaks so each figure starts on a new page
- normalize `Figure n.` captions and figure notes
- auto-generate missing figure captions
- clean up non-standard figure caption prefixes

## Grouping Rules

Default medical grouping rules live in:

`skills/paper-figure-polisher/assets/medical-grouping-rules.json`

Edit that file if you want to rename headings, change keyword coverage, or customize how follow-up and outcome rows are grouped.

## Tests

A small regression suite is included at:

`skills/paper-figure-polisher/tests/run_regression_tests.py`

It covers:

- plain spreadsheet-to-table conversion
- inferred medical grouping headings
- existing DOCX table caption cleanup and auto-generation
- existing DOCX figure caption cleanup and auto-generation
- complex DOCX table normalization behavior

Run it with:

```bash
python skills/paper-figure-polisher/tests/run_regression_tests.py
```

---

## 中文说明

`paper-figure-polisher` 是一个面向医学论文排版的 Codex Skill，主要用于整理 `.docx` 文稿、规范表格和图表格式，并将 `csv/xlsx` 数据转换为期刊风格的三线表。

它适合用于论文初稿、投稿稿件、统计结果表和图表整理等场景，重点是把 Word 里的表格结构、caption 格式、图表分页和样式规则统一起来，减少手工调整。

### 功能列表

- 将 `.csv` 和 `.xlsx` 转换为可编辑的 Word 三线表
- 自动识别医学分组标题，例如 `人口学与住院信息`、`既往史与合并症`、`基线实验室与辅助检查`、`随访指标`、`结局事件`
- 整理已有 `.docx` 表格，不必从头重建
- 自动识别分组主标题，并整理为“黑体加粗 + 子项缩进”的层级结构
- 清理不规范 caption，例如 `Figure1`、`figure 1 :`、`Table 1 :`
- 自动补全缺失的 `Table n.` 和 `Figure n.` 编号
- 让单个图表独占一页，并规范图题与图注
- 将真正超宽的表格单独放入横向分节，同时保留周围正文为纵向
- 支持跨页表格的表头重复与尽量减少断行

### 安装方法

这个仓库里已经包含了 skill，路径是：

`skills/paper-figure-polisher`

如果要手动安装，把这个文件夹复制到你的 Codex skills 目录中：

```text
~/.codex/skills/paper-figure-polisher
```

在 Windows 上通常是：

```text
C:\Users\<your-user>\.codex\skills\paper-figure-polisher
```

复制完成后，重启 Codex 即可识别该 skill。

### 从 GitHub 安装

如果你不会用 Git，也可以直接从 GitHub 下载并安装。

#### 方式一：下载 ZIP

1. 打开这个 GitHub 仓库。
2. 点击绿色的 `Code` 按钮。
3. 点击 `Download ZIP`。
4. 解压下载好的压缩包。
5. 找到其中的：

   `skills/paper-figure-polisher`

6. 把这个文件夹复制到你的 Codex skills 目录：

   ```text
   ~/.codex/skills/paper-figure-polisher
   ```

   Windows 上一般是：

   ```text
   C:\Users\<your-user>\.codex\skills\paper-figure-polisher
   ```

7. 重启 Codex。

#### 方式二：只复制 skill 文件夹

如果你只需要这个 skill，可以直接打开：

`skills/paper-figure-polisher`

把整个文件夹内容复制到本地 Codex skills 目录即可。

### 支持的输入类型

当前支持：

- `.docx` 文稿
- `.csv` 表格数据
- `.xlsx` 表格数据
- 已有的 `.docx` 表格
- 已有的 `.docx` 图表

这个 skill 更适合“论文排版整理”，不是用来做统计分析或从零生成图表的。

### 主要脚本说明

#### `scripts/spreadsheet_to_three_line_table.py`

用于把表格数据转换成 Word 三线表。

功能包括：

- 将 `.csv` 或 `.xlsx` 转成 `.docx`
- 自动应用三线表边框
- 统一字体、对齐、表题和表注
- 根据行内容识别医学分组标题
- 从 `assets/medical-grouping-rules.json` 读取固定分组规则

#### `scripts/normalize_docx_tables.py`

用于整理已经存在的 `.docx` 表格。

功能包括：

- 将已有 Word 表格规范成三线表
- 保留原始内容和行顺序
- 清理或自动生成 `Table n.` 表题
- 保留合并单元格结构
- 改善跨页表格显示
- 将真正宽表单独放入横向分节

#### `scripts/paginate_docx_figures.py`

用于整理 `.docx` 中的图表分页和图题。

功能包括：

- 插入分页，让每个图表从新页开始
- 统一 `Figure n.` 图题和图注
- 自动补全缺失的图题编号
- 清理不规范的图题前缀

### 分组规则

默认医学分组规则文件在：

`skills/paper-figure-polisher/assets/medical-grouping-rules.json`

如果你想修改分组标题、关键词覆盖范围，或者调整随访和结局的归类方式，可以直接编辑这个文件。

### 测试

仓库中包含一个小型回归测试：

`skills/paper-figure-polisher/tests/run_regression_tests.py`

它覆盖了：

- 表格转换
- 医学分组标题识别
- 现有 DOCX 表格 caption 清理与自动补全
- 现有 DOCX 图表 caption 清理与自动补全
- 复杂 DOCX 表格整理行为

运行方式：

```bash
python skills/paper-figure-polisher/tests/run_regression_tests.py
```
