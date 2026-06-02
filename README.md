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
