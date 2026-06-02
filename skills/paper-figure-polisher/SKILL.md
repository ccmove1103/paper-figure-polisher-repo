---
name: paper-figure-polisher
description: Organize academic manuscripts in .docx format and standardize tables and figures for submission-style workflows. Use when Codex needs to clean paper structure, convert CSV/XLSX data into three-line tables in Word, place each figure on its own page, normalize captions and numbering, or enforce consistent chart formatting such as fonts, sizes, line weights, spacing, and export layout.
---

# Paper Figure Polisher

## Overview

Use this skill as a focused paper-finishing workflow for Word manuscripts. The primary targets are `.docx` manuscripts plus table sources in `.csv` or `.xlsx`. The first version is optimized for:

- converting spreadsheet tables into three-line tables;
- forcing one figure or chart block per page;
- normalizing captions, numbering, and figure/table references;
- applying a consistent publication-style visual standard.

## Workflow

1. Confirm the working set:
   - manuscript file;
   - source table files;
   - figure files or embedded charts;
   - any target journal or lab style notes.
2. Inspect the manuscript structure before editing:
   - heading order;
   - existing figure and table numbering;
   - whether figures are embedded, linked, or separate files;
   - whether tables already exist in Word.
3. Route by task:
   - for spreadsheet-to-table conversion, read `references/table-workflow.md`;
   - for one-figure-per-page layout, read `references/figure-pagination.md`;
   - for typography, spacing, line weight, and color consistency, read `references/style-guide.md`;
   - for page setup, fonts, headings, captions, equations, and units, read `references/manuscript-style.md`.
4. Prefer minimally invasive edits:
   - preserve wording unless the user asks for rewriting;
   - preserve numbering if it is already correct;
   - change layout and formatting before changing content.
5. End with a manuscript audit:
   - table numbering is sequential;
   - figure numbering is sequential;
   - each in-text citation points to an existing figure or table;
   - each figure starts on a new page;
   - table and figure styles are consistent.

## Operating Rules

- Treat `.docx` as the source of truth for final layout.
- When converting tables from `.csv` or `.xlsx`, create editable Word tables rather than pasting screenshots.
- Use page breaks or section-aware layout so each figure starts on a fresh page.
- Keep captions attached to their target figure or table.
- Do not compress fonts or scale objects arbitrarily to force fit. If a table is too wide, shorten headers, rotate the page for that section, or split the table logically.
- When the user cites a style article or journal guide, follow it where explicit. If the source is incomplete or inaccessible, apply the closest documented style profile and state that the result is an informed approximation.
- Default to A4 paper, `2.5 cm` margins, pure black text, `1.5` line spacing, and explicit page breaks instead of stacked blank lines.

## Table Conversion Rules

For spreadsheet data:

1. Normalize the source data first:
   - header row must be unique;
   - units should be included in headers where possible;
   - blank spacer rows should be removed unless they encode grouped structure.
2. Convert to a Word table with a three-line design:
   - top border;
   - header separator border;
   - bottom border;
   - no vertical borders unless the user explicitly requests them.
3. Keep numeric alignment consistent inside a column.
4. Keep the caption above the table and notes below the table unless a journal rule says otherwise.
5. Default English table text to `Times New Roman 11 pt`; use `宋体` for Chinese body text and `黑体` for Chinese grouped headings or emphasis when needed.

Read `references/table-workflow.md` for the detailed standard.

Use `scripts/spreadsheet_to_three_line_table.py` when a spreadsheet needs to become an editable grouped three-line Word table.

Use `scripts/normalize_docx_tables.py` when a `.docx` already contains tables that need to be normalized into the same three-line style without rebuilding them from spreadsheet data. This path should also detect grouped heading rows and apply heading-plus-indented-children hierarchy when the table structure supports it.
Prefer this path for complex Word-native cases including merged cells, multiple tables mixed with prose, wide tables, and multi-page tables that need better cross-page appearance.
For wide-table blocks, prefer isolating only the actual wide table with configurable `page` vs `continuous` section behavior and adjustable spacing around the block.

## Figure Pagination Rules

- Each figure block includes panel image or chart, caption, and optional notes.
- Start each figure block on a new page.
- Keep one main figure block per page by default.
- If a figure has multiple panels, keep the full panel set together on the same page when possible.
- Avoid orphan captions separated from the figure.

Read `references/figure-pagination.md` for layout guidance.

Use `scripts/paginate_docx_figures.py` when an existing `.docx` needs page breaks inserted so each figure starts on a fresh page.
This path should also normalize detectable `Figure n.` captions and auto-generate missing figure captions in sequence when the figure block is otherwise unlabeled.

## Style Normalization Rules

Use a restrained publication-style visual profile:

- consistent manuscript-style text, with English output defaulting to Times New Roman unless the target template specifies otherwise;
- consistent axis, legend, and panel-label sizing;
- restrained line weights;
- simple fills and limited palette;
- no decorative effects, gradients, shadows, or chart junk.

Read `references/style-guide.md` before changing visual details.

## Deliverables

When asked to produce outputs, prefer:

- edited `.docx` manuscript;
- cleaned table assets converted into Word tables;
- optional audit notes listing assumptions, unresolved layout conflicts, and places where the source data or style standard was ambiguous.

## Regression Tests

Use `tests/run_regression_tests.py` when changing formatting logic, caption handling, grouping inference, or DOCX normalization behavior. The test set is intentionally small and covers:

- plain spreadsheet-to-table conversion without forced grouping;
- inferred medical grouping headings from row labels;
- existing DOCX table caption cleanup and caption auto-generation;
- existing DOCX figure caption cleanup and caption auto-generation.

## Fixed Grouping Dictionary

Keep the default medical grouping rules in `assets/medical-grouping-rules.json`.

Edit that file instead of hard-coding grouping changes into Python when the user wants to:

- rename the fixed medical heading set;
- expand or narrow keyword coverage for a heading;
- merge or split follow-up and outcome categories;
- force special labels into a specific heading.
