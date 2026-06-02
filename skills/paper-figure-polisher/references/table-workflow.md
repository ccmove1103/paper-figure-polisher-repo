# Table Workflow

Use this guide when the task is to convert spreadsheet data into Word tables or normalize existing Word tables into the same three-line style.

## Scope

Supported sources for v1:

- `.csv`
- `.xlsx`
- `.docx` tables that already exist inside a manuscript

Target output for v1:

- editable table inside a `.docx` manuscript
- three-line table style
- optional grouped section headings such as `Demographics`, `Comorbidities`, and `Outcomes`

## Conversion Sequence

1. Read the spreadsheet and identify:
   - title or caption text;
   - header row;
   - units;
   - grouped rows or subtotal rows;
   - purely decorative merged cells.
2. Clean the data before formatting:
   - remove duplicate empty columns;
   - normalize number formats inside each column;
   - keep missing values explicit;
   - preserve scientific notation where meaningful.
3. Build the Word table.
4. Apply three-line styling.
5. Add caption and table notes.
6. Recheck page fit and readability.

For existing `.docx` tables:

1. Inspect whether the first row is the header row.
2. Preserve existing table content and row order.
3. Normalize borders, fonts, alignment, and caption placement.
4. Keep any nearby caption above the table and note below the table when they already exist.
5. Recheck whether wide tables need landscape orientation.

## Three-Line Table Standard

Apply these defaults unless the user provides a stricter journal template:

- Use only three primary horizontal rules:
  - top rule for the table start;
  - middle rule below the header row;
  - bottom rule at the table end.
- Remove vertical rules.
- Keep header text concise and centered or left-aligned consistently.
- Align numeric data consistently, preferably right-aligned.
- Keep text columns left-aligned unless there is a clear reason not to.
- Keep units in the header, not repeated in every cell.
- Put explanatory notes immediately below the table.
- Put the table caption above the table, centered, and formatted as `Table 1.Title text`.
- Use `Times New Roman`, `11 pt` for English table text by default.
- Use bold styling for the header row.
- Use `10 pt` for table notes unless the journal gives another size.

## Grouped Heading Standard

When the journal or manuscript style expects variables to be grouped, render the table as logical blocks under a section heading row.

Example:

- `Demographics`
- `Comorbidities`
- `Outcomes`

Preferred input patterns:

1. The spreadsheet already has a category column such as `group`, `category`, or `section`.
2. A separate JSON map tells the script which labels belong under each heading.
3. If neither is present, the script should infer grouped heading blocks from row-label content and preserve source order using the editable fixed-rule file `assets/medical-grouping-rules.json`.

For a grouping JSON map, use a structure like:

```json
{
  "Demographics": ["Sex", "Age", "Race"],
  "Comorbidities": ["Hypertension", "Diabetes", "CKD"],
  "Outcomes": ["Mortality", "Length of stay", "Readmission"]
}
```

An editable starter file is available at `assets/group-map.example.json`.

If neither a group column nor a group map is available, preserve source order and do not invent headings silently.
If meaningful content cues exist in the first column, infer concise grouped headings such as:

- `人口学与住院信息`
- `既往史与合并症`
- `基线实验室与辅助检查`
- `随访指标`
- `结局事件`

Preserve original row order and infer contiguous group blocks rather than globally reshuffling the table.

## Fixed Medical Grouping Rules

The default medical grouping dictionary lives at:

- `assets/medical-grouping-rules.json`

Edit this file when you need to:

- rename default group headings;
- change which keywords map to `人口学与住院信息`, `既往史与合并症`, `基线实验室与辅助检查`, `随访指标`, or `结局事件`;
- change how follow-up markers are recognized;
- force labels such as `死亡时间` into `结局事件`.

The spreadsheet conversion script reads this file by default and can be pointed to another rules file with:

```bash
python scripts/spreadsheet_to_three_line_table.py input.xlsx output.docx --group-rules assets/medical-grouping-rules.json
```

## Readability Rules

- Do not shrink the font below readability just to fit a wide table.
- If the table is too wide:
  - shorten verbose headers;
  - convert repeated labels into grouped notes;
  - rotate that page section if needed;
  - split the table into continuations only when unavoidable.
- Avoid merged cells unless they encode real grouping.

## Caption Rules

- Keep the table caption above the table by default.
- Use sequential numbering matching manuscript order.
- Make sure the in-text citation matches the final table number.

## Audit Checklist

- table number exists and is unique;
- caption matches the table content;
- caption is above the table and centered;
- top, middle, and bottom rules are present;
- vertical borders are absent;
- alignment is consistent by column;
- numeric columns are right-aligned when appropriate;
- units are visible;
- notes are below the table;
- in-text references are correct.

## Existing DOCX Tables

Use `scripts/normalize_docx_tables.py` when the manuscript already contains editable Word tables and the main need is formatting repair rather than data conversion.

The script should:

- preserve table content;
- preserve row order;
- keep the first row as the default header row unless the user specifies otherwise;
- apply three-line borders;
- normalize fonts, bold header row, and numeric alignment;
- center the table;
- normalize adjacent caption and note paragraphs when they are recognizable;
- auto-detect grouped heading rows in existing `.docx` tables;
- style grouped headings in bold `黑体`;
- indent child items under those grouped headings to create a clearer journal-style hierarchy;
- preserve merged-cell visual structure while styling;
- keep multiple tables in document order without disturbing surrounding body paragraphs;
- switch to landscape automatically for obviously wide tables;
- isolate only the truly wide table block, not adjacent normal-width tables;
- support wide-table isolation mode selection with `page` or `continuous`;
- support adjustable spacing before and after isolated wide-table blocks;
- repeat the header row across page breaks and reduce row splitting where possible.

Recognize table captions with patterns like:

- `Table 1.Title text`
- `表1.标题`

If the caption is missing, generate `Table n.` using encounter order inside the document.

Recognize table notes with patterns like:

- `Note. ...`
- `注：...`
