# Manuscript Style

Use this guide when the user wants the output to follow the default manuscript layout and typography profile for `.docx` delivery.

## Page Setup

- Paper size: A4 (`210 x 297 mm`)
- Margins: `2.5 cm` on top, bottom, left, and right
- Alignment: left aligned by default; do not force full justification unless a journal explicitly requires it
- Line spacing: `1.5`
- First-line indent for body paragraphs: `0.5 cm`
- Paragraph spacing: `6 pt` before and `6 pt` after
- Text color: pure black (`RGB 0,0,0`)
- Use explicit page breaks before major sections, figures, and tables; do not simulate page breaks with repeated blank lines
- Page number: footer, centered, starting from page 1 unless the user supplies a different journal rule

## Typeface Defaults

- English body text: `Times New Roman`
- Chinese body text: `宋体`
- Chinese emphasis or heading text: `黑体`

When a paragraph mixes English and Chinese:

- keep Western characters in `Times New Roman`;
- keep Chinese characters in `宋体`;
- use `黑体` for Chinese heading-like emphasis only.

## Heading Scale

- Paper title: `Times New Roman`, `16 pt`, bold, centered
- Level 1 heading: `Times New Roman`, `14 pt`, bold, with `12 pt` before and `6 pt` after
- Level 2 heading: `Times New Roman`, `12 pt`, bold, left aligned
- Level 3 heading: `Times New Roman`, usually `12 pt`, regular or italic depending on the manuscript pattern
- Body text: `Times New Roman`, `12 pt`, `1.5` line spacing

Use heading numbering consistently when the manuscript already uses hierarchical numbering, for example:

- `1. Introduction`
- `1.1 Experimental setup`
- `1.1.1 Environmental control`

## Figure Captions

- Place the figure caption below the figure
- Format the label as `Figure 1.Title text`
- Capitalize `Figure`
- Use Arabic numerals
- Center the caption
- Omit the terminal period unless the journal explicitly requires one
- Default font: `Times New Roman`, `11 pt`
- Use italic style for the caption text when the manuscript follows the supplied baseline
- Keep a figure note directly below the figure caption when present
- Format figure notes in `10 pt`, left aligned by default
- If a figure caption is missing in an editable `.docx`, auto-generate `Figure n.` in sequence order
- If a figure caption exists but uses non-standard prefixes such as `Figure1`, `figure 1 :`, or `图1：`, normalize it to `Figure n.`

Example:

`Figure 1.Schematic diagram of the 3D canopy reconstruction process`

## Table Captions and Table Body

- Place the table caption above the table
- Format the label as `Table 1.Title text`
- Center the caption
- Default caption font: `Times New Roman`, `11 pt`, bold
- Table body font: `Times New Roman`, `11 pt`
- Table header: bold
- Keep table notes directly below the table when present
- Keep only the top rule, header separator rule, and bottom rule
- Remove vertical lines
- Center the table on the page
- Right-align numeric columns unless the journal requires decimal alignment or another numeric convention
- If a table caption is missing in an editable `.docx`, auto-generate `Table n.` in sequence order
- If a table caption exists but uses non-standard prefixes such as `Table1`, `table 1 :`, or `表1：`, normalize it to `Table n.`
- For wide tables, switch to landscape automatically when practical
- For wide tables, isolate only the wide table block instead of turning unrelated neighboring content sideways
- Default wide-table isolation mode is page-style separation; switch to continuous mode only when the user explicitly wants tighter flow
- Allow a small adjustable spacer before and after the isolated wide-table block
- Repeat the header row across page breaks when the table spans multiple pages
- Prevent row splitting across pages when Word honors the layout constraint

Example:

`Table 1.Summary of experimental conditions`

## Notes and Footers

- Table notes: default `10 pt`
- Footer and page number: `Times New Roman`, `10 pt`

## Spelling and Grammar

- Run Word spelling and grammar review before final delivery
- Treat this as a manual QA step unless the editing environment provides a real proofing API

## Equations and Units

- Use Word built-in equations or MathType for formulas
- Set variables in italics, for example `x` and `n`
- Keep functions and units upright, for example `sin` and `mg/L`
- Insert a space between values and units, for example `25 °C` and `3.5 mg·L⁻¹`
- Follow SI units by default
