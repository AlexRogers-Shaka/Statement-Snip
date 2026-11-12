# Developer notes (no client data here)

- `build_workbook.py`: builds `Pavilion_Servicing_Workbook.xlsx` (`python3 build_workbook.py out.xlsx pavilion-logo.png`). Its layout constants (Structure rows 7–21, Servicing rows, IM blocks) must match `LAYOUT` in `mapper.js`.
- `taskpane.template.html`: the panel source. `taskpane.html` is this file with the parser block (from the original Statement Snip) and `mapper.js` (between the MAPPER-START and MAPPER-END markers) pasted in at `/*PARSER*/` and `/*MAPPER*/`.
- `mapper.js`: matches snipped lines to template rows (P&L, balance sheet, tax return), detects year columns, and aggregates the figures.
- `mapper_test.js` and `simulated_excel_test.js`: Node tests. The second runs the full panel against a fake Excel using jsdom.
