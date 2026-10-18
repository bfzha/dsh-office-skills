---
# Ported from DSH: @deepseek-ai/dsh-skill-office (MIT) — see NOTICE.
name: office-xlsx
description: Read, create, edit, and validate Excel workbooks, data, and formulas (.xlsx).
license: MIT
---

# Excel workbooks

Follow an explicit user or applicable AGENTS.md requirement for an environment or library. Otherwise follow `../../python/SETUP.md` (relative to this skill's directory) to select a Python environment with `openpyxl` (and `pandas` for analysis) available; it defines the consent-first install for missing libraries. Do not install packages without the user's agreement. If no usable environment can be configured, report a missing dependency only when it prevents the requested operation.

Keep scripts, working files, and outputs in the task workspace. The runtime and skill directory are shared read-only resources. Save to a new workbook unless the user requests an in-place edit.

Name Python scripts for their task, such as `edit_xlsx.py`; avoid module names such as `inspect.py`, `openpyxl.py`, or `pandas.py` that shadow imported libraries.

## Read and modify

Use `openpyxl` for existing XLSX workbooks and targeted cell or style changes. Load with `data_only=False` when formulas must survive. Inspect sheet names, the affected cell types, formulas, styles, merged ranges, and referenced ranges before editing. Check the saved file by reopening it, including unchanged content that the request requires preserving.

Use pandas for data analysis and transformations. A DataFrame is not the workbook: exporting it over an existing file can lose sheets, formulas, charts, and formatting. Write analysis results back to the intended ranges with openpyxl. `XlsxWriter` can create new workbooks, but cannot read or modify existing ones.

For a new workbook, the selected Python environment can write editable values, styles, and formulas directly:

```python
from openpyxl import Workbook
from openpyxl.styles import Font

workbook = Workbook()
sheet = workbook.active
sheet.title = "Revenue"
for row in [("Quarter", "Revenue"), ("Q1", 12), ("Q2", 18), ("Total", "=SUM(B2:B3)")]:
    sheet.append(row)
for cell in sheet[1]:
    cell.font = Font(bold=True)
sheet.column_dimensions["A"].width = 18
sheet.column_dimensions["B"].width = 18
workbook.save("report.xlsx")
```

Preserve numbers, dates, booleans, and identifiers as the intended cell types; formatting is not a type conversion. For modifications, load the existing file with `openpyxl.load_workbook("input.xlsx", data_only=False)`, change the requested ranges, and save a separate result. Avoid a DataFrame round trip when workbook features must survive.

Writing a formula does not calculate its result. openpyxl and XlsxWriter do not evaluate formulas; `data_only=True` returns cached values that may be absent or stale. When calculated results must be refreshed, load `office-recalculate`. Do not replace requested formulas with constants or report stale values as newly calculated results.

Do not rename `.xls`, `.xlsb`, encrypted files, or macro-enabled files to `.xlsx`. They need an appropriate supported operation. `keep_vba=True` can preserve VBA package content in an XLSM workflow, but does not execute or edit macros and does not guarantee every advanced workbook feature survives. Preserve the original and verify such requirements explicitly.

## Check and deliver

Reopen the saved workbook with openpyxl and verify the requested sheet names, values, types, formulas, styles, and totals against the source data, including unchanged content that must be preserved. These checks do not calculate formulas or judge workbook appearance.

Load `office-render` for requested images or visual inspection and whenever the user's primary task targets appearance, including cell fill colors, text alignment, spacing, or sizing. These appearance changes require rendering even without a separate visual-check request, unless the user explicitly skips it. New workbooks and data or formula edits do not require rendering solely because they include incidental basic formatting. In particular, bold headers and number formats accompanying data edits do not require loading `office-render` or running its commands.

Load `office-recalculate` when refreshed calculated values are needed and `office-export-pdf` for PDF output. Deliver after the requested data checks pass and the required operations finish.

Report the actual final workbook path in the reply so the user can open it from the workspace; keep the file in place and omit intermediate scripts and reports unless requested. If the session provides a file-delivery or presentation mechanism, use it with the same final path.
