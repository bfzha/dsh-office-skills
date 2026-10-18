---
# Ported from DSH: @deepseek-ai/dsh-skill-office (MIT) — see NOTICE.
name: office-recalculate
description: Refresh spreadsheet formula caches and verify calculated values with LibreOffice Kit.
license: MIT
---

# Recalculate spreadsheet formulas

Use the installed LibreOffice Kit CLI through the shell. Resolve the absolute Node and CLI paths once per session by running, from this skill's directory, `node ../../scripts/resolve-dsoffice.mjs`; it prints `{ "libreofficeKit": { "node": "<node>", "cli": "<cli>" } }`. In Bash, invoke `"<node>" "<cli>"`; in PowerShell, invoke `& "<node>" "<cli>"`. Prefer bundled LibreOffice unless the user requests another installation. If path resolution fails or the engine is missing, follow [sidecar setup](../office/references/sidecar.md): fix or fetch the engine only after the user agrees. Do not invoke system soffice, use npx, or install a replacement to work around a failed recalculation with an installed sidecar.

Keep the original workbook and write a new file in the task workspace. Quote each input and output path.

```text
"<node>" "<cli>" recalculate --input <workbook.xlsx> --output <recalculated.xlsx>
```

Verify the saved result with openpyxl, using the Python environment from `../../python/SETUP.md` unless a project environment is required. Open it with both `data_only=False` and `data_only=True`. Check that formulas survive and cached results match the requested logic, expected totals, and cell types. Check required unchanged sheets, styles, and data. Recalculation does not validate business logic or guarantee Excel-specific feature compatibility.

If no supported recalculation engine is available or the command fails, preserve the workbook and explain that calculated values could not be refreshed. Do not replace requested formulas with constants or present stale/absent caches as fresh results. Deliver the verified output by reporting its final path, keeping the file in place; use the session's file-delivery mechanism if it provides one.
