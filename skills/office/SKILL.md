---
# Ported from DSH: @deepseek-ai/dsh-skill-office (MIT) — see NOTICE.
name: office
description: Create, edit, inspect, or convert Office documents, spreadsheets, presentations, and PDFs. Load to choose the relevant workflow.
license: MIT
---

# Office task routing

Before creating, reading, or editing Office files, load the format skill for the task by reading its SKILL.md from the sibling directory next to this file. Standalone rendering, PDF export, and recalculation need only their operation skill.

- Create, read, or edit: `office-docx` (Word), `office-pptx` (PowerPoint), or `office-xlsx` (Excel).
- Images, visual inspection, or tasks primarily about layout, spacing, sizing, or alignment: `office-render`, even without an explicit visual-check request.
- Refresh spreadsheet formula results: `office-recalculate`.
- Export an Office file to PDF: `office-export-pdf`.

Honor requests to skip rendering. New files, content edits, formulas, and incidental formatting such as bold headers or number formats need no render unless visual inspection is requested.

The sibling skill directories sit next to this file: `../office-docx`, `../office-pptx`, `../office-xlsx`, `../office-render`, `../office-recalculate`, `../office-export-pdf`. Read the chosen skill's SKILL.md before acting. Do not reload a skill you have already read in this session.
