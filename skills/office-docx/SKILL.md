---
name: office-docx
description: Create, read, edit, and validate Word documents (.docx).
license: MIT
---

# Word documents

Use `python-docx` for DOCX creation and ordinary edits. Follow an explicit user or applicable AGENTS.md requirement for a project environment or another library. Otherwise follow `../../python/SETUP.md` (relative to this skill's directory) to select a Python environment with `python-docx` available; it defines the consent-first install for missing libraries. Do not install packages without the user's agreement. If no usable environment can be configured, report a missing dependency only when it prevents the requested operation.

Keep source scripts, intermediate files, and final documents in the task workspace; the runtime and this skill directory are read-only resources. Use the user's requested language and preserve an existing document's design unless a redesign is requested.

Name Python scripts for their task, such as `edit_docx.py`; avoid module names such as `inspect.py`, `docx.py`, or `zipfile.py` that shadow imported libraries.

## Create and edit

For text edits, locate the target by its text or position and inspect only its runs and formatting. A title can be an ordinary paragraph without a Title style or document-property title. Update the relevant runs; assigning paragraph `.text` destroys run formatting. Save a new file unless an in-place edit is requested. Read other structures only when the edit involves them; rebuilding a document can lose unsupported features.

Use paragraph styles for headings and body text. Size tables for the section that contains them, and account for merged cells and nested tables. Chinese, Japanese, and Korean text may need an explicit `w:eastAsia` font assignment in addition to `run.font.name`; font names alone do not establish glyph availability or rendered appearance.

For a new document, use the selected Python executable:

```python
from docx import Document

document = Document()
document.add_heading("Project report", level=0)
document.add_paragraph("Summary", style="Heading 1")
document.add_paragraph("The requested findings go here.")
document.save("report.docx")
```

python-docx does not paginate or render documents. Do not represent ordinary replacement, colored text, or comments as tracked changes. When real revisions or unsupported OOXML features matter, preserve their package parts and verify the requested operation rather than silently discarding them. When writing package parts directly, create a new ZIP and copy each unchanged member once; appending a replacement member duplicates its name and corrupts the package.

## Check and deliver

Reopen the saved document with python-docx and verify the requested content, related formatting, and unchanged content that must be preserved. Deliver when these checks pass. Investigate package-level details only for a concrete failure, ambiguous edit target, unsupported feature, or explicit user requirement.

Creating a file or changing content with incidental basic formatting does not require additional operations. Load `office-render` for requested images or visual inspection. When styling, font color, spacing, sizing, or text alignment is the requested edit, load `office-render` and inspect the changed page even without a separate visual-check request, unless the user explicitly skips rendering. For PDF output load `office-export-pdf`.

Report the actual final DOCX path in the reply so the user can open it from the workspace; keep the file in place and omit temporary QA reports unless requested. If the session provides a file-delivery or presentation mechanism, use it with the same final path.

<!-- Skill text ported from @deepseek-ai/dsh-skill-office (MIT, Copyright (c) 2026 DeepSeek): https://github.com/deepseek-ai/deepseek-harness. Unofficial pi packaging in dsh-office-skills; see NOTICE. -->
