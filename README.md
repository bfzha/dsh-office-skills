# dsh-office-skills

Office document skills for [pi](https://github.com/earendil-works/pi-coding-agent), ported from [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness) (DSH): create and edit Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) with Python, and render pages to images, recalculate spreadsheet formulas, and export to PDF with the bundled LibreOffice Kit engine.

## Install

```bash
pi install npm:dsh-office-skills
```

The install downloads the LibreOffice Kit engine via npm optional dependencies: your platform's native engine (~150-190 MB unpacked) plus the universal WASM fallback (~145 MB), about 335 MB on Windows x64. No system LibreOffice or Microsoft Office is required.

## Skills provided

| Skill | Use |
| --- | --- |
| `office` | Router: picks the right skill for the task |
| `office-docx` | Create, read, edit, and validate Word documents (python-docx) |
| `office-xlsx` | Read, create, edit, and validate Excel workbooks, data, and formulas (openpyxl, pandas) |
| `office-pptx` | Create, read, edit, and validate PowerPoint presentations (python-pptx) |
| `office-render` | Render requested pages/sheets/ranges to images and inspect visual layout |
| `office-recalculate` | Refresh spreadsheet formula caches and verify calculated values |
| `office-export-pdf` | Convert an Office document to PDF |

Just ask pi to build or edit an Office file; it loads the matching skill. Force the router with `/skill:office`.

## Requirements

- pi (which itself requires Node.js >= 22.19).
- Python 3.9+ for the authoring skills. Missing libraries (python-docx, python-pptx, openpyxl, pandas) are installed with pip only after you agree; see `python/SETUP.md`.

## Provenance and license

**This is an unofficial community port** — not affiliated with, or endorsed by, DeepSeek. All credit for the skill texts and the document engine goes to the DeepSeek Harness team. The package name keeps the `dsh` prefix to mark that origin.

- Skill texts: ported from [`@deepseek-ai/dsh-skill-office@0.2.1-rc.2`](https://github.com/deepseek-ai/deepseek-harness/tree/main/packages/skill/skill-office) (MIT, Copyright (c) 2026 DeepSeek). The original copyright notice is preserved in LICENSE and NOTICE, and every SKILL.md carries a source comment.
- Document engine: [`@deepseek-ai/libreoffice-kit`](https://github.com/deepseek-ai/dsh-libreoffice-kit) (MPL-2.0, Copyright (c) DeepSeek), consumed as an npm dependency — not redistributed in this repository.
- The pi packaging itself (`python/SETUP.md`, `scripts/resolve-dsoffice.mjs`, `scripts/sync-from-dsh.mjs`) is original to this port (MIT).

## Tracking upstream DSH updates

`scripts/sync-from-dsh.mjs` regenerates `skills/*/SKILL.md` from an upstream `dsh-skill-office` assets root, re-applying the pi adaptations with exact-match assertions:

```bash
npm pack @deepseek-ai/dsh-skill-office@<version>
tar -xzf deepseek-ai-dsh-skill-office-<version>.tgz
node scripts/sync-from-dsh.mjs ./package/assets
```

If upstream reworded a sentence, the script stops with the offending text instead of silently drifting; update the replacement table inside the script and commit both together. `skills/office/references/sidecar.md` is a pi-specific rewrite — diff it by hand against upstream `assets/office/references/sidecar.md` when syncing.
