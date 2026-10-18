# dsh-office-skills

Office document skills for [pi](https://github.com/earendil-works/pi-coding-agent), **ported from [DeepSeek Harness (DSH)](https://github.com/deepseek-ai/deepseek-harness)**.

**Unofficial community port** — not affiliated with, endorsed by, or maintained by DeepSeek.

This is not an independent reimplementation. The skill texts come from DSH's
[`@deepseek-ai/dsh-skill-office`](https://github.com/deepseek-ai/deepseek-harness) (MIT, (c) 2026 DeepSeek) and are adapted so pi can use them. See [NOTICE](./NOTICE) for the upstream copyright notice, the full MIT text, and the exact list of changes.

## Install

```bash
pi install npm:dsh-office-skills
```

The install pulls the LibreOffice Kit engine through npm optional dependencies:
your platform's native engine (~150-190 MB unpacked) plus the universal WASM
fallback (~145 MB), about 335 MB on Windows x64. No system LibreOffice and no
Microsoft Office installation are required.

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

Ask pi to build or edit an Office file and it loads the matching skill. Force the
router with `/skill:office`.

## Requirements

- pi (which itself requires Node.js >= 22.19).
- Python 3.9+ for the authoring skills. Missing libraries (python-docx,
  python-pptx, openpyxl, pandas) are installed with pip only after you agree;
  see [python/SETUP.md](./python/SETUP.md) and `python/requirements.txt`.

## Keeping up with upstream (DSH)

The port is mechanical and reproducible. `upstream.json` records which upstream
version this package is based on plus a sha256 per upstream skill file.

```bash
npm run check:upstream   # compare the baseline with the newest published upstream version
npm run apply:upstream   # re-port from the newest published upstream version
```

`scripts/sync-upstream.mjs` downloads the upstream tarball, applies the adaptations
below, and **fails loudly** when an upstream sentence it must rewrite has changed.
Drift is never silently dropped: update `RULES` in that script and re-run.

| Upstream (DSH) | This port (pi) |
| --- | --- |
| `skill` tool invocation | read the sibling skill's `SKILL.md` |
| `load_workspace_dependencies` | `python/SETUP.md` consent-first flow |
| node/CLI paths injected at load time | `scripts/resolve-dsoffice.mjs` resolves them at runtime |
| `present({"files": [...]})` delivery | report the final file path |

Hand-maintained files that the sync script never rewrites:
`skills/office/references/sidecar.md` (rewritten for npm-installed dependencies),
`python/`, and `scripts/resolve-dsoffice.mjs`.

## License

- Skill texts: MIT, derived from `@deepseek-ai/dsh-skill-office` — Copyright (c) 2026 DeepSeek. The upstream notice is reproduced in [NOTICE](./NOTICE); the package itself is [MIT](./LICENSE).
- Engine (`@deepseek-ai/libreoffice-kit` and platform packages): MPL-2.0, Copyright (c) DeepSeek. It is installed as an npm dependency and is not redistributed here.
