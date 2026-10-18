# Missing LibreOffice Kit engine

The office-render, office-recalculate, and office-export-pdf skills call the LibreOffice Kit CLI. This package declares `@deepseek-ai/libreoffice-kit` as an npm dependency, and npm installs the engine matching the current platform automatically. Work through this page only when that dependency is missing or broken.

1. Run `node ../../scripts/resolve-dsoffice.mjs` from the skill directory. Success prints the absolute `node` and `cli` paths and ends the check.
2. If it reports a missing dependency:
   - For npm or git package installs, run `pi update --extensions` to reconcile package installations, then retry the resolve script.
   - For local path installs the dependency tree is the package author's responsibility: run `npm install` inside the package directory (safe to add `--ignore-scripts`; the engine packages are prebuilt).
3. If the engine still cannot be obtained, and the user has not already authorized a download, ask whether they want to download the Office conversion engine into a task-local tools directory. Wait for explicit agreement; if they decline, continue work that does not need the engine and report the unavailable operation.

After agreement, install `@deepseek-ai/libreoffice-kit` together with the matching engine package, using the engine version declared by the main package's optional dependencies. Use an available npm in a task-local tools directory, with `--ignore-scripts`: the packages provide prebuilt binaries and do not need installation scripts. Keep the package manager's cache or store in that writable task directory too (for npm, pass `--cache <task-tools-directory>/.npm-cache`). Select the package for the operating system and architecture of the Node.js process that will run the CLI, including when Node runs under emulation:

| Runtime platform | Engine package |
| --- | --- |
| Windows x64 | `@deepseek-ai/libreoffice-kit-win32-x64` |
| Windows ARM64 | `@deepseek-ai/libreoffice-kit-win32-arm64` |
| macOS Apple Silicon | `@deepseek-ai/libreoffice-kit-darwin-arm64` |
| macOS Intel | `@deepseek-ai/libreoffice-kit-darwin-x64` |
| No matching native package, including Linux | `@deepseek-ai/libreoffice-kit-wasm` |

Prefer the matching native package; otherwise use WASM when its package supports the current platform. Check package platform restrictions before installation rather than forcing an incompatible package. Permission to download the Office engine does not include installing Python, Node.js, or a package manager; if those prerequisites are missing, explain what is needed. Run the new installation's `lib/cli.js` with the selected Node executable by path; do not assume installation changed the session's PATH. Run its `capabilities` command before continuing the requested operation.

<!-- Skill text ported from @deepseek-ai/dsh-skill-office (MIT, Copyright (c) 2026 DeepSeek): https://github.com/deepseek-ai/deepseek-harness. Unofficial pi packaging in dsh-office-skills; see NOTICE. -->
