# Python environment procedure (office skills)

The office-docx, office-xlsx, and office-pptx skills drive `python-docx`, `openpyxl`, `python-pptx`, and `pandas`. This package does not bundle Python; when a skill points here, follow this procedure once per task.

1. Prefer the project environment. If the user or the applicable AGENTS.md specifies a Python interpreter, virtualenv, or library, use it and skip the remaining detection.
2. Detect an interpreter: run `python --version`; on Windows also try `py -3 --version`. Accept Python 3.9 or newer.
3. Check the libraries the current skill needs: `python -c "import docx"` (office-docx), `python -c "import openpyxl"` and `python -c "import pandas"` (office-xlsx), `python -c "import pptx"` (office-pptx). Every import used by the active skill must succeed.
4. If an import fails, ask the user for consent before installing anything. With consent run `python -m pip install --user -r <this-directory>/requirements.txt`, where `<this-directory>` is the directory containing this SETUP.md (the skill reads it, so its path is known). On PEP 668 systems (pip reports "externally-managed-environment"), prefer a task-local virtualenv: `python -m venv <task-workspace>/.venv-office` and install into it with that venv's pip; use `--break-system-packages` only with explicit user approval.
5. Re-run the import check after installing. Report a missing dependency only when it prevents the requested operation; never describe libraries as installed when the check failed.

Do not install a different Python, and do not install anything the user did not approve.
