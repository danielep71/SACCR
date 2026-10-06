# SACCR tooling

Requirements: Git and Python 3.10+ on PATH. No Python packages or Excel
installation are needed. Run from the repository root before every push:

```sh
python tools/check.py
```

All gates must pass; findings return a nonzero exit code. Reports and logs go to
the ignored `test-results/` directory. `--ci` checks committed whitespace instead
of local edits and requires a clean tree; add `--base <base-sha>` to inspect a
pull request's whole range.

| Gate | Checks |
| --- | --- |
| `tool-tests` | Negative fixtures for the source gate (`test_tooling.py`) |
| `check_committed_whitespace-fixtures` | Self-test of the whitespace gate |
| `check_source` | Every tracked text file is stored with LF in Git (CRLF, mixed or lone-CR blobs declared as text fail); exported VBA (`.bas`, `.cls`, `.frm`) checks out as CRLF, decodes as cp1252, has `Option Explicit` and a `VB_Name` equal to its filename and unique in the project; each `.frm` references a tracked `.frx` large enough for its offset; `CHANGELOG.md` starts with `## [Unreleased]`, uses `## [X.Y.Z] - YYYY-MM-DD` release headings with real calendar dates and has a link reference for each |
| `check_committed_whitespace` | `git diff --check` on staged and unstaged changes (local) or on the committed range (`--ci`) |

These are static checks only. None of them compiles VBA, opens Excel or runs a
test harness, and a pass is never evidence that the workbook works in Excel.
