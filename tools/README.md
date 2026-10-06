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

## GitHub Actions

`.github/workflows/static-checks.yml` runs the same command with `--ci` on
pull requests targeting `main` or `release/**`, pushes to those branches, and
manual dispatch. It is installed on both `main` and `release/0.0.1`.
The job is named **Repository integrity**.

Pull requests pass their base SHA through `--base` to check the complete
proposed change. Checkout fetches the full history and verifies the exact
event SHA; for pull requests this is GitHub's test merge commit.
Push and manual runs use the runner's first-parent whitespace scope.

All four gates must pass. Reports and logs from `test-results/` are uploaded
even when checks fail and retained for 30 days. The artifact name includes
the candidate SHA, run ID and attempt; `static-checks.json` also records the
candidate SHA. Missing evidence fails the upload step.

The workflow uses a read-only token, does not persist checkout credentials,
and pins actions to full commit SHAs taken from the template.
It uses GitHub-hosted Ubuntu and Python 3.10 with no additional Python packages.
A successful job is static-check evidence only, not Excel validation.
Whether GitHub blocks merging on this check depends on branch-protection
availability and configuration.
