# SACCR tooling

Requirements: Git and Python 3.10+ on PATH. `check.py` needs no Python packages
and no Excel. Run from the repository root before every push:

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
| `check_source` | Every tracked text file is stored with LF in Git (CRLF, mixed or lone-CR blobs declared as text fail); VBA components sit only in the locations defined in `docs/REPOSITORY_STRUCTURE.md`, and modules in `src/core/` declare `Option Private Module`; exported VBA (`.bas`, `.cls`, `.frm`) checks out as CRLF, decodes as cp1252, has `Option Explicit` and a `VB_Name` equal to its filename and unique in the project; each `.frm` references a tracked `.frx` large enough for its offset; `CHANGELOG.md` starts with `## [Unreleased]`, uses `## [X.Y.Z] - YYYY-MM-DD` release headings with real calendar dates and has a link reference for each |
| `check_committed_whitespace` | `git diff --check` on staged and unstaged changes (local) or on the committed range (`--ci`) |
| `check_vba_jumps` | Every `GoTo`, `GoSub`, `Resume` and `On Error GoTo` target is a label in the same procedure |
| `check_vba_conditionals` | `#If`/`#ElseIf`/`#Else`/`#End If` are balanced and use only `VBA6`, `VBA7`, `Win32`, `Win64`; `Declare` in reachable 64-bit branches is `PtrSafe`; no `#Const` |
| `check_vba_public_api` | Every `Public` declaration in `src/modules/` is listed, with its exact signature, in `docs/PUBLIC_API.txt`, and nothing else is; no implicit public procedures; one identifier per public `Const` or variable; no name collisions |
| `*-fixtures` | Each VBA checker and the whitespace gate run their own positive and negative self-tests first |

These are static checks only. None of them compiles VBA, opens Excel, runs a
test harness or validates any SA-CCR number, and a pass is never evidence that
the workbook works in Excel.

## VBA checkers: source and adaptations

`check_vba_jumps.py`, `check_vba_conditionals.py` and `check_vba_public_api.py`
come from EXCEL-VBA-PROJECT-TEMPLATE at commit
`b903fe44ef6a032c4689870b83745afa1c22490d`. SACCR now evaluates jumps per reachable compilation environment and checks
`PtrSafe` in the declaration modifier position. The
public-API checker takes each component's role from its folder, as defined in
[`docs/REPOSITORY_STRUCTURE.md`](../docs/REPOSITORY_STRUCTURE.md): `src/modules/`
is public; the other `src/` folders are internal; `tests/` and `examples/` are
test and example. The template instead reads a component registry from
`.github/repository-profile.json`, which SACCR does not use. `test_tooling.py`
adds fixtures for that role mapping. `test_review_regressions.py` covers the
post-merge findings, including colon-separated public declarations.

### Direct-call validation: evaluated, not adopted yet

VBA-DATETIMEPICKER's `check_vba_calls.py` resolves direct, module-qualified and
typed-class calls, argument counts and literal `Application.Run`/`OnAction`
targets. It reports anything it cannot resolve (calls through `Object` or
`Variant`, library members, dynamic names) as *unknown*, not as a defect. That
repository runs it as advisory only, because it needs a per-project manifest of
VBA projects and its value grows with the number of modules.

SACCR's only VBA so far is a three-module scaffold and harness, so there is
little for it to resolve. It should be adopted, advisory first, once the engine
spans several core modules and a facade.
Until then the VBA compiler in Excel is the call-resolution check.

## GitHub Actions

`.github/workflows/static-checks.yml` runs the same command with `--ci` on
pull requests targeting `main` or `release/**`, pushes to those branches, and
manual dispatch. It is installed on both `main` and `release/0.0.1`.
The job is named **Repository integrity**.

Pull requests pass their base SHA through `--base` to check the complete
proposed change. Checkout fetches the full history and verifies the exact
event SHA; for pull requests this is GitHub's test merge commit.
Push and manual runs use the runner's first-parent whitespace scope.

All gates must pass. Reports and logs from `test-results/` are uploaded
even when checks fail and retained for 30 days. The artifact name includes
the candidate SHA, run ID and attempt; `static-checks.json` also records the
candidate SHA. Missing evidence fails the upload step.

The workflow uses a read-only token, does not persist checkout credentials,
and pins actions to full commit SHAs taken from the template.
It uses GitHub-hosted Ubuntu and Python 3.10.
A successful job is static-check evidence only, not Excel validation.

After `check.py`, the same job runs three quality steps. Each runs even if an
earlier step failed, so one run reports every problem, and each failure fails the
job:

| Step | Checks |
| --- | --- |
| Ruff | `ruff check tools` with the rules in `pyproject.toml`; a probe proves the complexity ceiling of 15 is active |
| mypy | `mypy --strict` over every module in `tools/`, tests included; a probe proves an unannotated function is rejected |
| actionlint | Every workflow in `.github/workflows/`, including shellcheck on `run:` scripts |

Ruff and mypy install from `tools/requirements-quality-ci.txt` with
`--require-hashes --only-binary=:all:`. The actionlint archive is downloaded at a
pinned version and verified against its SHA-256 before use. Their outputs are in
the uploaded report. To run Ruff and mypy locally, with Python 3.10:

```sh
python -m pip install --require-hashes --only-binary=:all: -r tools/requirements-quality-ci.txt
ruff check tools
mypy
```

## Dependency updates

`.github/dependabot.yml` asks Dependabot for weekly version updates of the
GitHub Actions pinned in the workflows. They are opened against the active
release branch (`target-branch`) and labelled `ci` and `P3`. Dependabot reads
this file from the default branch; PR #24 integrated it into `main`. Update `target-branch` whenever a new release branch is opened.

Review each update pull request like any other:

1. Check the new version's release notes and the commit it pins. The pin must
   stay a full 40-character SHA with the version as a trailing comment.
2. Let **Repository integrity** run, including actionlint, and confirm it is
   green.
3. Squash-merge it yourself. Nothing is merged or approved automatically, and
   Dependabot alerts are never dismissed automatically.

The quality tools are not covered by Dependabot. To change Ruff, mypy or
actionlint, update the version and its hash or checksum in
`tools/requirements-quality-ci.txt` or the workflow `env` in one reviewed pull
request.
Whether GitHub blocks merging on this check depends on the repository's
branch rulesets.
