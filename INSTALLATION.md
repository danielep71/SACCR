<div align="center">

# 📦 Installation and Developer Setup

### Set up, validate, and later import SACCR into Excel

[![Status](https://img.shields.io/badge/Status-Pre--release-6e7781?style=flat-square)](#current-status)
[![Checks](https://img.shields.io/badge/Checks-python_tools%2Fcheck.py-0969da?style=flat-square)](#run-the-checks)
[![Security](https://img.shields.io/badge/Security-Private_policy-d73a49?style=flat-square)](SECURITY.md)

<br>

**One identifiable commit · Clean import · Compile · Validate · Preserve caller state**

</div>

---

This document is authoritative for **developer setup, import, upgrade, recovery
and removal**. Contribution workflow is owned by
[`CONTRIBUTING.md`](CONTRIBUTING.md), vulnerability handling by
[`SECURITY.md`](SECURITY.md), and publication by [`RELEASING.md`](RELEASING.md).

> [!IMPORTANT]
> VBA executes with the user's Office permissions. Review the exact source,
> follow organizational macro policy, and never enable macros in an untrusted
> workbook.

<a id="current-status"></a>

## 🧭 Current status

SACCR is in its repository-setup milestone, **v0.0.1**. There is **no SA-CCR
calculation yet**: the only VBA is a neutral scaffold and the regression harness
that tests it. There is no workbook or add-in to install.

| Topic | Status |
| --- | --- |
| VBA source layout | Defined in [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) |
| Import/export conventions and host support | Defined below and in [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md) |
| Regression harness | `tests/modules/TestHarness.bas`; see [running the harness](#running-the-harness) |
| Excel evidence | Record and validator defined in [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md) |
| Released versions | None |

The only VBA so far is the setup scaffold and the regression harness; the
component list below grows as the SA-CCR engine is added.

## 🧰 Prerequisites

| Tool | Needed for |
| --- | --- |
| Git | Cloning and all repository work |
| Python 3.10 or later | `python tools/check.py`; standard library only, no packages |
| Node.js 20 or later | Optional: local label-catalogue checks |
| Microsoft Excel for Windows | Importing, compiling and running the VBA; see [supported hosts](#supported-hosts) |

<a id="supported-hosts"></a>

## 🖥️ Supported hosts

Decided by the owner on 2026-10-06 in issue #4.

| Host | Support level |
| --- | --- |
| Microsoft 365 or Excel 2016 and later, **Windows, 64-bit** | Supported target |
| Same versions, Windows, 32-bit | Best effort: kept compiling, not certified |
| Excel for Mac, Excel for the web, Excel 2013 and earlier | Not supported |

**References:** only the four defaults of a new workbook: *Visual Basic For
Applications*, *Microsoft Excel 16.0 Object Library*, *OLE Automation* and
*Microsoft Office 16.0 Object Library*. Adding any other reference needs an
issue and an update to this section.

**Evidence:** this table is a support **commitment**, not a test result. Host
evidence is recorded per change under [Validation record](#validation-record),
or as an evidence bundle bound to a commit as described in
[`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md).

## 📥 Get the source

Use a **Git clone**:

```sh
git clone https://github.com/danielep71/SACCR.git
cd SACCR
git switch release/0.0.1
```

`release/0.0.1` is the active development branch; `main` receives it only on
the owner's request.

Do not use **Code → Download ZIP**. `.gitattributes` excludes repository
plumbing such as `.gitattributes`, `.gitignore` and `.editorconfig` from GitHub
archives, so a ZIP is not a complete checkout and the checks will not behave as
documented.

A clone checks VBA files out with the CRLF line endings the VBE expects. Do not
download individual modules through GitHub's **Raw** view: it serves the LF
form stored in Git.

<a id="run-the-checks"></a>

## ✅ Run the checks

From the repository root:

```sh
python tools/check.py
```

All gates must pass. Reports and logs are written to the ignored
`test-results/` directory. Before pushing a committed candidate, the stricter
form also requires a clean tree:

```sh
python tools/check.py --ci
python tools/check.py --ci --base <base-sha>   # for a pull request's full range
```

Optional label-catalogue checks:

```sh
node .github/scripts/labels-sync.mjs --policy .github/label-policy.json --self-test
node .github/scripts/labels-sync.mjs --policy .github/label-policy.json --mode validate
node .github/scripts/labels-drift.mjs --policy .github/label-policy.json --self-test
```

The gates are described in [`tools/README.md`](tools/README.md). They are static
checks: they never compile VBA or run Excel.

<a id="importing-vba-into-excel"></a>

## 📂 Importing VBA into Excel

The layout and file conventions are defined in
[`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) and
[`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md#export-compatibility).
The exact component list is added here with the first VBA source.

### Building a workbook from source

1. Start from one exact commit in a Git checkout, so `.bas`, `.cls` and `.frm`
   files have CRLF line endings. Never mix components from different commits.
2. Create a new blank workbook and save it as **Excel Macro-Enabled Workbook
   (`.xlsm`)** outside the checkout, or in an ignored location such as
   `test-results/`. Workbooks are never committed.
3. Open the VBE (**Alt+F11**) and check **Tools → References** shows only the
   four default references listed under [supported hosts](#supported-hosts).
4. Import with **File → Import File**, in this order:
   1. every file in `src/core/`;
   2. every file in `src/classes/`;
   3. every file in `src/modules/`;
   4. every `.frm` in `src/forms/` (its `.frx` loads with it; never import a
      `.frx`);
   5. for a development workbook only: `tests/` and `examples/` modules.

   Current components, in import order:

   | # | File | Component | Role |
   | ---: | --- | --- | --- |
   | 1 | `src/core/CoreScaffold.bas` | `CoreScaffold` | Internal; neutral checked division for the scaffold |
   | 2 | `src/modules/SaccrScaffold.bas` | `SaccrScaffold` | Public facade (`docs/PUBLIC_API.txt`) |
   | 3 | `tests/modules/TestHarness.bas` | `TestHarness` | Regression harness; development workbook only |
5. **Document modules** in `src/workbook/` (`ThisWorkbook.cls` and sheet
   modules) cannot be imported: the VBE would create a new class such as
   `ThisWorkbook1`. Instead, open the `.cls` file in a text editor, copy the code
   below the `Attribute` lines, and paste it into the existing `ThisWorkbook` or
   sheet module. Sheet code names must match the file names.
6. Run **Debug → Compile VBAProject**; it must complete with no error.
7. Save, close and reopen when a clean session is needed, then run the harness
   and the specific scenario under test.

Do not paste source into arbitrarily named modules: the component name
(`VB_Name`) is part of the contract, and `tools/check_source.py` requires it to
match the file name.

### Exporting changes back

1. In the VBE, right-click each changed component → **Export File**, and save
   it over its file in the checkout, in the same folder. For a document module,
   copy its code back into the matching `src/workbook/` file below the existing
   `Attribute` lines instead.
2. Review `git diff`. The VBE silently changes the capitalization of an
   identifier everywhere it appears when one declaration changes case; revert
   case-only changes you did not intend.
3. Run `python tools/check.py`. It verifies placement, `VB_Name`,
   `Option Explicit`, `Option Private Module`, form resources, encoding and
   line endings.
4. Commit. Git stores the files with LF; you do not convert anything by hand.

<a id="running-the-harness"></a>

## ▶️ Running the harness

After importing the three components above and compiling, open the Immediate
window (**Ctrl+G**) and run:

```text
TestHarness.RunTests
```

A passing run prints a `MODE=NORMAL` line, the environment, one `CASE=` line per
case, the counts, the cleanup verdict and finally:

```text
RESULT=PASS; completeness=COMPLETE; cases=4; assertions=6; failures=0; cleanup=PASS
```

Anything else is a failure, including a run that stops early: the harness
requires exactly 4 cases and 6 assertions before it can report `PASS`, and it
raises an error after printing a failed result.

To see the failure path, run:

```text
TestHarness.RunTestsWithInjectedFailure
```

It runs the same suite with one deliberately wrong expectation and must print
`MODE=INJECTED_FAILURE` and
`RESULT=FAIL; completeness=COMPLETE; cases=4; assertions=6; failures=1; cleanup=PASS`,
followed by the suite failure error. If a run is interrupted, run
`TestHarness.ResetTests` before running again.

The harness never changes Excel settings. It checks that `Calculation`,
`DisplayAlerts`, `EnableEvents` and `ScreenUpdating` are the same after the run
as before, and reports `cleanup=FAIL` if not.

<a id="validation-record"></a>

## 🧪 Validation record

A successful import is not certification. Record at least:

```text
Commit (full SHA):
Files imported:
Excel / Office version and build:
Office bitness:
Operating system:
Compile:
Regression harness:
Specific scenario:
Cleanup:
Skipped or unverified:
```

A skipped, incomplete or cleanup-failed run is not a pass. Static checks cannot
substitute for Excel execution.

When the run must be bound to an exact commit, for example for a release or an
issue's acceptance evidence, record it as an evidence bundle and validate it
with `tools/check_excel_evidence.py`; see
[`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md).

## ⬆️ Upgrade

Before moving to a newer commit:

1. read the changelog between the two commits;
2. back up the host and any user configuration;
3. identify the complete component set at the target commit; and
4. review compatibility notes and known limitations.

Replace the whole component set, compile again and repeat the validation. Do not
infer backward compatibility from successful compilation alone.

Treat a locally modified copy as a fork: diff it against the old and new source,
reapply modifications deliberately and retest them.

## 🧯 Troubleshooting

| Symptom | Check first |
| --- | --- |
| Compile error or missing procedure | All components come from one commit, and required references are set. |
| Ambiguous name | Remove duplicate or older copies of a component. |
| Form controls missing or corrupt | Re-import the exact `.frm` with its adjacent `.frx`. |
| Garbled accented characters after import | The module was copied from a browser or re-saved in another encoding; import the file from the checkout. |
| Different result from a reference | Confirm the commit, inputs, configuration and the independent reference used. |
| Security warning | Verify the source origin and organizational macro settings. |
| `tools/check.py` reports CRLF in Git | Renormalize with `git add --renormalize .` and commit the result separately. |

Suspected security problems follow [`SECURITY.md`](SECURITY.md), not issues.

## 🗑️ Removal

1. Run any project-owned shutdown or cleanup procedure.
2. Remove the components and integrations the project owns.
3. Compile the remaining VBA project.
4. Close and reopen, and verify nothing still depends on removed components.

Removing modules does not remove formulas, names, links, connections or other
host integrations. Remove only state the project owns.

## 📚 Related documents

- [`README.md`](README.md) — overview and status
- [`CONTRIBUTING.md`](CONTRIBUTING.md) — change and review workflow
- [`tools/README.md`](tools/README.md) — static checks and CI
- [`RELEASING.md`](RELEASING.md) — release sequence
- [`SECURITY.md`](SECURITY.md) — vulnerability and data-handling policy

---

**Installation principle:** use one identifiable commit, compile it, exercise
its real behavior in Excel, and record what was and was not validated.
