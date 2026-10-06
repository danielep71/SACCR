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

SACCR is in its repository-setup milestone, **v0.0.1**. There is **nothing to
install into Excel yet**: the repository holds tooling and documentation, but no
VBA source, workbook or add-in.

| Topic | Status |
| --- | --- |
| VBA source layout | To be defined in issue #3 |
| Import/export conventions and host support | To be defined in issue #4 |
| Regression harness | To be defined in issue #7 |
| Excel evidence and smoke test | To be defined in issue #9 |
| Released versions | None |

The Excel sections below state the rules that will apply. They will be
completed with concrete component lists and commands as those issues land.

## 🧰 Prerequisites

| Tool | Needed for |
| --- | --- |
| Git | Cloning and all repository work |
| Python 3.10 or later | `python tools/check.py`; standard library only, no packages |
| Node.js 20 or later | Optional: local label-catalogue checks |
| Microsoft Excel for Windows | Importing, compiling and running VBA, once source exists |

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

All four gates must pass. Reports and logs are written to the ignored
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

## 📂 Importing VBA into Excel

These rules apply once VBA source exists. The component list and import order
will be added with issue #4.

1. Back up the destination workbook or add-in and any user data.
2. Use one exact commit from a Git checkout. Never mix components from
   different commits or local exports.
3. Import every required component through the VBE (**File → Import File**).
   Import the `.frm`; its adjacent `.frx` is loaded with it and must never be
   imported or edited as text.
4. Run **Debug → Compile VBAProject**.
5. Save in a macro-enabled format.
6. Close and reopen the host when a clean session is required.
7. Run the regression harness and the specific scenario under test.

Do not paste source into arbitrarily named modules. Component identity
(`VB_Name`) and form resources are part of a reproducible import.

### Exporting changes back

Export changed components from the VBE over their files in the checkout, then
let Git normalize line endings on commit. Run `python tools/check.py` before
committing: it verifies `VB_Name`, `Option Explicit`, form resources and line
endings.

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
