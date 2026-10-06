<div align="center">

# 📜 Changelog

### Release history for SACCR counterparty credit risk in Excel/VBA

[![Format](https://img.shields.io/badge/Format-Keep_a_Changelog-0969da?style=flat-square)](https://keepachangelog.com/en/1.1.0/)
[![Versioning](https://img.shields.io/badge/Versioning-SemVer-6f42c1?style=flat-square)](https://semver.org/spec/v2.0.0.html)
[![Dates](https://img.shields.io/badge/Dates-YYYY--MM--DD-217346?style=flat-square)](#date-and-version-rules)
[![Staging](https://img.shields.io/badge/Staging-Unreleased_first-d97706?style=flat-square)](#unreleased)
[![Contributing](https://img.shields.io/badge/Changes-Contribution_guide-2ea44f?style=flat-square)](CONTRIBUTING.md)

<br>

**User-visible history · Explicit compatibility · Reproducible evidence · Immutable releases**

</div>

---

All notable changes to **SACCR** are documented here.

This changelog follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and [Semantic Versioning](https://semver.org/spec/v2.0.0.html). It records
released behavior and material unreleased changes; it is not a commit log, issue
tracker, or substitute for release evidence.

---

## 🧭 Maintenance policy

- Add material changes under **Unreleased** in the same pull request as the
  behavior or documentation they describe.
- Write from the user's perspective: describe the observable result, contract,
  compatibility impact, and migration need.
- Link the owning issue or pull request when it contains useful engineering
  detail.
- Record only validation actually performed. Static checks are never Excel
  evidence; state what was and was not run in Excel.
- Move Unreleased entries into a dated version section only when a release is
  explicitly prepared.
- Do not edit a published release entry except to correct a demonstrable factual
  or link error; annotate material corrections instead of rewriting history.

See [CONTRIBUTING.md](CONTRIBUTING.md) for change and evidence requirements and
[SECURITY.md](SECURITY.md) for private vulnerability reporting.

<a id="date-and-version-rules"></a>

### Date and version rules

| Rule | Standard |
|---|---|
| Version | `MAJOR.MINOR.PATCH`, without the leading `v` in headings |
| Release heading | `## [X.Y.Z] - YYYY-MM-DD` |
| Date | Gregorian calendar date in ISO `YYYY-MM-DD` format |
| Ordering | Unreleased first; released versions newest to oldest |
| Comparison | Unreleased → latest tag; each release → preceding tag |
| Patch | Backward-compatible correction or hardening |
| Minor | Backward-compatible capability |
| Major | Incompatible public-contract change |

The repository remains below `1.0.0` while its supported surface is still
forming. Pre-release status does not excuse undocumented breaking changes.

<details>
<summary><strong>Entry categories</strong></summary>

<br>

| Category | Use for |
|---|---|
| **Added** | New supported capabilities, APIs, files, or tests |
| **Changed** | Changes to existing behavior, contracts, tooling, or documentation |
| **Deprecated** | Supported behavior scheduled for removal |
| **Removed** | Removed capabilities or compatibility |
| **Fixed** | Corrected defects |
| **Security** | Safely disclosed security corrections |
| **Documentation** | Material documentation-only changes |
| **Validation** | Evidence actually produced |
| **Compatibility** | Upgrade or migration effects |
| **Known limitations** | Deliberate, unresolved boundaries |

Use only the categories needed by a release.

</details>

---

<a id="unreleased"></a>

## [Unreleased]

> Not yet released. All development and setup work takes place on `release/0.0.1`
> until milestone v0.0.1 closes; changes to `main` require an explicit owner instruction.

### Added

- Repository-control files adopted from the Excel VBA project template:
  `.editorconfig`, `.gitattributes` and `.gitignore`. Exported VBA source is
  stored with LF in Git and checked out as CRLF; Office packages are binary and
  ignored unless an exact path is re-included.
- One local command, `python tools/check.py`, runs the static gates before every
  push: Git storage and line endings, exported VBA integrity (`Option Explicit`,
  `VB_Name`, form resources), changelog structure and whitespace. See
  [`tools/README.md`](tools/README.md).
- Automatic **Repository integrity** checks on pull requests and pushes to
  `main` and `release/**`, plus manual execution. Reports and logs identify the
  checked commit and are retained for 30 days, including failed checks.
- This changelog.
- Repository governance records the active release branch, the pull-request
  workflow, issue metadata requirements, the review checklist and the current
  GitHub controls.
- Twenty-label issue catalogue from the Excel VBA project template, with a
  workflow that reconciles live labels from `.github/labels.json` and a
  read-only daily drift check. See [`docs/LABELS.md`](docs/LABELS.md).
- Root project documents in the same form as the other Excel/VBA repositories:
  [`CONTRIBUTING.md`](CONTRIBUTING.md), [`INSTALLATION.md`](INSTALLATION.md),
  [`SECURITY.md`](SECURITY.md), [`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md) and
  [`RELEASING.md`](RELEASING.md), tailored to SACCR's pre-release state. The
  README now gives the project status, a getting-started path and a
  documentation map.
- Repository structure for the **application** profile:
  [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) records the
  profile decision, where core, facade, workbook glue, tests, fixtures, examples
  and methodology live, the dependency direction and the public-API boundary,
  with an empty [`docs/PUBLIC_API.txt`](docs/PUBLIC_API.txt) manifest.
  `python tools/check.py` now rejects VBA components outside those locations and
  core modules without `Option Private Module`.
- VBA conventions and host support:
  [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md) defines naming, units and
  domains, error contracts, Excel-state cleanup that preserves the original
  error, and export rules (cp1252, LF in Git and CRLF on checkout, forms,
  workbook policy). [`INSTALLATION.md`](INSTALLATION.md) declares the supported
  hosts (Excel 2016+/Microsoft 365 on Windows, 64-bit supported, 32-bit best
  effort, default references only) and gives the step-by-step build-from-source
  and export procedure.
- Issue forms for bugs, features and documentation, and a pull-request
  template, adapted to SACCR and its label catalogue. Forms set the assignee,
  type label and a default priority; `docs/GOVERNANCE.md` documents the manual
  milestone step and a three-search triage check for issue metadata.
- Three VBA static checks from the template run in `python tools/check.py`,
  each with its own fixtures: jump targets stay within their procedure,
  conditional compilation is balanced and uses known symbols with `PtrSafe`
  declarations, and every public facade declaration matches
  `docs/PUBLIC_API.txt`. None of them compiles VBA or checks SA-CCR results.

### Known limitations

- The repository contains no VBA source yet, so the VBA checks have nothing to
  inspect. No check compiles VBA or runs Excel.

---

[Unreleased]: https://github.com/danielep71/SACCR/commits/main
