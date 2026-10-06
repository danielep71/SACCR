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

> Not yet released. Development takes place on the active release branch,
> `release/1.0.0`; changes to `main` require an explicit owner instruction.
> Repository setup (milestone v0.0.1) closed without a release.

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
  with the [`docs/PUBLIC_API.txt`](docs/PUBLIC_API.txt) manifest.
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
- The **Repository integrity** job also lints the Python tooling with Ruff and
  strict mypy, from hash-locked versions, and the workflows with a
  checksum-verified actionlint. Dependabot proposes weekly GitHub Actions
  updates to the release branch for manual review; nothing merges
  automatically.
- A deterministic VBA regression harness, `tests/modules/TestHarness.bas`, run
  with `TestHarness.RunTests`. It supports exact, tolerance and expected-error
  assertions with stable case names, refuses to report `PASS` unless all
  expected cases and assertions ran, verifies Excel settings are unchanged, and
  prints a machine-readable `RESULT=` line. `RunTestsWithInjectedFailure`
  demonstrates the failure path. Its four cases exercise the engine's worksheet
  functions: replacement cost, maturity factor, the CDO-delta `#NUM!` result
  for an invalid tranche, and option-delta repeatability. The v0.0.1 setup
  scaffold it first tested (`CoreScaffold`, `SaccrScaffold`) is removed.
- Excel evidence bound to an exact commit:
  [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md) gives the manual
  Windows/Excel procedure and record format, and
  `tools/check_excel_evidence.py` checks a record against the candidate's
  source digests, `.github/excel-evidence-policy.json`, the default references
  and the retained harness log, including failed reports. It
  reports import, compile, test, cleanup, incomplete and unavailable outcomes
  separately and never runs Excel.
- Methodology basis and numerical test-case format: the EU CRR is the baseline
  regime and Basel CRE52 is also supported, as recorded differences.
  [`docs/methodology/`](docs/methodology/README.md) holds the source register,
  assumptions, open decisions and traceability table;
  [`TEST_CASES.md`](docs/methodology/TEST_CASES.md) defines fixture and
  expected-result files, reference classes (published, independent,
  illustrative), tolerances, case categories and the completeness policy, with
  one clearly illustrative case. No SA-CCR formula is implemented.
- Versioning conventions in [`RELEASING.md`](RELEASING.md#versioning): a root
  `VERSION` file is created when the first release is prepared, and
  `python tools/check.py` then requires it to match the newest dated changelog
  release. Release branches are `release/X.Y.Z`; tags are annotated `vX.Y.Z`.
  Development continues on `release/1.0.0`.
- Daily traffic export, as in VBA-DATETIMEPICKER: GitHub's 14-day traffic data
  is kept on the orphan `traffic-history` branch, with alert issues on spikes
  and new referrers. See [`tools/README.md`](tools/README.md#traffic-history).
- Issue forms and the pull-request template rewritten in the style of the
  Excel VBA project template: bug and feature forms ask for the regime (CRR,
  Basel CRE52 or both), evidence and alternatives; the issue chooser links to
  installation help and the methodology; the PR template covers candidate
  identity, static checks, Excel evidence, regression coverage, risk,
  provenance and SA-CCR-specific review.

- Prototype SA-CCR engine imported from the owner's `SACCR_Calculator.xlsm`
  (engine v1.1.0): `M_Config`, `M_Engine` and `M_Util` in `src/core`, the ten
  `SACCR_*` worksheet functions in `src/modules/M_Formulas` (added to
  `docs/PUBLIC_API.txt`), and the sheet-button macros in `src/workbook/M_Main`
  with the 13 document modules. It covers both regimes (CRR default, Basel
  CRE52 per netting set), the CRR other-risks class and the margined-EAD cap.
  The five standard modules are laid out in the house style (#60): module and
  procedure banners giving purpose, inputs, results and regulatory
  references, one commented declaration per line, and section comments. Their
  statements are unchanged from the prototype apart from `Option Private
  Module` in the three core modules, which the import had declared twice.
- Workbook template `src/workbook/SACCR_Template.xlsx` (#58): the prototype's
  12 sheets, formulas, named ranges and buttons with its VBA project removed.
  The workbook is built by saving it as `.xlsm` and importing the source; the
  template is part of the Excel evidence source inventory. Not yet built and
  run in Excel from the repository.

- The prototype TestCatalogue's 28 checks as versioned test cases (#59): 12
  fixtures in `tests/fixtures` with expected files per regime in
  `tests/expected`, mapped in `tests/README.md`. The 14 figures printed in BCBS
  279 Annex 4 are `published`; the 14 that came from an unavailable Python
  script or unreviewed hand calculations are `illustrative`. Each value and
  tolerance matches the catalogue. `docs/methodology/TEST_CASES.md` defines
  the full fixture vocabulary, trade-level and text results, and currencies,
  and `tools/check_test_cases.py` validates the files in `python
  tools/check.py`. The harness does not run the cases yet.

### Fixed

- Allow scheduled traffic exports without relying on a webhook payload, while
  retaining the default-branch restriction for manual dispatches (review #53).
  The analytics environment also restricts access to main independently.

- VBA jump checks resolve labels separately in each reachable compilation
  environment; mutually exclusive labels cannot hide missing targets or create
  false duplicates (PRs #21 and #24).
- Public API checks inspect colon-separated statements without splitting strings
  or named arguments; the conditional checker requires `PtrSafe` in the actual
  declaration modifier position (PRs #21 and #24).
- The cleanup example reports restoration failures separately from the primary
  error and attempts both restorations. Metadata triage includes closed issues
  and documents the maintainer blank-issue bypass (PRs #18 and #19).
- Validation hardening: a failed Excel run must cite a matching failure report, a
  timed-out run claims no results, the harness must report a Windows host, and
  changelog releases must be listed newest first (PRs #28, #30, #47 and #49).

### Known limitations

- The imported engine is not yet validated under the repository's test-case
  policy and keeps the deviations listed in `docs/REPOSITORY_STRUCTURE.md`.
  Building it from the template has not yet been verified in Excel. No automated
  check compiles VBA or runs Excel.
- Excel evidence is manual and covers one 64-bit host; 32-bit is untested.
- The methodology sources are registered but not yet verified against their
  official texts.

---

[Unreleased]: https://github.com/danielep71/SACCR/commits/main
