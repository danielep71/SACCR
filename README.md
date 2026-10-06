<div align="center">

# 🏦 SACCR

### Counterparty credit risk under SA-CCR, built in Excel/VBA

[![Excel VBA](https://img.shields.io/badge/Excel_VBA-Windows-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](#status)
[![Status](https://img.shields.io/badge/Status-Repository_setup-6e7781?style=for-the-badge)](#status)
[![Milestone](https://img.shields.io/badge/Milestone-v0.0.1-6f42c1?style=for-the-badge)](https://github.com/danielep71/SACCR/milestone/1)
[![License](https://img.shields.io/badge/License-MIT-2ea44f?style=for-the-badge)](LICENSE)

<br>

**Source-first · Synthetic data only · Static checks on every change · Excel-verified calculations**

</div>

---

## ✨ What this project is

SACCR will calculate exposure at default for derivative netting sets under the   
Standardised Approach for Counterparty Credit Risk (SA-CCR), as an Excel
workbook driven by exported, reviewable VBA source.

<a id="status"></a>

## 🧭 Status

The project is in its **repository-setup milestone, v0.0.1**. No calculation
code exists yet. SACCR uses the **application** profile: the deliverable is a
workbook built from the exported source, around a host-independent calculation
core ([repository structure](docs/REPOSITORY_STRUCTURE.md)). The current work
establishes:

- repository conventions, static checks and CI;
- governance, issue metadata and labels;
- the methodology sources and numerical test-case contract (#8); and
- the regression harness and Excel evidence (#7, #9).

Nothing has been released. See the [changelog](CHANGELOG.md) for what has been
added so far.

## 🌿 Development during repository setup

Until milestone **v0.0.1** closes, development and setup take place on
`release/0.0.1`. Every change goes through a task branch and a pull request
into that branch. Changes to `main` require an explicit owner instruction.

## 🚀 Getting started

```sh
git clone https://github.com/danielep71/SACCR.git
cd SACCR
git switch release/0.0.1
python tools/check.py
```

Requirements, the import procedure and troubleshooting are in
[`INSTALLATION.md`](INSTALLATION.md).

## 📚 Documentation

| Document | Covers |
| --- | --- |
| [`INSTALLATION.md`](INSTALLATION.md) | Developer setup, checks, Excel import and validation record |
| [`CONTRIBUTING.md`](CONTRIBUTING.md) | Workflow, commit messages, evidence and pull requests |
| [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) | Profile decision, source layout, public API boundary |
| [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md) | VBA naming, contracts, errors, Excel-state cleanup, export rules |
| [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md) | Branches, issue metadata, review checklist, GitHub controls |
| [`tools/README.md`](tools/README.md) | Static checks and the CI workflow |
| [`docs/LABELS.md`](docs/LABELS.md) | Issue-label catalogue and its workflows |
| [`RELEASING.md`](RELEASING.md) | Integration into `main` and the release sequence |
| [`SECURITY.md`](SECURITY.md) | Private vulnerability reporting and data handling |
| [`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md) | Participant behavior and conduct reporting |
| [`CHANGELOG.md`](CHANGELOG.md) | Change history |

## 🔐 Data and security

Use synthetic data only. Never commit or attach real trades, netting sets,
counterparty or collateral data, or any client or employer material. Report
suspected vulnerabilities privately as described in [`SECURITY.md`](SECURITY.md).

## 👤 Author

**Daniele Penza** — [@danielep71](https://github.com/danielep71)

## 📄 License

[MIT](LICENSE)
