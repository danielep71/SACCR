<div align="center">

# 🏦 SACCR

### Counterparty credit risk under SA-CCR, built in Excel/VBA

[![Excel VBA](https://img.shields.io/badge/Excel_VBA-Windows-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](#status)
[![Status](https://img.shields.io/badge/Status-Pre--release-6e7781?style=for-the-badge)](#status)
[![Branch](https://img.shields.io/badge/Branch-release%2F1.0.0-6f42c1?style=for-the-badge)](https://github.com/danielep71/SACCR/tree/release/1.0.0)
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

Repository setup, milestone **v0.0.1**, is complete; its accepted baseline is
recorded in [issue #13](https://github.com/danielep71/SACCR/issues/13). No
calculation code exists yet. SACCR uses the **application** profile: the
deliverable is a workbook built from the exported source, around a
host-independent calculation core
([repository structure](docs/REPOSITORY_STRUCTURE.md)). Setup established:

- repository conventions, static checks and CI;
- governance, issue metadata and labels;
- the methodology basis, CRR with Basel CRE52 also supported, and the numerical
  test-case format ([`docs/methodology/`](docs/methodology/README.md)); and
- Excel evidence bound to a commit ([`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md))
  and the first real Excel smoke test (#9).

Nothing has been released. See the [changelog](CHANGELOG.md) for what has been
added so far.

## 🌿 Development

Development takes place on the active release branch, `release/1.0.0`. Every
change goes through a task branch and a pull request into that branch. Changes
to `main` require an explicit owner instruction.

## 🚀 Getting started

```sh
git clone https://github.com/danielep71/SACCR.git
cd SACCR
git switch release/1.0.0
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
| [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md) | Manual Excel run record, procedure and validator |
| [`docs/methodology/`](docs/methodology/README.md) | Regulatory basis, source register, numerical test-case format |
| [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md) | Branches, issue metadata, review checklist, GitHub controls |
| [`tools/README.md`](tools/README.md) | Static checks, CI, evidence validator and traffic export |
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
