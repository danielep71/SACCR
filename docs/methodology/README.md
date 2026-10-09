# 📐 Methodology

This directory records **what SACCR implements and on what authority**: the
regulatory texts and their versions, assumptions, parameter sources, open
decisions, and the trace from each rule to code and tests. The numerical
test-case format is in [`TEST_CASES.md`](TEST_CASES.md).

> [!IMPORTANT]
> The prototype engine imported from `SACCR_Calculator.xlsm` implements CRR and
> Basel CRE52, but no rule is yet validated under [`TEST_CASES.md`](TEST_CASES.md)
> and the traceability table below is still empty. Every source is still to be
> verified against its official text.

<a id="regulatory-basis"></a>

## ⚖️ Regulatory basis

Decided by the owner on 2026-10-06 in issue #8.

| Regime | Identifier | Role |
| --- | --- | --- |
| EU Capital Requirements Regulation, SA-CCR | `CRR` | **Baseline.** Default for every calculation and test case |
| Basel Framework, CRE52 | `BCBS` | Also supported; implemented as differences from the baseline |

A caller may choose the regime; when it does not, the facade uses `CRR`. Every
result reports the regime it was calculated under, and every expected result
names its regime. Where the two texts agree, the engine has one implementation
tagged with both rule identifiers. Where they differ, the difference is recorded
in [Regime differences](#regime-differences) before any code depends on it.

<a id="source-register"></a>

## 📚 Source register

Cite sources by ID in code `REFERENCE` sections, in the traceability table and
in test cases. Record the exact version date when a source is first relied on;
a later version is a new row, not an edit.

| ID | Source | Scope relied on | Version / date | Status |
| --- | --- | --- | --- | --- |
| `CRR` | Regulation (EU) No 575/2013, Part Three, Title II, Chapter 6, Section 3 (Articles 274–280f), as amended by Regulation (EU) 2019/876 and Regulation (EU) 2024/1623 | SA-CCR exposure value, replacement cost, PFE, add-ons | Consolidated text: *to record* | To verify |
| `CRR-RTS` | Commission Delegated Regulation (EU) 2021/931 | Primary risk driver and risk-category mapping; supervisory delta of interest-rate options | *to record* | To verify |
| `BCBS` | Basel Framework, chapter CRE52 | Basel SA-CCR | Effective version: *to record* | To verify |
| `BCBS-279` | BCBS, *The standardised approach for measuring counterparty credit risk exposures* (2014) | Background; Annex 4a worked examples (printed pages 22–30) | March 2014, revised April 2014 | Example 4 checked 2026-10-09 |
| `CRE99` | Basel Framework, chapter CRE99, *Application guidance*: the SA-CCR worked examples | Published values of the `cre99-example-*` cases. Amounts in USD thousands; intermediate results unrounded, displayed results and final EAD rounded (CRE99.20); 250 business days a year (CRE99.28) | Effective 1 January 2023, last updated 27 March 2020 | Example 4 checked 2026-10-09; examples 1–3 and 5 to check |

Reference values may only be taken from a source in this register. Add a row
before citing a new one.

<a id="assumptions-and-scope"></a>

## 🧭 Assumptions and scope

| Topic | Current position |
| --- | --- |
| Units | Amounts in the netting set's calculation currency, unscaled; rates and factors as decimals; periods in years. See [`VBA_HOUSE_STYLE.md`](../VBA_HOUSE_STYLE.md#units-and-domains) |
| Inputs | Trade-level data supplied by the caller; SACCR does not price trades or source market data |
| Supervisory parameters | Taken from the cited regime text and recorded in the parameter table below, never typed in code without a reference |
| Running test cases | Decided 2026-10-06 (decision 7, #44): VBA generated from the JSON by a checked tool; cases feed the engine through its input sheets. See [`TEST_CASES.md`](TEST_CASES.md#consumption) |
| Evidence for v1.0.0 | Decided 2026-10-09 (decision 8, #44): v1.0.0 validates rules only with the `published` values of BCBS 279 Annex 4 (CRE99), checked against the registered text. It has no `independent` values: the project has one maintainer, who also wrote the code under test. Every other case stays `illustrative`. Because the published examples are Basel examples, rules under `CRR` are implemented but not validated in v1.0.0, and the release says so. A later release cross-checks the results against a second source |

<a id="open-decisions"></a>

## ❓ Open decisions

These are deliberately undecided. Each is settled by the owner in an issue and
moved into the tables above; none may be assumed by code or tests meanwhile.

| # | Decision | Notes |
| ---: | --- | --- |
| 1 | Exact consolidated CRR version date to implement | Fixes the `CRR` row of the source register |
| 2 | Effective Basel CRE52 version | Fixes the `BCBS` row |
| 3 | In scope beyond the full SA-CCR: simplified SA-CCR (CRR Article 281), original exposure method (Article 282), CCP exposures | Assumed out of scope until decided |
| 4 | Margined netting sets and collateral (NICA, thresholds, MPOR) in the first engine milestone | Unmargined first is the smallest useful scope |
| 5 | Date and maturity conventions: business days, day count, floors | Part of each rule's parameter entry |
| 6 | Calculation currency and FX conversion of trade amounts | Caller-supplied converted amounts is the simplest contract |

<a id="regime-differences"></a>

## 🔀 Regime differences

Each confirmed difference between `CRR` and `BCBS` that affects a result gets a
row, with locators in both texts. No difference is asserted until both texts
have been compared for that rule.

| Topic | `CRR` (locator) | `BCBS` (locator) | Effect | Test cases |
| --- | --- | --- | --- | --- |
| *none recorded yet* | | | | |

<a id="parameters"></a>

## 🔢 Supervisory parameters

| Parameter | Value | Regime | Source locator | Used by |
| --- | --- | --- | --- | --- |
| *none recorded yet* | | | | |

<a id="traceability"></a>

## 🔗 Traceability

One row per implemented rule. Rule IDs are the regime identifier and the
article or paragraph, for example `CRR.275.1` or `CRE52.<paragraph>`; `CRE52` alone
marks a paragraph not yet located. A rule shared by
both regimes lists both IDs.

| Rule ID | Requirement | VBA procedure | Test cases | Status |
| --- | --- | --- | --- | --- |
| *none implemented yet* | | | | |

A rule is **validated** only when at least one case with a `published` or
`independent` reference passes for it in each regime it claims; see the
[completeness policy](TEST_CASES.md#completeness-policy). Illustrative cases
never count.
