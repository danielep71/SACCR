# 📐 Methodology

This directory records **what SACCR implements and on what authority**: the
regulatory texts and their versions, assumptions, parameter sources, open
decisions, and the trace from each rule to code and tests. The numerical
test-case format is in [`TEST_CASES.md`](TEST_CASES.md).

> [!IMPORTANT]
> The engine implements CRR and Basel CRE52. Under
> [`TEST_CASES.md`](TEST_CASES.md), only Basel rules are validated, by the
> published CRE99 examples (T01–T14); no CRR rule is validated in v1.0.0
> ([decision 8](#assumptions-and-scope)). The [traceability](#traceability)
> table gives the status of each rule. The CRR and CRE52 rows of the source
> register are still to be verified against their official texts.

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
| `CRE99` | Basel Framework, chapter CRE99, *Application guidance*: the SA-CCR worked examples | Published values of the `cre99-example-*` cases. Amounts in USD thousands; intermediate results unrounded, displayed results and final EAD rounded (CRE99.20); 250 business days a year (CRE99.28) | Effective 1 January 2023, last updated 27 March 2020 | Examples 1–5: the values used by the cases checked 2026-10-09 |

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
| Sources | Decided 2026-10-09 (decisions 1 and 2, #33): `CRR` is the consolidated CRR as at a date recorded when its text is checked, with Delegated Regulation (EU) 2021/931 as amended by (EU) 2025/855 and EBA Q&A 2023_6962; `BCBS` is the Basel Framework CRE52 in force, with its effective date recorded the same way. Until then those source-register rows stay "To verify" |
| Scope | Decided 2026-10-09 (decision 3, #33): the full SA-CCR only. Simplified SA-CCR (CRR Article 281), the original exposure method (Article 282), exposures to central counterparties, securities financing transactions, CVA and RWA are outside v1.0.0 and planned in #107 (v1.2.0). Inputs that ask for them are rejected; a limitation stated only in prose is not enough |
| Margined netting sets | Decided 2026-10-09 (decision 4, #33): in scope for v1.0.0, with VM, NICA, threshold, MTA and MPOR as inputs |
| Day count | Decided 2026-10-09 (decision 5, #33): S, E, M and T in years of 365 calendar days (`DaysPerYear`); floors and MPOR in business days of 250 a year (`BusinessDaysPerYear`). A test case may set `DaysPerYear` to reproduce a published maturity |
| Currency | Decided 2026-10-09 (decision 6, #33): trade amounts are converted into the reporting currency with the FX table on Params, whose rates the user supplies; SACCR sources no market data. A test fixture holds amounts already in its calculation currency, whose rate is set to 1 |
| Interest-rate aggregation | Decided 2026-10-09 (decision 9, #33): under `BCBS` both CRE52.57 formulas are supported, chosen with `IRBucketOffset`: the bucket formula that recognises offsets across maturity buckets, and the sum of absolute bucket values. Under `CRR` only what Article 280a allows is supported, once its text is checked. Each supported formula needs a value test |
| Open treatments | Decided 2026-10-09 (decision 10, #33): the climatic-conditions factor is the one printed in CRR Article 280e, and climatic trades are rejected if the text gives none; an unrated credit reference is rejected, because the CRR mapping is not automated; other risks keep one hedging set per primary risk driver if Article 277a confirms it; inflation is treated as interest rate under its own risk-factor label; sold options and sold credit protection (Article 274(5) and (7)) are rejected in an ordinary netting set unless declared, and their special treatment comes after v1.0.0 |

<a id="open-decisions"></a>

## ❓ Open decisions

These are deliberately undecided. Each is settled by the owner in an issue and
moved into the tables above; none may be assumed by code or tests meanwhile.

| # | Decision | Notes |
| ---: | --- | --- |
| 1 | Consolidation date of the CRR text implemented | Principle decided (Sources, above); the date is recorded when the text is checked (#33) |
| 2 | Effective date of the CRE52 text implemented | Principle decided (Sources, above); the date is recorded when the text is checked (#33) |

Decisions 3 to 6, 9 and 10 were settled on 2026-10-09 in #33 and are recorded
in [Assumptions and scope](#assumptions-and-scope).

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

Paragraph and article numbers are those cited by the code and the test cases.
They are checked against the official texts when the `CRR` and `BCBS` rows of
the source register are verified ([open decisions](#open-decisions) 1 and 2).
Test IDs `T01`–`T28` are listed in
[`tests/README.md`](../../tests/README.md#ported-prototype-catalogue). Status
uses four terms:

- **Validated**: a `published` output whose rule is this one passes in Excel
  for that regime. Only these count for the
  [completeness policy](TEST_CASES.md#completeness-policy).
- **Exercised**: no published output names the rule, but a published value
  that passes depends on it, so an error in it would very probably fail that
  value. Not counted as validated.
- **Not validated**: implemented; any cases are illustrative.
- **Not implemented**: listed below the table.

| Rule ID | Requirement | VBA procedure | Test cases | Status |
| --- | --- | --- | --- | --- |
| `CRE52.1`, `CRR.274.2` | EAD = alpha x (RC + PFE), alpha 1.4 | `CORE_Engine.ComputeNettingSets` | Published: T04, T06, T07, T10, T14. Illustrative: T15, T19, T25. Relations: `TEST_Invariants` | BCBS validated; CRR not validated |
| `CRE52.2`, `CRR.274.3` | EAD of a margined netting set capped at its unmargined EAD; under CRR the cap uses C = NICA (EBA Q&A 2023_6962) | `CORE_Engine.ComputeNettingSets` | Illustrative: T15, T16, T17, T18. Relations: `TEST_Invariants` | Not validated: no published example makes the cap bind |
| `CRE52.10`, `CRR.275.1` | RC of an unmargined netting set = max(V - C, 0) | `SACCR_ReplacementCost`, `CORE_Engine.ComputeNettingSets` | Published: T08. Illustrative: `illustrative-ir-swap-unmargined` | BCBS validated; CRR not validated |
| `CRE52.18`, `CRR.275.2` | RC of a margined netting set = max(V - C, TH + MTA - NICA, 0) | `SACCR_ReplacementCost`, `CORE_Engine.ComputeNettingSets` | Published EAD T14 (Example 5) | BCBS exercised; CRR not validated |
| `CRE52.20` | PFE = multiplier x aggregate add-on | `CORE_Engine.ComputeNettingSets` | Published EADs T04, T06, T07, T10, T14. Relations: `TEST_Invariants` | BCBS exercised; CRR not validated |
| `CRE52.23`, `CRR.278.3` | Multiplier, floor 5% | `SACCR_Multiplier`, `CORE_Engine.ComputeNettingSets` | Published: T05, T13. Relations: `TEST_Invariants` | BCBS validated; CRR not validated |
| `CRE52.25` | Aggregate add-on = sum of the asset-class add-ons | `CORE_Engine.ComputeNettingSets` | Published: T09, T12. Relations: `TEST_Invariants` | BCBS validated; CRR not validated |
| `CRE52.34`, `CRR.279b` | Supervisory duration, floored at 10 business days, and adjusted notional of interest-rate and credit trades | `SACCR_SupervisoryDuration`, `CORE_Engine.ProcessTrades` | Published: T01 | BCBS validated; CRR not validated |
| `CRE52.40`, `CRR.279a` | Supervisory delta of options | `SACCR_OptionDelta`, `CORE_Engine.ProcessTrades` | Published: T02. Illustrative: T21 | BCBS validated; CRR not validated |
| `CRR-RTS.5` | CRR lambda shift for interest-rate and commodity options with negative or low prices | `SACCR_LambdaCRR`, `CORE_Engine.ProcessTrades` | Illustrative: T20, T23; T22 (rejected under BCBS) | Not validated |
| `CRE52.41` | Supervisory delta of CDO tranches | `SACCR_CDODelta`, `CORE_Engine.ProcessTrades` | `TEST_Harness` (input domain only) | Not validated |
| `CRE52.46`, `CRE52.47` | Basis and volatility transactions: own hedging sets, factor x 0.5 and x 5 | `CORE_Engine.ProcessTrades` | `TEST_Aggregation` (relations) | Not validated |
| `CRE52.48`, `CRR.279c` | Maturity factor of an unmargined trade, sqrt(min(M, 1)), M floored at 10 business days | `SACCR_MaturityFactor`, `CORE_Engine.ProcessTrades` | Published EAD T07 (9-month trade) | BCBS exercised; CRR not validated |
| `CRE52.50`, `CRE52.52`, `CRR.279c` | MPOR floors and maturity factor of a margined netting set, 1.5 x sqrt(MPOR / 250) | `CORE_Engine.LoadNettingSets`, `SACCR_MaturityFactor` | Published: T11; EAD T14 | BCBS validated (MPOR), exercised (maturity factor); CRR not validated |
| `CRE52.56`, `CRE52.57`, `CRR.280a` | Interest rate: hedging set per currency, maturity buckets, bucket formula or sum of absolutes | `CORE_Engine.ProcessTrades`, `SACCR_IREffectiveNotional`, `CORE_Engine.ComputeHedgingSets` | Published: T03. Relations: `TEST_Invariants`, `TEST_Aggregation` | BCBS validated; CRR not validated |
| `CRE52.58`, `CRE52.59`, `CRR.280b` | Foreign exchange: hedging set per currency pair, add-on = factor x abs(effective notional) | `CORE_Engine.ProcessTrades`, `CORE_Engine.ComputeHedgingSets` | None | Not validated |
| `CRE52.60`, `CRE52.61`, `CRR.280c` | Credit: entity buckets, single-factor aggregation, factor by rating (BCBS) or credit quality step (CRR) | `CORE_Engine.ProcessTrades`, `SACCR_FactorAggregation`, `CORE_Engine.ComputeHedgingSets` | Published EAD T06. Illustrative: T28. Relations: `TEST_Aggregation` | BCBS exercised; CRR not validated |
| `CRE52.64`, `CRE52.66`, `CRR.280d` | Equity: entity buckets, single-factor aggregation | `CORE_Engine.ProcessTrades`, `SACCR_FactorAggregation`, `CORE_Engine.ComputeHedgingSets` | None | Not validated |
| `CRE52.68`, `CRE52.70`, `CRR.280e` | Commodity: hedging sets energy, metals, agricultural and other, commodity-type buckets, single-factor aggregation | `CORE_Engine.ProcessTrades`, `SACCR_FactorAggregation`, `CORE_Engine.ComputeHedgingSets` | Published EADs T07, T12 | BCBS exercised; CRR not validated |
| `CRR.277a` | CRR climatic-conditions commodity hedging set | `CORE_Engine.ProcessTrades` | Illustrative: T27. Its supervisory factor is a placeholder to confirm in Art. 280e | Not validated |
| `CRR.277.1`, `CRR.280f` | CRR other-risks asset class, factor 8%; rejected under BCBS | `CORE_Engine.ProcessTrades`, `CORE_Engine.ComputeHedgingSets` | Illustrative: T24, T25, T26 | Not validated |
| `CRE52.72` | Supervisory factors, correlations and option volatilities on Params | `CORE_Engine.LoadSFTable` | Published values T01–T14, for the factors their trades use | BCBS exercised for those factors; CRR not validated |

Not implemented, as recorded on the template's Basel_vs_CRR sheet; where the
sheet says so, the user enters such trades in a form the engine supports: several margin agreements in one netting set
(CRE52, `CRR.274.4`); one margin agreement covering several netting sets
(`CRR.275.3`); the special treatment of sold options and sold credit
protection (`CRR.274.5`, `CRR.274.7`); decomposition of option combinations
(`CRR.274.6`); automated mapping of risk drivers (`CRR.277.2`, `CRR.277.3`);
simplified SA-CCR (`CRR.281`) and the original exposure method (`CRR.282`).
Collateral is an input after haircuts, and trade amounts are converted with
the FX table on Params ([decision 6](#assumptions-and-scope)); neither is
validated.

A rule is **validated** only when at least one case with a `published` or
`independent` reference passes for it in each regime it claims; see the
[completeness policy](TEST_CASES.md#completeness-policy). Illustrative cases
never count.
