# 🧪 Tests

`tests/` holds all verification source and stable test data. Layout rules are in
[`docs/REPOSITORY_STRUCTURE.md`](../docs/REPOSITORY_STRUCTURE.md).

| Location | Contents | Defined in |
| --- | --- | --- |
| `modules/` | Regression modules; `TestHarness.bas` is the harness and its entry point | #7 |
| `fixtures/` | Synthetic inputs: trades, netting sets, collateral terms | [`TEST_CASES.md`](../docs/methodology/TEST_CASES.md) |
| `expected/` | Expected values per regime, each with its reference class and source | [`TEST_CASES.md`](../docs/methodology/TEST_CASES.md) |

The harness entry point is `TestHarness.RunTests`. How to run it and what a
passing log looks like is in
[`INSTALLATION.md`](../INSTALLATION.md#running-the-harness). Each other
subdirectory is created with its first real file.

Rules:

- Test modules are never part of the production import set.
- Fixtures are synthetic. Never commit real trades, counterparties, collateral
  agreements or portfolio extracts.
- Expected values come from a source independent of the implementation, such as
  published regulatory examples. Cases named `illustrative-*` hold only
  illustrative values: they show a behaviour and validate nothing.
- Run output, logs and generated workbooks are not committed.
- `python tools/check.py` validates every fixture and expected file against
  the format (`check_test_cases`) and checks that the generated
  `modules/TestCases.bas` is current. `TestCases.RunCaseTests` runs them in
  Excel through `modules/CaseRunner.bas`.

<a id="ported-prototype-catalogue"></a>

## Ported prototype catalogue

The 28 checks of the prototype workbook's TestCatalogue sheet (#59), with the
catalogue's expected value and tolerance. Inputs come from the template's
NettingSets and Trades sheets, all in EUR, valued on 2026-09-30. Netting sets
that the workbook duplicates per regime with identical inputs (`BCBS-EX5` and
`CRR-EX5`, `BCBS-VMCAP` and `CRR-VMCAP`, `BCBS-NEGIR` and `CRR-NEGIR`) are one
fixture with an expected file per regime. T22 and T26 read trades `NIR-B` and
`OT-X` in the workbook; the fixtures hold the same inputs as `NIR-C` and
`OT-1`. T28 needs only trade `B03` of `NS-GAMMA`.

| Test | Case and regime | Quantity (trade) | Expected | Class |
| --- | --- | --- | --- | --- |
| T01 | `cre99-example-1.bcbs` | `adjusted_notional` (EX1-T1) | 78,694 EUR | published |
| T02 | `cre99-example-1.bcbs` | `supervisory_delta` (EX1-T3) | -0.27 | published |
| T03 | `cre99-example-1.bcbs` | `add_on.interest_rate` | 347 EUR | published |
| T04 | `cre99-example-1.bcbs` | `exposure_value` | 569 EUR | published |
| T05 | `cre99-example-2.bcbs` | `multiplier` | 0.965 | published |
| T06 | `cre99-example-2.bcbs` | `exposure_value` | 381 EUR | published |
| T07 | `cre99-example-3.bcbs` | `exposure_value` | 5,406 EUR | published |
| T08 | `cre99-example-4.bcbs` | `replacement_cost` | 40 EUR | published |
| T09 | `cre99-example-4.bcbs` | `aggregate_add_on` | 629 EUR | published |
| T10 | `cre99-example-4.bcbs` | `exposure_value` | 936 EUR | published |
| T11 | `cre99-example-5.bcbs` | `margin_period_of_risk` | 14 business days | published |
| T12 | `cre99-example-5.bcbs` | `aggregate_add_on` | 1,401 EUR | published |
| T13 | `cre99-example-5.bcbs` | `multiplier` | 0.958 | published |
| T14 | `cre99-example-5.bcbs` | `exposure_value` | 1,879 EUR | published |
| T15 | `cre99-example-5.crr` | `exposure_value` | 1,879 EUR | illustrative |
| T16 | `illustrative-posted-vm-cap.bcbs` | `exposure_value` | 376592788.3471456 EUR | illustrative |
| T17 | `illustrative-posted-vm-cap.crr` | `exposure_value` | 40774849.34860364 EUR | illustrative |
| T18 | `illustrative-posted-vm-cap.crr` | `cap_applied` | `Y` | illustrative |
| T19 | `illustrative-swaption-forward-1pct.crr` | `exposure_value` | 760.4939328393385 EUR | illustrative |
| T20 | `illustrative-negative-ir-forward.crr` | `lambda_shift` (NIR-C) | 0.003 | illustrative |
| T21 | `illustrative-negative-ir-forward.crr` | `supervisory_delta` (NIR-C) | -0.9941752720429311 | illustrative |
| T22 | `illustrative-negative-ir-forward.bcbs` | `trade_status` (NIR-C) | `EXCLUDED` | illustrative |
| T23 | `illustrative-negative-commodity-price.crr` | `lambda_shift` (NCO-1) | 5.5 price | illustrative |
| T24 | `illustrative-other-risks.crr` | `add_on.other` | 880 EUR | illustrative |
| T25 | `illustrative-other-risks.crr` | `exposure_value` | 1,232 EUR | illustrative |
| T26 | `illustrative-other-risks.bcbs` | `trade_status` (OT-1) | `EXCLUDED` | illustrative |
| T27 | `illustrative-climatic-commodity.crr` | `hedging_set` (CL-1) | `CO\|CLIMATIC\|STD` | illustrative |
| T28 | `illustrative-credit-quality-step.crr` | `supervisory_factor` (B03) | 0.0042 | illustrative |

`published` values are printed in BCBS 279 Annex 4 (CRE99), as transcribed in
the catalogue, rounded to the printed digit; they have not yet been re-checked
against the registered text. `illustrative` values came from a Python
implementation that is not in the repository or from unreviewed hand
calculations, and stay illustrative until an independent derivation is
committed and reviewed. Every expected value and tolerance matches the
catalogue, and the prototype engine's saved results in the template are within
each tolerance.
