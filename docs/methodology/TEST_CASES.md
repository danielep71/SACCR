# 🧪 Numerical Test Cases

This document defines how SACCR's numerical test cases are written: inputs,
units, expected results, their provenance, tolerances, case categories and when
the suite counts as complete. Sources and rule IDs come from the
[methodology](README.md).

> [!IMPORTANT]
> The format is defined; no real SA-CCR case exists yet. The only committed case
> is **illustrative**: it shows the shape of the files and validates nothing.

## 📁 Files

A case is one input file and one expected-result file per regime:

| File | Holds |
| --- | --- |
| `tests/fixtures/<case-id>.json` | Synthetic inputs: netting set, trades, collateral |
| `tests/expected/<case-id>.<regime>.json` | Expected results for one regime (`crr` or `bcbs`), each with its provenance |

Separating them lets the same inputs carry a `CRR` and a `BCBS` result. Case IDs
are lowercase kebab-case and never reused. A case whose expected values are all
illustrative starts with `illustrative-`.

Files are UTF-8 JSON, stored with LF. All data is synthetic: never real trades,
counterparties, collateral agreements or portfolio extracts.

<a id="fixture-format"></a>

## 📥 Fixture format

```json
{
  "schema_version": 1,
  "kind": "saccr-fixture",
  "id": "illustrative-ir-swap-unmargined",
  "synthetic": true,
  "description": "What the case exercises, in one sentence",
  "category": "nominal",
  "calculation_currency": "EUR",
  "valuation_date": "2026-01-02",
  "netting_set": {
    "id": "NS1",
    "margined": false,
    "collateral_net_amount": 0
  },
  "trades": [
    {
      "id": "T1",
      "asset_class": "interest_rate",
      "currency": "EUR",
      "notional_amount": 10000000,
      "start_date": "2026-01-02",
      "end_date": "2031-01-02",
      "market_value_amount": 100000
    }
  ]
}
```

| Rule | Detail |
| --- | --- |
| Envelope | `schema_version`, `kind`, `id` (equal to the file name), `synthetic: true`, `description` and `category` are required |
| Units | In `netting_set` and `trades`, a field's suffix states its unit: `_amount` in `calculation_currency`, unscaled; `_date` as ISO `YYYY-MM-DD`; `_years` in years; `_rate` and `_factor` as decimals. A quantity there never lacks a unit suffix. Envelope fields such as `schema_version` are metadata, not quantities |
| Signs | Market values are from the bank's side: positive is an asset. `collateral_net_amount` is collateral held minus collateral posted |
| Vocabulary | Trade and netting-set fields beyond the envelope are **provisional** and are fixed with the first engine milestone; adding one updates this table |

<a id="expected-format"></a>

## 📤 Expected-result format

```json
{
  "schema_version": 1,
  "kind": "saccr-expected",
  "fixture": "illustrative-ir-swap-unmargined",
  "regime": "CRR",
  "outputs": [
    {
      "quantity": "replacement_cost",
      "rule": "CRR.275.1",
      "value": 100000,
      "unit": "EUR",
      "tolerance": {"absolute": 0.000001, "relative": 0.000000001},
      "reference": {
        "class": "illustrative",
        "source": "CRR",
        "locator": "Article 275(1)",
        "derivation": "How the value was obtained",
        "derived_by": "Name",
        "reviewed_by": null,
        "date": "2026-10-06"
      }
    }
  ]
}
```

An invalid-input case replaces `value`, `unit` and `tolerance` with the error
the facade must raise, by its public constant name. The name below is an
example; no such constant exists yet:

```json
{"quantity": "exposure_value", "rule": "CRR.274.2", "error": "SACCR_ERROR_INVALID_NOTIONAL", "reference": {"class": "illustrative", "...": "..."}}
```

Quantities use stable names: `exposure_value`, `replacement_cost`,
`potential_future_exposure`, `multiplier`, `aggregate_add_on`, and
`add_on.<asset_class>` for each asset class. Intermediate quantities may be
added when a rule needs them; each name is used with one meaning only.

<a id="reference-classes"></a>

## 🏷️ Reference classes

Every expected value carries exactly one class. Only the first two are
evidence that the implementation is correct.

| Class | Meaning | Required fields |
| --- | --- | --- |
| `published` | Printed in a source from the register, for example a worked example | `source`, exact `locator` (section, example, table) |
| `independent` | Computed outside SACCR's code, by a separate method such as a reviewer's spreadsheet or script, from the cited rules | `source`, `locator`, `derivation`, `derived_by`, `reviewed_by` |
| `illustrative` | Shows the format or a behaviour; not a validated SA-CCR figure | `derivation` saying why it exists |

An `independent` value is never produced by running SACCR itself, and never
reviewed only by the person who wrote the code under test. Test output reports
illustrative results separately and never as validation.

<a id="tolerance"></a>

## 📏 Tolerance

An actual result `a` matches the expected value `e` when

```text
|a − e| ≤ max(absolute, relative × |e|)
```

Both bounds are stated explicitly on every output; there is no hidden default.
Use these conventions when choosing them:

| Reference | `absolute` | `relative` |
| --- | --- | --- |
| `independent` or `illustrative`, full precision | `1e-6` currency units (or `1e-9` for factors) | `1e-9` |
| `published`, rounded | Half a unit of the last printed digit: printed in thousands → `500` | `0` |

Near zero, the absolute bound governs: when `|e|` is small the relative term
vanishes, so an expected `0` is matched within `absolute` only. Domain rules
apply before tolerance: a quantity that must be non-negative, such as
`exposure_value`, fails if it is negative, however small.

<a id="categories"></a>

## 🗂️ Case categories

Each fixture declares one `category`:

| Category | Covers |
| --- | --- |
| `nominal` | Typical inputs within the domain |
| `boundary` | Values on a domain edge: zero notional or market value, results that should be exactly zero (fully offsetting trades, negative net value floored at zero), maturity at a regulatory floor or cap, a single trade, a single hedging set |
| `invalid` | Inputs the facade must reject: missing or unknown fields, negative amounts where not allowed, end date before start date, unknown asset class or currency. Expected results are errors, not values. JSON cannot hold NaN or infinity, so non-finite inputs are tested directly in a VBA test module instead of a fixture |
| `regime` | Inputs chosen to show a recorded [regime difference](README.md#regime-differences), with a `crr` and a `bcbs` expected file |

<a id="completeness-policy"></a>

## ✅ Completeness policy

- A rule in the [traceability table](README.md#traceability) is **validated**
  for a regime only when at least one `published` or `independent` case passes
  for it in that regime.
- Each implemented rule also has at least one `boundary` case, and each public
  input domain has at least one `invalid` case, as a fixture or, for inputs
  JSON cannot express, as a direct test in a VBA test module.
- Every recorded regime difference has a `regime` case.
- A release states which rules are validated, for which regime, and lists
  every rule that is implemented but not validated. Illustrative cases never
  count toward any of this.
- The harness keeps its own completeness check: a run that does not execute
  every expected case and assertion is incomplete, not a pass.

<a id="consumption"></a>

## ⚙️ Consumption

How the VBA harness reads these files is still open
([decision 7](README.md#open-decisions)): a JSON parser in VBA, or VBA test
modules generated from the JSON by a checked tool. A validator for the file
format is added with the first real case, so the format is enforced from then
on. Until then these rules are applied in review.
