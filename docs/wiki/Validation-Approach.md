# Validation Approach

> **Guide, not policy:** [Tests](../../tests/README.md) and the [test-case methodology](../methodology/TEST_CASES.md) define what counts as validation evidence.

SA-CCR Benchmark is intended as an **independent challenger and reconciliation
calculator**, not merely a second copy of a bank's production implementation.

Validation evidence is built from:

1. published regulatory worked examples;
2. independently derived and reviewed expected results;
3. property, invariant and boundary tests;
4. independent external implementations used as comparators, not regulatory authority; and
5. reproducible Excel execution evidence tied to an exact source commit.

Illustrative cases demonstrate behaviour but do not count as independent
numerical validation.

Reconciliation should localise differences progressively from EAD to replacement
cost and PFE, multiplier and aggregate add-on, asset class, hedging set or
maturity bucket, and finally trade-level quantities such as adjusted notional,
supervisory delta, maturity factor and supervisory factor.

Static checks cannot prove that VBA compiles or that Excel produces the expected
numbers. Host execution evidence is maintained separately under
[Excel Evidence](../EXCEL_EVIDENCE.md).
