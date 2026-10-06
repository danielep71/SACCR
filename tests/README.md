# 🧪 Tests

`tests/` holds all verification source and stable test data. Layout rules are in
[`docs/REPOSITORY_STRUCTURE.md`](../docs/REPOSITORY_STRUCTURE.md).

| Location | Contents | Defined in |
| --- | --- | --- |
| `modules/` | Regression modules; `TestHarness.bas` is the harness and its entry point | #7 |
| `fixtures/` | Synthetic inputs: trades, netting sets, collateral terms | #8 |
| `expected/` | Reviewed expected values with their independent source | #8 |

The harness entry point is `TestHarness.RunTests`. How to run it and what a
passing log looks like is in
[`INSTALLATION.md`](../INSTALLATION.md#running-the-harness). Each other
subdirectory is created with its first real file.

Rules:

- Test modules are never part of the production import set.
- Fixtures are synthetic. Never commit real trades, counterparties, collateral
  agreements or portfolio extracts.
- Expected values come from a source independent of the implementation, such as
  published regulatory examples.
- Run output, logs and generated workbooks are not committed.
