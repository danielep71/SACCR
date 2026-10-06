<!--
Target branch: the active release branch (release/1.0.0), never main unless the
owner asked for an integration. Synthetic data only. Delete guidance comments.
-->

## Summary

<!-- What changed and why, in two or three sentences. -->

Closes / part of #

## Change type

- [ ] Calculation (src/core)
- [ ] Public API (src/modules, docs/PUBLIC_API.txt)
- [ ] Workbook or user interface
- [ ] Tests or fixtures
- [ ] Tooling or CI
- [ ] Documentation only

## Compatibility

<!-- Effect on public declarations, computed results, input/output layouts or
supported hosts. Write "None" if there is none. -->

## Validation

| Check | Result |
| --- | --- |
| `python tools/check.py --ci` | <!-- PASS / FAIL --> |
| Repository integrity (CI) | <!-- green / red --> |
| Excel compile (Debug → Compile VBAProject) | <!-- PASS / Not run: no VBA changed --> |
| Regression harness | <!-- passed/run, Failed=0 / Not run --> |
| Specific scenario | <!-- what was exercised / Not run --> |
| Excel version, bitness, Windows | <!-- e.g. M365 2409 64-bit, Windows 11 --> |

<!-- Static checks are never Excel evidence. A PR that changes VBA is merged
only after the owner has verified it in Excel. State plainly what was not run. -->

## Checklist

- [ ] The issue is assigned to danielep71, has a P1/P2/P3 label and the milestone.
- [ ] The `[Unreleased]` changelog entry is updated, or the change is not user-facing.
- [ ] Documentation matches the new behavior.
- [ ] No real trades, counterparties, credentials, private links or workbooks are added.
