# 🧩 Production Source

`src/` holds the production VBA source: the only input from which the SACCR
workbook is built. Layout, visibility and dependency rules are defined in
[`docs/REPOSITORY_STRUCTURE.md`](../docs/REPOSITORY_STRUCTURE.md).

| Location | Contents |
| --- | --- |
| `core/` | Calculation engine. Every module declares `Option Private Module`; no Excel object-model access. |
| `modules/` | Public facade: supported entry points listed in [`docs/PUBLIC_API.txt`](../docs/PUBLIC_API.txt). |
| `classes/` | Class modules; each header states whether it is public surface or internal. |
| `workbook/` | Exported `ThisWorkbook` and sheet modules; host glue only. |
| `forms/` | UserForms with `.frm` beside `.frx`, only if one is ever needed. |

The only source so far is the setup scaffold that the regression harness tests:
`core/CoreScaffold.bas` and `modules/SaccrScaffold.bas`. It is replaced by the
SA-CCR engine. Each other subdirectory is created with its first real
component.

Do not place tests, examples, workbooks or generated output here.
