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

Current source:

- the prototype engine imported from `SACCR_Calculator.xlsm`: `core/M_Config`,
  `core/M_Util`, `core/M_Engine`, the worksheet functions in
  `modules/M_Formulas`, and `workbook/M_Main` with the 13 document modules;
- the setup scaffold the regression harness tests: `core/CoreScaffold.bas` and
  `modules/SaccrScaffold.bas`.

The imported engine does not yet follow every rule here: `M_Engine` and
`M_Util` read and write worksheets from `src/core`, and `M_Main` is a standard
module kept with the workbook glue because the sheet buttons call it. See
[known deviations](../docs/REPOSITORY_STRUCTURE.md#known-deviations).

Do not place tests, examples, workbooks or generated output here.
