# ✍️ VBA House Style

[![Scope: VBA exports](https://img.shields.io/badge/scope-VBA%20exports-217346)](REPOSITORY_STRUCTURE.md)
[![Contracts: explicit](https://img.shields.io/badge/contracts-explicit-6f42c1)](#procedure-contracts)
[![State: restored](https://img.shields.io/badge/Excel%20state-restored-0969da)](#operation-and-cleanup)

This document is authoritative for **how SACCR's VBA is written**: naming,
contracts, errors, Excel-state handling and source presentation. Where a
component lives is owned by
[`REPOSITORY_STRUCTURE.md`](REPOSITORY_STRUCTURE.md); importing and exporting
by [`INSTALLATION.md`](../INSTALLATION.md#importing-vba-into-excel). It adapts
`docs/VBA_HOUSE_STYLE.md` from EXCEL-VBA-PROJECT-TEMPLATE.

## 🏷️ Naming

The VBE lists all components in one flat project, so component names carry
their role.

| Element | Convention | Example |
| --- | --- | --- |
| Core module (`src/core/`) | `Core` + subject | `CoreReplacementCost` |
| Facade module (`src/modules/`) | `Saccr` + subject | `SaccrApi` |
| Class (`src/classes/`) | `C` + noun | `CNettingSet` |
| Workbook module (`src/workbook/`) | `ThisWorkbook`; sheets by code name `sh` + noun | `shInputs` |
| Test module (`tests/`) | `Test` + subject | `TestReplacementCost` |
| Example module (`examples/`) | `Example` + subject | `ExampleSingleNettingSet` |
| Procedure | PascalCase verb phrase | `CalculateAddOn` |
| Parameter, local variable | camelCase noun | `notionalAmount` |
| Constant | UPPER_SNAKE_CASE | `ALPHA_FACTOR` |
| Enum type / member | `Saccr` + noun / short prefix + PascalCase | `SaccrAssetClass` / `acInterestRate` |

Names are English and spelled out; avoid abbreviations other than established
SA-CCR terms (`EAD`, `RC`, `PFE`, `MPOR`), written in full in the procedure's
comment the first time.

## ⚙️ Module rules

- Every module declares `Option Explicit`.
- Every module in `src/core/` also declares `Option Private Module`. Both rules
  are enforced by `tools/check_source.py`.
- No `Option Base 1`: arrays are declared with explicit bounds
  (`Dim values(1 To n)`).
- No `Variant` where a specific type works. Arguments are `ByVal` unless the
  procedure's contract is to write an output parameter.
- No global mutable state in `src/core/`. Core procedures depend only on their
  arguments.
- A `Declare` statement, if ever needed, uses `PtrSafe` and `LongPtr`, because
  64-bit Excel is the primary target.

<a id="procedure-contracts"></a>

## 📜 Procedure contracts

Every procedure documents its contract in a banner immediately after its
signature. Omit sections that are empty.

| Section | Explain |
| --- | --- |
| `PURPOSE` | The responsibility the procedure owns |
| `INPUTS` | Meaning, **unit**, **valid domain** and ownership of each argument |
| `RETURNS` | Meaning and unit of the result and any output parameter |
| `ERROR POLICY` | Errors raised, propagated or contained, and cleanup |
| `DEPENDENCIES` | Boundaries that matter to callers or maintainers |
| `STATE OWNERSHIP` | Excel or module state read, changed or restored |
| `REFERENCE` | Regulatory paragraph or published example implemented |
| `UPDATED` | Date the procedure was last revised |

### Units and domains

- Amounts are `Double` in the netting set's calculation currency, never scaled
  (no thousands or millions).
- Rates, factors and correlations are decimals: `0.05`, never `5` meaning 5%.
- Time periods are `Double` in years unless the name says otherwise
  (`mporBusinessDays`). Dates are VBA `Date`.
- Every numeric input states its valid domain, for example `notionalAmount >= 0`
  or `0 < correlation <= 1`, and the facade rejects values outside it.
- The exact day-count, maturity and supervisory parameter conventions are part
  of the methodology (#8) and are cited in `REFERENCE`.

### Errors

- Error numbers are named constants in the range `vbObjectError + 2048` to
  `vbObjectError + 4095`, defined once in a core module and re-exposed by the
  facade when callers need them.
- Raise with `Err.Raise number, source, description`. The source names the
  procedure (`"SACCR.CalculateEad"`); the description states what was invalid
  and the offending value.
- The facade validates inputs and raises; the core may assume validated
  arguments but still raises on internal impossibilities.
- `On Error Resume Next` is allowed only around a single statement whose error
  is checked on the next line, followed by `On Error GoTo 0` or the procedure's
  handler.
- Errors are never swallowed silently. A procedure either handles an error
  completely or propagates it unchanged.

### Public API changes

The supported API is the list in [`PUBLIC_API.txt`](PUBLIC_API.txt). Adding,
removing or changing a listed declaration's name, arguments, defaults, units,
return value or errors is a contract change. It updates the manifest, the
regression tests and the `[Unreleased]` changelog entry in the same pull
request, and states the compatibility impact.

<a id="operation-and-cleanup"></a>

## 🧹 Operation and cleanup

Any procedure that changes Excel application state owns restoring it.

1. **Capture** the values it will change before changing them
   (`Calculation`, `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, `Cursor`,
   `StatusBar`).
2. **Change** only what the operation needs.
3. **Restore** in one cleanup block reached on both the success and the error
   path, restoring only what this procedure changed.
4. **Preserve the primary error.** Copy `Err.Number`, `Err.Source` and
   `Err.Description` into locals before cleanup runs, and re-raise those after
   cleanup. Restoring is best effort: a failure inside cleanup must never
   replace the original error.

```vb
Public Sub RunCalculation()
'==============================================================================
'                              RUN CALCULATION
'------------------------------------------------------------------------------
' PURPOSE         Recalculate all netting sets on the input sheet.
' STATE OWNERSHIP Changes ScreenUpdating and Calculation; restores both.
' ERROR POLICY    Restores state, then re-raises the original error unchanged.
'==============================================================================
    Dim savedScreen        As Boolean   'ScreenUpdating before this run
    Dim savedCalculation   As Long      'Calculation mode before this run
    Dim errNumber          As Long      'Primary error, preserved across cleanup
    Dim errSource          As String    'Primary error source
    Dim errDescription     As String    'Primary error description

        savedScreen = Application.ScreenUpdating
        savedCalculation = Application.Calculation
        On Error GoTo HandleError
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

        '... work ...

CleanUp:
        On Error Resume Next
        Application.Calculation = savedCalculation
        Application.ScreenUpdating = savedScreen
        On Error GoTo 0
        If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
        Exit Sub

HandleError:
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Resume CleanUp

End Sub
```

Core procedures never change Excel state, so they need none of this.

## 🧱 Module layout

1. `Attribute VB_Name` on the first line of a `.bas` export.
2. A full-width module banner: name, purpose, public surface, dependencies,
   state ownership, error policy, compatibility, updated date and author.
3. `MODULE SETTINGS`: `Option Explicit`, plus `Option Private Module` in core.
4. `MODULE CONSTANTS` and, outside core, `MODULE STATE` when needed.
5. Centered, uppercase procedure-group headings, then the procedures.

Banners use an apostrophe and 78 `=`; section rules an apostrophe and 78 `-`.
Source decorations are ASCII only.

## 📐 Indentation and wrapping

- Four-space levels; no tabs, no trailing whitespace.
- Local declarations at four spaces with `As` aligned in each group; body
  statements at eight spaces, nested blocks adding four.
- Procedure declarations, `End Sub` / `End Function`, labels and `#If`
  directives at column one.
- A signature with parameters puts one parameter per continuation line, and a
  function's return type on its own continuation line:

  ```vb
  Public Function CalculateAddOn( _
      ByVal notionalAmount As Double, _
      ByVal supervisoryFactor As Double) _
      As Double
  ```

- Two empty lines between procedures. Aim for 79-column comments and about 100
  columns of code.
- Comments explain intent and invariants, not assignments. Each declaration
  carries a short trailing comment stating its role or unit.

<a id="export-compatibility"></a>

## 💾 Export compatibility

- **Encoding:** the VBE exports in the Windows ANSI code page, which is cp1252
  on the supported Western-European hosts. Keep source ASCII; build any other
  character at run time with `ChrW`. `tools/check_source.py` rejects files that
  do not decode as cp1252. No byte-order mark.
- **Line endings:** stored with LF in Git, checked out with CRLF by
  `.gitattributes`, which is what the VBE expects. Never change these rules or
  the export header when reformatting.
- **Forms:** a `.frm` and its `.frx` are one component, kept side by side. The
  `.frx` is binary and is never edited as text.
- **Workbooks and other Office files** are build outputs, ignored by Git. One
  may be committed only at an exact path re-included in `.gitignore` and
  approved in an issue.

## 🔎 Review and validation

- A formatting-only change must leave names, signatures, visibility, constants,
  calculations, tests and error behavior unchanged; a discovered defect goes in
  a separate change.
- `python tools/check.py` must pass. It checks storage, naming, placement and
  `Option Explicit` / `Option Private Module`. It does not compile VBA.
- Changed VBA is merged only after it has been compiled and tested in Excel, as
  described in [`CONTRIBUTING.md`](../CONTRIBUTING.md#validation-and-evidence).

---

**Style principle:** state the contract, name the unit, restore what you
change, and never lose the original error.
