Attribute VB_Name = "TestInputValidation"
'==============================================================================
' MODULE: TestInputValidation
'------------------------------------------------------------------------------
' PURPOSE
'   Check that malformed inputs are rejected instead of silently changing
'   the trade population, the netting-set terms or the parameters (#35):
'   trade rows with data but no ID, duplicate trade IDs, a missing MtM, an
'   unreadable date, netting-set fields that cannot be read, a moved or
'   renamed column, a broken parameter name, a parameter, supervisory
'   factor or currency listed twice with different values, and a factor or
'   rate the table loader would not read because of a blank row or key. A
'   blank optional field, and a duplicate with the same values, must still
'   be accepted.
'
' PUBLIC SURFACE
'   RunInputValidationTests is the entry point. Option Private Module keeps
'   it out of the external workbook automation API.
'
' DEPENDENCIES
'   CaseRunner, which writes the inputs, runs the engine, restores the
'   workbook and reports. These inputs cannot be expressed as JSON
'   fixtures, so they are written here directly, as TEST_CASES.md allows
'   for invalid inputs. The table cases find their rows on Params by key
'   (CO_OTHER, OT, USD, JPY, CHF), so they need the template's Params
'   sheet.
'
' WORKSHEET SAFETY
'   As CaseRunner: inputs, outputs, and the headers, Params cells and
'   workbook name a case patches are rewritten during the run and
'   restored. Use a development workbook.
'
' USAGE
'   Run TestInputValidation.RunInputValidationTests from the Immediate
'   window.
'
' UPDATED
'   2026-10-07
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
        Private Const VALUATION   As String = "2026-09-30"    'Valuation date of every case
        Private Const CASES       As Long = 23                'Cases in a complete run
        Private Const CHECKS      As Long = 23                'Checks in a complete run


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunInputValidationTests()
'
'==============================================================================
'                           RunInputValidationTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case and print the CaseRunner report.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Params row a table case patches

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
        CaseRunner.BeginSuite CASES, CHECKS
        On Error GoTo Failed

    'Trade rows.
        StartCase "trade-middle-row-without-id", "N"
        AddSwap "T1"
        AddSwap "T2"
        AddSwap "T3"
        CaseRunner.SetInputCell SH_TRADES, 2, TR_ID, Empty
        ExpectStatus "INCOMPLETE: 1 of 3 trade(s) rejected"

        StartCase "trade-trailing-row-without-id", "N"
        AddSwap "T1"
        AddSwap "T2"
        CaseRunner.SetInputCell SH_TRADES, 2, TR_ID, Empty
        ExpectStatus "INCOMPLETE: 1 of 2 trade(s) rejected"

        StartCase "trade-duplicate-id", "N"
        AddSwap "T1"
        AddSwap "T1"
        ExpectStatus "INCOMPLETE: 1 of 2 trade(s) rejected"

        StartCase "trade-missing-mtm", "N"
        AddSwap "T1"
        CaseRunner.SetInputCell SH_TRADES, 1, TR_MTM, Empty
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"

        StartCase "trade-unreadable-date", "N"
        AddSwap "T1"
        CaseRunner.SetInputCell SH_TRADES, 1, TR_MAT, "next year"
        ExpectStatus "INCOMPLETE: 1 of 1 trade(s) rejected"

    'Netting-set fields.
        StartCase "netting-set-unrecognised-flag", "N"
        CaseRunner.SetInputCell SH_NS, 1, NS_MARGINED, "Maybe"
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "netting-set-non-numeric-amount", "N"
        CaseRunner.SetInputCell SH_NS, 1, NS_NICA, "abc"
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "netting-set-unknown-regime", "N"
        CaseRunner.SetInputCell SH_NS, 1, NS_REGIME, "EU"
        AddSwap "T1"
        ExpectStatus "INVALID: 1 input error(s)"

        StartCase "netting-set-blank-flags-take-defaults", ""
        AddSwap "T1"
        ExpectStatus "VALID"

    'Sheet layout: a header that differs means a column moved.
        StartCase "layout-trades-header-renamed", "N"
        AddSwap "T1"
        CaseRunner.PatchCell SH_TRADES, HEADER_ROW, TR_NOTIONAL, "Nominal"
        ExpectStop "Header must be 'Notional'"

        StartCase "layout-netting-sets-column-shifted", "N"
        AddSwap "T1"
        CaseRunner.PatchCell SH_NS, HEADER_ROW, NS_NICA, "Threshold TH"
        ExpectStop "Header must be 'NICA'"

        StartCase "layout-params-value-header-renamed", "N"
        AddSwap "T1"
        CaseRunner.PatchCell SH_PARAMS, HEADER_ROW, PRM_VALUE_COL, "Amount"
        ExpectStop "Header must be 'Value'"

        StartCase "layout-broken-parameter-name", "N"
        AddSwap "T1"
        CaseRunner.PatchName PRM_ALPHA, "=#REF!"
        ExpectStop "no longer refers to a cell"

    'Duplicates: the same values are accepted, different values stop the run.
        StartCase "params-duplicate-different-value", "N"
        AddSwap "T1"
        AddParamRow PRM_ALPHA, 1.2
        ExpectStop "Parameter listed twice with different values"

        StartCase "params-duplicate-same-value", "N"
        AddSwap "T1"
        AddParamRow PRM_ALPHA, ParamsValue(PRM_ALPHA, PRM_VALUE_COL)
        ExpectStatus "VALID"

    'CO_OTHER becomes a second CO_METALS row: it differs in category and
    'hedging set, unless those are patched too.
        StartCase "factor-table-duplicate-different-values", "N"
        AddSwap "T1"
        r = ParamsRow("CO_OTHER")
        CaseRunner.PatchCell SH_PARAMS, r, 1, "CO_METALS"
        ExpectStop "Supervisory factor key listed twice with different values"

        StartCase "factor-table-duplicate-same-values", "N"
        AddSwap "T1"
        r = ParamsRow("CO_OTHER")
        CaseRunner.PatchCell SH_PARAMS, r, 1, "CO_METALS"
        CaseRunner.PatchCell SH_PARAMS, r, 3, "METALS"
        CaseRunner.PatchCell SH_PARAMS, r, 7, "METALS"
        ExpectStatus "VALID"

    'CHF becomes a second USD row, at CHF's rate or at USD's.
        StartCase "fx-table-duplicate-different-rates", "N"
        AddSwap "T1"
        r = ParamsRow("CHF")
        CaseRunner.PatchCell SH_PARAMS, r, 1, "USD"
        ExpectStop "Currency listed twice with different rates"

        StartCase "fx-table-duplicate-same-rate", "N"
        AddSwap "T1"
        r = ParamsRow("CHF")
        CaseRunner.PatchCell SH_PARAMS, r, 1, "USD"
        CaseRunner.PatchCell SH_PARAMS, r, 3, ParamsValue("USD", 3)
        ExpectStatus "VALID"

    'Gaps: a table ends at its first blank key, so a row below a blank row,
    'or a value without a key, would not be read.
        StartCase "factor-table-blank-row-inside", "N"
        AddSwap "T1"
        ClearParamsRow ParamsRow("CO_OTHER"), 8
        ExpectStop "below a blank row"

        StartCase "factor-table-value-without-key", "N"
        AddSwap "T1"
        CaseRunner.PatchCell SH_PARAMS, ParamsRow("OT"), 1, Empty
        ExpectStop "has a value but no key"

        StartCase "fx-table-blank-row-inside", "N"
        AddSwap "T1"
        ClearParamsRow ParamsRow("JPY"), 3
        ExpectStop "below a blank row"

        StartCase "fx-table-rate-without-currency", "N"
        AddSwap "T1"
        CaseRunner.PatchCell SH_PARAMS, ParamsRow("CHF"), 1, Empty
        ExpectStop "has a value but no key"

        CaseRunner.EndSuite ""
        Exit Sub

Failed:
        CaseRunner.EndSuite "unexpected error " & Err.Number & ": " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub StartCase( _
    ByVal caseId As String, _
    ByVal flagValue As String)
'
'==============================================================================
'                                  StartCase
'------------------------------------------------------------------------------
' PURPOSE
'   Start a CRR case with one unmargined netting set NS1. flagValue is
'   written to every Y/N field: "N", or "" to leave them blank.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        CaseRunner.BeginCase "input-" & caseId, "CRR", VALUATION, "EUR"
        CaseRunner.AddNettingSet "NS1", flagValue, flagValue, "", flagValue, flagValue, "", _
                                 "0", "0", "0", "0", ""

End Sub


Private Sub AddSwap( _
    ByVal tradeId As String)
'
'==============================================================================
'                                   AddSwap
'------------------------------------------------------------------------------
' PURPOSE
'   Add a valid five-year EUR interest-rate swap to NS1.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        CaseRunner.AddTrade tradeId, "IR", "", "EUR", "Linear", "Long", "", "Standard", "", _
                            "10000", "30", "", "2031-09-30", "2031-09-30", "", "", "", "", "", ""

End Sub


Private Sub ExpectStop( _
    ByVal messagePart As String)
'
'==============================================================================
'                                  ExpectStop
'------------------------------------------------------------------------------
' PURPOSE
'   Run the case and check that the run stopped with an ERROR on Checks
'   containing messagePart.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        CaseRunner.RunCaseExpectingStop "stop", messagePart, "illustrative"

End Sub


Private Function ParamsRow( _
    ByVal key As String) _
    As Long
'
'==============================================================================
'                                  ParamsRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the Params row whose column A holds key: a parameter code, a
'   supervisory-factor key or a currency.
'
' ERROR POLICY
'   Raises ERR_TEST_SETUP when the key is not on Params, so that a changed
'   template fails the suite instead of patching the wrong row.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Row found; 0 when absent

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        r = FindHeaderRow(GetSheet(SH_PARAMS), 1, key)
        If r = 0 Then
            Err.Raise ERR_TEST_SETUP, "TestInputValidation.ParamsRow", "'" & key & "' not found in column A of Params."
        End If
        ParamsRow = r

End Function


Private Function ParamsValue( _
    ByVal key As String, _
    ByVal col As Long) _
    As Variant
'
'==============================================================================
'                                 ParamsValue
'------------------------------------------------------------------------------
' PURPOSE
'   Return a value from the Params row whose column A holds key.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        ParamsValue = GetSheet(SH_PARAMS).Cells(ParamsRow(key), col).Value

End Function


Private Sub AddParamRow( _
    ByVal code As String, _
    ByVal newValue As Variant)
'
'==============================================================================
'                                 AddParamRow
'------------------------------------------------------------------------------
' PURPOSE
'   Add a second Params row for a parameter, two rows below the last used
'   row so that no table on Params runs into it.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'The new row

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        r = UsedLastRow(GetSheet(SH_PARAMS)) + 2
        CaseRunner.PatchCell SH_PARAMS, r, PRM_CODE_COL, code
        CaseRunner.PatchCell SH_PARAMS, r, PRM_VALUE_COL, newValue

End Sub


Private Sub ClearParamsRow( _
    ByVal rowNum As Long, _
    ByVal nCols As Long)
'
'==============================================================================
'                                ClearParamsRow
'------------------------------------------------------------------------------
' PURPOSE
'   Blank columns A to nCols of one Params row, as if a blank row had been
'   inserted in a table there. Restored with the case's other patches.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim c   As Long    'Column being cleared

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        For c = 1 To nCols
            CaseRunner.PatchCell SH_PARAMS, rowNum, c, Empty
        Next c

End Sub


Private Sub ExpectStatus( _
    ByVal expected As String)
'
'==============================================================================
'                                 ExpectStatus
'------------------------------------------------------------------------------
' PURPOSE
'   Run the case and check the netting set's status on Results.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        CaseRunner.RunCase
        CaseRunner.ExpectText "status", "", "netting_set_status", expected, "illustrative"

End Sub
