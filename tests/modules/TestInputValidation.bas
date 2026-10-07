Attribute VB_Name = "TestInputValidation"
'==============================================================================
' MODULE: TestInputValidation
'------------------------------------------------------------------------------
' PURPOSE
'   Check that malformed inputs are rejected instead of silently changing
'   the trade population or the netting-set terms (#35): trade rows with
'   data but no ID, duplicate trade IDs, a missing MtM, an unreadable date,
'   and netting-set fields that cannot be read. A blank optional field must
'   still take its default.
'
' PUBLIC SURFACE
'   RunInputValidationTests is the entry point. Option Private Module keeps
'   it out of the external workbook automation API.
'
' DEPENDENCIES
'   CaseRunner, which writes the inputs, runs the engine, restores the
'   workbook and reports. These inputs cannot be expressed as JSON
'   fixtures, so they are written here directly, as TEST_CASES.md allows
'   for invalid inputs.
'
' WORKSHEET SAFETY
'   As CaseRunner: inputs and outputs are rewritten during the run and
'   restored at the end. Use a development workbook.
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
        Private Const CASES       As Long = 9                 'Cases in a complete run
        Private Const CHECKS      As Long = 9                 'Checks in a complete run


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
