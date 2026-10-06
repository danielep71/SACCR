Attribute VB_Name = "TestMainState"
'==============================================================================
' MODULE: TestMainState
'------------------------------------------------------------------------------
' PURPOSE
'   Check that the workbook macros in M_Main put Excel back exactly as they
'   found it, on success and on failure, report operation and cleanup
'   failures separately, and leave the workbook ready for the next run
'   (issue #43).
'
' PUBLIC SURFACE
'   RunMainStateTests is the entry point. Option Private Module keeps it out
'   of the external workbook automation API.
'
' DEPENDENCIES
'   M_Main, its test seam gTestFault, and the error numbers in M_Config. The
'   workbook must be built from the template with the full source, because
'   the macros calculate and write the output sheets.
'
' WORKSHEET SAFETY
'   Unlike TestHarness, this module runs the real macros: they clear and
'   rewrite TradeCalc, Buckets, HedgingSets, Results and Checks, and
'   activate the Checks sheet. The last case leaves the outputs of a normal
'   run. Use a development workbook only.
'
' STATE OWNERSHIP
'   Sets calculation mode, events and screen updating to known values for
'   each case and restores the caller's values, gSilent and gTestFault at
'   the end, also after an unexpected error.
'
' ERROR POLICY
'   Expected errors are captured and asserted. An unexpected error is
'   recorded against its case and the remaining cases still run.
'
' USAGE
'   Import with the production modules, compile, then run
'   TestMainState.RunMainStateTests from the Immediate window.
'
' UPDATED
'   2026-10-06
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the tests out of the external API.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
        Private Const EXPECTED_CASES   As Long = 7    'Cases in a complete run

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mCaseCount      As Long      'Cases started
        Private mCheckCount     As Long      'Checks evaluated
        Private mFailureCount   As Long      'Failed checks and unexpected errors
        Private mCurrentCase    As String    'Case being run, for messages


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunMainStateTests()
'
'==============================================================================
'                              RunMainStateTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case and print one CASE line per case, each failed check,
'   and a final RESULT line.
'
' STATE OWNERSHIP
'   Restores the caller's calculation mode, events, screen updating,
'   gSilent and gTestFault before printing the result.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim callerCalculation   As XlCalculation    'Caller's calculation mode
    Dim callerEvents        As Boolean          'Caller's EnableEvents
    Dim callerScreen        As Boolean          'Caller's ScreenUpdating
    Dim callerSilent        As Boolean          'Caller's M_Main.gSilent
    Dim restored            As Boolean          'Caller's settings read back equal

'------------------------------------------------------------------------------
' START
'------------------------------------------------------------------------------
        mCaseCount = 0
        mCheckCount = 0
        mFailureCount = 0
        callerCalculation = Application.Calculation
        callerEvents = Application.EnableEvents
        callerScreen = Application.ScreenUpdating
        callerSilent = M_Main.gSilent
        Debug.Print "SACCR MAIN STATE TESTS"

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Silent mode keeps message boxes from blocking the run and makes the
    'macros raise their errors.
        M_Main.gSilent = True
        M_Main.gTestFault = ""
        CaseRestoresState "clear.automatic-on", "clear", xlCalculationAutomatic, True, True
        CaseRestoresState "clear.manual-off", "clear", xlCalculationManual, False, False
        CaseRestoresState "validate.manual-off", "validate", xlCalculationManual, False, False
        CaseRestoresState "run.manual-off", "run", xlCalculationManual, False, False
        CaseOperationFailure
        CaseCleanupFailure
        CaseSilentFlagPreserved

'------------------------------------------------------------------------------
' RESTORE THE CALLER AND REPORT
'------------------------------------------------------------------------------
        M_Main.gTestFault = ""
        M_Main.gSilent = callerSilent
        Application.Calculation = callerCalculation
        Application.EnableEvents = callerEvents
        Application.ScreenUpdating = callerScreen
        restored = (Application.Calculation = callerCalculation) And _
                   (Application.EnableEvents = callerEvents) And _
                   (Application.ScreenUpdating = callerScreen)
        If mCaseCount <> EXPECTED_CASES Then
            Fail "suite", "expected " & EXPECTED_CASES & " cases, ran " & mCaseCount
        End If
        Debug.Print "CASES=" & mCaseCount & "; CHECKS=" & mCheckCount & "; FAILURES=" & mFailureCount
        Debug.Print "RESULT=" & IIf(mFailureCount = 0 And restored, "PASS", "FAIL") & _
                    "; cases=" & mCaseCount & "; checks=" & mCheckCount & _
                    "; failures=" & mFailureCount & "; caller_state=" & IIf(restored, "RESTORED", "CHANGED")

End Sub


'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseRestoresState( _
    ByVal caseName As String, _
    ByVal operation As String, _
    ByVal calcMode As XlCalculation, _
    ByVal eventsOn As Boolean, _
    ByVal screenOn As Boolean)
'
'==============================================================================
'                              CaseRestoresState
'------------------------------------------------------------------------------
' PURPOSE
'   Set known settings, run one operation, and check that every setting is
'   back to its value, including settings that were off before the run.
'
' INPUTS
'   caseName: case identifier.
'   operation: "clear", "validate" or "run".
'   calcMode, eventsOn, screenOn: the settings before the operation.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim result   As String    'Result line of a run

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase caseName
        SetState calcMode, eventsOn, screenOn
        Select Case operation
            Case "clear"
                M_Main.ClearOutputs
            Case "validate"
                M_Main.ValidateInputs
            Case "run"
                result = M_Main.RunSACCR_Silent()
                Check Left$(result, 9) = "RESULT=OK", "run result: " & result
        End Select
        CheckState calcMode, eventsOn, screenOn
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail caseName, "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseOperationFailure()
'
'==============================================================================
'                             CaseOperationFailure
'------------------------------------------------------------------------------
' PURPOSE
'   An error during the operation is raised unchanged after Excel has been
'   restored, and the next run works.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim result      As String    'Result line of the follow-up run
    Dim errNumber   As Long      'Error raised by the failed run
    Dim errText     As String    'Its description

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "run.operation-failure"
        SetState xlCalculationAutomatic, False, False
        M_Main.gTestFault = "operation"
        On Error Resume Next
        result = M_Main.RunSACCR_Silent()
        errNumber = Err.Number
        errText = Err.Description
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check errNumber = ERR_INJECTED_FAULT, "expected the injected error, got " & errNumber & ": " & errText
        CheckState xlCalculationAutomatic, False, False

    'The re-entry flag must have been cleared: the next run completes.
        result = M_Main.RunSACCR_Silent()
        Check Left$(result, 9) = "RESULT=OK", "follow-up run: " & result
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail "run.operation-failure", "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseCleanupFailure()
'
'==============================================================================
'                              CaseCleanupFailure
'------------------------------------------------------------------------------
' PURPOSE
'   A failed restoration is reported as ERR_CLEANUP_FAILED, the other
'   settings are still restored, and the next run works.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim result      As String    'Result line of the follow-up run
    Dim errNumber   As Long      'Error raised by the failed run
    Dim errText     As String    'Its description

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The injected fault stops the calculation-mode restoration, so the mode
    'stays manual; events and screen updating must still come back False.
        On Error GoTo Unexpected
        BeginCase "run.cleanup-failure"
        SetState xlCalculationAutomatic, False, False
        M_Main.gTestFault = "cleanup"
        On Error Resume Next
        result = M_Main.RunSACCR_Silent()
        errNumber = Err.Number
        errText = Err.Description
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check errNumber = ERR_CLEANUP_FAILED, "expected the cleanup error, got " & errNumber & ": " & errText
        Check InStr(errText, "Calculation:") > 0, "cleanup detail names calculation: " & errText
        Check InStr(errText, "RESULT=OK") > 0, "cleanup error keeps the run result: " & errText
        CheckState xlCalculationManual, False, False

    'After the user restores the mode by hand, the next run works.
        Application.Calculation = xlCalculationAutomatic
        result = M_Main.RunSACCR_Silent()
        Check Left$(result, 9) = "RESULT=OK", "follow-up run: " & result
        CheckState xlCalculationAutomatic, False, False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        Fail "run.cleanup-failure", "unexpected error " & Err.Number & ": " & Err.Description

End Sub


Private Sub CaseSilentFlagPreserved()
'
'==============================================================================
'                           CaseSilentFlagPreserved
'------------------------------------------------------------------------------
' PURPOSE
'   RunSACCR_Silent restores gSilent to its previous value, False or True,
'   on success and when the run raises.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim result   As String    'Result line

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo Unexpected
        BeginCase "silent.flag-preserved"
        SetState xlCalculationAutomatic, True, True
        M_Main.gSilent = False
        result = M_Main.RunSACCR_Silent()
        Check M_Main.gSilent = False, "gSilent was False before a successful run"
        M_Main.gSilent = True
        M_Main.gTestFault = "operation"
        On Error Resume Next
        result = M_Main.RunSACCR_Silent()
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check M_Main.gSilent = True, "gSilent was True before a failed run"
        M_Main.gSilent = False
        M_Main.gTestFault = "operation"
        On Error Resume Next
        result = M_Main.RunSACCR_Silent()
        On Error GoTo Unexpected
        M_Main.gTestFault = ""
        Check M_Main.gSilent = False, "gSilent was False before a failed run"
        M_Main.gSilent = True
        CheckState xlCalculationAutomatic, True, True
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
Unexpected:
        M_Main.gTestFault = ""
        M_Main.gSilent = True
        Fail "silent.flag-preserved", "unexpected error " & Err.Number & ": " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub BeginCase( _
    ByVal caseName As String)
'
'==============================================================================
'                                  BeginCase
'------------------------------------------------------------------------------
' PURPOSE
'   Count a case and print its CASE line.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        mCaseCount = mCaseCount + 1
        mCurrentCase = caseName
        Debug.Print "CASE=" & caseName

End Sub


Private Sub SetState( _
    ByVal calcMode As XlCalculation, _
    ByVal eventsOn As Boolean, _
    ByVal screenOn As Boolean)
'
'==============================================================================
'                                   SetState
'------------------------------------------------------------------------------
' PURPOSE
'   Put the three settings into a known state before an operation.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        Application.Calculation = calcMode
        Application.EnableEvents = eventsOn
        Application.ScreenUpdating = screenOn

End Sub


Private Sub CheckState( _
    ByVal calcMode As XlCalculation, _
    ByVal eventsOn As Boolean, _
    ByVal screenOn As Boolean)
'
'==============================================================================
'                                  CheckState
'------------------------------------------------------------------------------
' PURPOSE
'   Check each of the three settings against its expected value.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        Check Application.Calculation = calcMode, _
              "Calculation expected " & calcMode & ", found " & Application.Calculation
        Check Application.EnableEvents = eventsOn, _
              "EnableEvents expected " & eventsOn & ", found " & Application.EnableEvents
        Check Application.ScreenUpdating = screenOn, _
              "ScreenUpdating expected " & screenOn & ", found " & Application.ScreenUpdating

End Sub


Private Sub Check( _
    ByVal passed As Boolean, _
    ByVal detail As String)
'
'==============================================================================
'                                    Check
'------------------------------------------------------------------------------
' PURPOSE
'   Count one check and record it when it fails.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        mCheckCount = mCheckCount + 1
        If Not passed Then
            Fail mCurrentCase, detail
        End If

End Sub


Private Sub Fail( _
    ByVal caseName As String, _
    ByVal detail As String)
'
'==============================================================================
'                                     Fail
'------------------------------------------------------------------------------
' PURPOSE
'   Count a failure and print it.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
        mFailureCount = mFailureCount + 1
        Debug.Print "FAILURE=" & caseName & ": " & detail

End Sub
