Attribute VB_Name = "M_Main"
'==============================================================================
' MODULE: M_Main
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the macros behind the buttons on the sheets, and report each run
'   to the user.
'
' PUBLIC SURFACE
'   RunSACCR: validate the inputs, calculate and write every output sheet.
'   ValidateInputs: validate only; findings go to the Checks sheet.
'   ClearOutputs: empty every output sheet.
'   RunSACCR_Silent: RunSACCR without message boxes, for automation.
'   The sheet buttons call the first three by name. These macros are not
'   listed in docs/PUBLIC_API.txt.
'
' DEPENDENCIES
'   M_Engine for the calculation and its run counters; M_Util and M_Config
'   for the output sheets.
'
' STATE OWNERSHIP
'   Owns gSilent, which switches messages from message boxes to the
'   Immediate window. Each macro switches off screen updating and events and
'   sets manual calculation for the run (BeginBatch), then sets them back
'   (EndBatch).
'
' ERROR POLICY
'   Each macro catches any unexpected error, sets Excel back with EndBatch and
'   shows the error number and description. Errors in the inputs are not
'   VBA errors: the engine lists them on the Checks sheet.
'
' KNOWN DEVIATION
'   EndBatch switches screen updating and events back on rather than
'   restoring the values captured before the run, and this standard module
'   lives in src/workbook. See the known deviations in
'   docs/REPOSITORY_STRUCTURE.md.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
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
    'Require explicit declarations. The macros must stay visible to the
    'sheet buttons, so this module is not private.
    Option Explicit

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'True only while RunSACCR_Silent runs: messages go to the Immediate
    'window instead of message boxes.
        Public gSilent   As Boolean    'Silent mode for automation


'
'------------------------------------------------------------------------------
'
'                                BUTTON MACROS
'
'------------------------------------------------------------------------------
'

Public Sub RunSACCR()
'
'==============================================================================
'                                   RunSACCR
'------------------------------------------------------------------------------
' PURPOSE
'   Run the full SA-CCR calculation and write every output sheet, then
'   report the trades used, the total EAD and the number of errors and
'   warnings.
'
' USAGE
'   Assigned to the Run SA-CCR button.
'
' STATE OWNERSHIP
'   Calculation, screen updating and events are changed for the run by
'   BeginBatch and set back by EndBatch.
'
' ERROR POLICY
'   An unexpected error is reported with its number and description after
'   EndBatch. A run stopped by invalid inputs activates the Checks sheet.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ok         As Boolean    'True when the engine completed the run
    Dim calcMode   As Long       'Calculation mode captured by BeginBatch
    Dim msg        As String     'Summary shown to the user

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        On Error GoTo Fail
        BeginBatch calcMode
        ok = M_Engine.Calculate(True)
        EndBatch calcMode

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'A completed run can still have excluded trades; point to the Checks
    'sheet whenever there are errors or warnings.
        If ok Then
            msg = "SA-CCR run completed." & vbCrLf & vbCrLf & _
                  "Trades used: " & M_Engine.TradesUsed & " of " & M_Engine.TradesRead & vbCrLf & _
                  "Total EAD: " & Format$(M_Engine.TotalEAD, "#,##0") & vbCrLf & _
                  "Errors: " & M_Engine.ErrorCount & "   Warnings: " & M_Engine.WarningCount
            If M_Engine.ErrorCount > 0 Or M_Engine.WarningCount > 0 Then
                msg = msg & vbCrLf & vbCrLf & "See the Checks sheet for details."
            End If
            Notify msg, IIf(M_Engine.ErrorCount > 0, vbExclamation, vbInformation)
        Else
            Notify "SA-CCR run stopped - see the Checks sheet.", vbCritical
            ThisWorkbook.Worksheets(SH_CHECKS).Activate
        End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Fail:
        EndBatch calcMode
        Notify "Unexpected error " & Err.Number & ": " & Err.Description, vbCritical

End Sub


Public Sub RunSACCR_Silent()
'
'==============================================================================
'                               RunSACCR_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   Run RunSACCR with its messages printed to the Immediate window instead
'   of shown in message boxes, so it can run unattended.
'
' STATE OWNERSHIP
'   Sets gSilent for the run and clears it afterwards.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN SILENTLY
'------------------------------------------------------------------------------
        gSilent = True
        RunSACCR
        gSilent = False

End Sub


Public Sub ValidateInputs()
'
'==============================================================================
'                                ValidateInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Validate every input without writing results, show the Checks sheet and
'   report the number of errors and warnings.
'
' USAGE
'   Assigned to the Validate inputs button.
'
' STATE OWNERSHIP
'   As RunSACCR. Only the Checks sheet is written.
'
' ERROR POLICY
'   As RunSACCR.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ok         As Boolean    'Engine result; not used for the message
    Dim calcMode   As Long       'Calculation mode captured by BeginBatch

'------------------------------------------------------------------------------
' VALIDATE AND REPORT
'------------------------------------------------------------------------------
        On Error GoTo Fail
        BeginBatch calcMode
        ok = M_Engine.Calculate(False)
        EndBatch calcMode
        ThisWorkbook.Worksheets(SH_CHECKS).Activate
        Notify "Validation finished: " & M_Engine.ErrorCount & " error(s), " & _
               M_Engine.WarningCount & " warning(s).", _
               IIf(M_Engine.ErrorCount > 0, vbExclamation, vbInformation)
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Fail:
        EndBatch calcMode
        Notify "Unexpected error " & Err.Number & ": " & Err.Description, vbCritical

End Sub


Public Sub ClearOutputs()
'
'==============================================================================
'                                 ClearOutputs
'------------------------------------------------------------------------------
' PURPOSE
'   Empty the five output sheets and stamp the time on the Results sheet.
'
' USAGE
'   Assigned to the Clear outputs button.
'
' STATE OWNERSHIP
'   As RunSACCR.
'
' ERROR POLICY
'   As RunSACCR.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim calcMode   As Long    'Calculation mode captured by BeginBatch

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        On Error GoTo Fail
        BeginBatch calcMode
        ClearOutputBlock GetSheet(SH_TRADECALC), FIRST_DATA_ROW, TC_NCOLS
        ClearOutputBlock GetSheet(SH_BUCKETS), FIRST_DATA_ROW, BK_NCOLS
        ClearOutputBlock GetSheet(SH_HEDGING), FIRST_DATA_ROW, HS_NCOLS
        ClearOutputBlock GetSheet(SH_RESULTS), FIRST_DATA_ROW, RS_NCOLS
        ClearOutputBlock GetSheet(SH_CHECKS), FIRST_DATA_ROW, CK_NCOLS
        GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value = "Outputs cleared " & Format$(Now, "yyyy-mm-dd hh:mm:ss")
        EndBatch calcMode
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Fail:
        EndBatch calcMode
        Notify "Unexpected error " & Err.Number & ": " & Err.Description, vbCritical

End Sub


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub BeginBatch( _
    ByRef calcMode As Long)
'
'==============================================================================
'                                  BeginBatch
'------------------------------------------------------------------------------
' PURPOSE
'   Prepare Excel for a run: no screen updating, no events, manual
'   calculation.
'
' RETURNS
'   calcMode: the calculation mode before the run, for EndBatch; 0 when it
'   could not be read.
'
' ERROR POLICY
'   Best effort: errors are ignored so that a protected or busy Excel does
'   not stop the run.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' PREPARE EXCEL
'------------------------------------------------------------------------------
        On Error Resume Next
        calcMode = Application.Calculation
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual

End Sub


Private Sub EndBatch( _
    ByVal calcMode As Long)
'
'==============================================================================
'                                   EndBatch
'------------------------------------------------------------------------------
' PURPOSE
'   Set Excel back after a run: the captured calculation mode, events on and
'   screen updating on.
'
' INPUTS
'   calcMode: the value captured by BeginBatch; 0 leaves calculation as it is.
'
' ERROR POLICY
'   Best effort: errors are ignored.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RESTORE EXCEL
'------------------------------------------------------------------------------
        On Error Resume Next
        If calcMode <> 0 Then
            Application.Calculation = calcMode
        End If
        Application.EnableEvents = True
        Application.ScreenUpdating = True

End Sub


Private Sub Notify( _
    ByVal msg As String, _
    ByVal style As VbMsgBoxStyle)
'
'==============================================================================
'                                    Notify
'------------------------------------------------------------------------------
' PURPOSE
'   Show a message to the user, or print it in silent mode.
'
' INPUTS
'   msg: the text to show.
'   style: message-box icon, for example vbInformation or vbCritical.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' SHOW OR PRINT
'------------------------------------------------------------------------------
        If gSilent Then
            Debug.Print msg
        Else
            MsgBox msg, style, "SA-CCR"
        End If

End Sub
