Attribute VB_Name = "M_Main"
'==============================================================================
' Module   : M_Main
' Purpose  : User-facing macros (assigned to the buttons on the sheets).
'   RunSACCR        - validate inputs, calculate, write all output sheets
'   ValidateInputs  - validation only, findings on the Checks sheet
'   ClearOutputs    - empty all output sheets
'   RunSACCR_Silent - same as RunSACCR without message boxes (automation)
'==============================================================================
Option Explicit

Public gSilent As Boolean

Public Sub RunSACCR()
    Dim ok As Boolean, calcMode As Long, msg As String
    On Error GoTo Fail
    BeginBatch calcMode
    ok = M_Engine.Calculate(True)
    EndBatch calcMode
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
Fail:
    EndBatch calcMode
    Notify "Unexpected error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Public Sub RunSACCR_Silent()
    gSilent = True
    RunSACCR
    gSilent = False
End Sub

Public Sub ValidateInputs()
    Dim ok As Boolean, calcMode As Long
    On Error GoTo Fail
    BeginBatch calcMode
    ok = M_Engine.Calculate(False)
    EndBatch calcMode
    ThisWorkbook.Worksheets(SH_CHECKS).Activate
    Notify "Validation finished: " & M_Engine.ErrorCount & " error(s), " & _
           M_Engine.WarningCount & " warning(s).", _
           IIf(M_Engine.ErrorCount > 0, vbExclamation, vbInformation)
    Exit Sub
Fail:
    EndBatch calcMode
    Notify "Unexpected error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Public Sub ClearOutputs()
    Dim calcMode As Long
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
Fail:
    EndBatch calcMode
    Notify "Unexpected error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

'--- helpers -------------------------------------------------------------------
Private Sub BeginBatch(ByRef calcMode As Long)
    On Error Resume Next
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub EndBatch(ByVal calcMode As Long)
    On Error Resume Next
    If calcMode <> 0 Then Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub Notify(ByVal msg As String, ByVal style As VbMsgBoxStyle)
    If gSilent Then
        Debug.Print msg
    Else
        MsgBox msg, style, "SA-CCR"
    End If
End Sub
