Attribute VB_Name = "M_Util"
'==============================================================================
' Module   : M_Util
' Purpose  : Type-safe conversions, keyed lookups (Collection based, no
'            external references: works on Windows and Mac), sheet helpers.
'==============================================================================
Option Explicit
Option Private Module
Option Private Module

'--- Conversions ---------------------------------------------------------------

' CStr that never fails on error values (#N/A, #VALUE!, ...)
Public Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Then
        SafeStr = ""
    ElseIf IsEmpty(v) Or IsNull(v) Then
        SafeStr = ""
    Else
        SafeStr = Trim$(CStr(v))
    End If
End Function

' Upper-case trimmed text
Public Function UTxt(ByVal v As Variant) As String
    UTxt = UCase$(SafeStr(v))
End Function

' True when the cell is empty or contains only blanks
Public Function IsBlankCell(ByVal v As Variant) As Boolean
    IsBlankCell = (Len(SafeStr(v)) = 0)
End Function

' True when the value is a usable number
Public Function IsNum(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then
        IsNum = False
    ElseIf VarType(v) = vbString Then
        IsNum = (Len(Trim$(v)) > 0) And IsNumeric(v)
    Else
        IsNum = IsNumeric(v)
    End If
End Function

' Number or default
Public Function ToDbl(ByVal v As Variant, Optional ByVal dflt As Double = 0#) As Double
    If IsNum(v) Then
        ToDbl = CDbl(v)
    Else
        ToDbl = dflt
    End If
End Function

' Date serial (as Double) or -1 when missing / invalid
Public Function ToSerial(ByVal v As Variant) As Double
    ToSerial = -1#
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        ToSerial = CDbl(v)
    ElseIf IsNum(v) Then
        ToSerial = CDbl(v)
    ElseIf VarType(v) = vbString Then
        If IsDate(v) Then ToSerial = CDbl(CDate(v))
    End If
End Function

' Y/YES/TRUE/1 -> True
Public Function ToBool(ByVal v As Variant, Optional ByVal dflt As Boolean = False) As Boolean
    Dim s As String
    If IsError(v) Or IsEmpty(v) Then
        ToBool = dflt
        Exit Function
    End If
    If VarType(v) = vbBoolean Then
        ToBool = v
        Exit Function
    End If
    s = UTxt(v)
    Select Case s
        Case "Y", "YES", "TRUE", "1", "VERO", "SI"
            ToBool = True
        Case "N", "NO", "FALSE", "0", "FALSO"
            ToBool = False
        Case Else
            ToBool = dflt
    End Select
End Function

Public Function Max2(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then Max2 = a Else Max2 = b
End Function

Public Function Min2(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then Min2 = a Else Min2 = b
End Function

'--- Keyed lookups (key -> 1-based index) --------------------------------------

Public Function KeyIndex(ByVal col As Collection, ByVal key As String) As Long
    Dim v As Variant
    On Error Resume Next
    v = col.Item(UCase$(key))
    If Err.Number <> 0 Then
        Err.Clear
        KeyIndex = 0
    Else
        KeyIndex = CLng(v)
    End If
    On Error GoTo 0
End Function

Public Sub KeyAdd(ByVal col As Collection, ByVal key As String, ByVal idx As Long)
    col.Add idx, UCase$(key)
End Sub

'--- Asset classes -------------------------------------------------------------

Public Function ACIndex(ByVal code As String) As Long
    Select Case UCase$(code)
        Case "IR": ACIndex = AC_IR
        Case "FX": ACIndex = AC_FX
        Case "CR": ACIndex = AC_CR
        Case "EQ": ACIndex = AC_EQ
        Case "CO": ACIndex = AC_CO
        Case "OT": ACIndex = AC_OT
        Case Else: ACIndex = 0
    End Select
End Function

Public Function ACCode(ByVal idx As Long) As String
    Select Case idx
        Case AC_IR: ACCode = "IR"
        Case AC_FX: ACCode = "FX"
        Case AC_CR: ACCode = "CR"
        Case AC_EQ: ACCode = "EQ"
        Case AC_CO: ACCode = "CO"
        Case AC_OT: ACCode = "OT"
        Case Else: ACCode = "?"
    End Select
End Function

'--- Sheet helpers -------------------------------------------------------------

Public Function GetSheet(ByVal sheetName As String) As Worksheet
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
End Function

' Last row of the sheet's used range (upper bound for data scans)
Public Function UsedLastRow(ByVal ws As Worksheet) As Long
    With ws.UsedRange
        UsedLastRow = .Row + .Rows.Count - 1
    End With
End Function

' Column values from row 1 to the last used row as a 1-based 2-D array
Private Function ColumnValues(ByVal ws As Worksheet, ByVal firstRow As Long, _
                              ByVal lastRow As Long, ByVal colIdx As Long) As Variant
    Dim v As Variant, a() As Variant
    v = ws.Range(ws.Cells(firstRow, colIdx), ws.Cells(lastRow, colIdx)).Value
    If IsArray(v) Then
        ColumnValues = v
    Else
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = v
        ColumnValues = a
    End If
End Function

' Last non-empty row in a column (returns firstRow - 1 when empty).
' Scans the used range bottom-up instead of End(xlUp) so that it behaves the
' same in hidden / automated sessions.
Public Function LastDataRow(ByVal ws As Worksheet, ByVal firstRow As Long, _
                            ByVal colIdx As Long) As Long
    Dim lastR As Long, v As Variant, r As Long
    LastDataRow = firstRow - 1
    lastR = UsedLastRow(ws)
    If lastR < firstRow Then Exit Function
    v = ColumnValues(ws, firstRow, lastR, colIdx)
    For r = UBound(v, 1) To 1 Step -1
        If Not IsBlankCell(v(r, 1)) Then
            LastDataRow = firstRow + r - 1
            Exit Function
        End If
    Next r
End Function

' Row of the cell in column colIdx whose text equals header (0 if not found)
Public Function FindHeaderRow(ByVal ws As Worksheet, ByVal colIdx As Long, _
                              ByVal header As String) As Long
    Dim lastR As Long, v As Variant, r As Long
    FindHeaderRow = 0
    lastR = UsedLastRow(ws)
    If lastR < 1 Then Exit Function
    v = ColumnValues(ws, 1, lastR, colIdx)
    For r = 1 To UBound(v, 1)
        If StrComp(SafeStr(v(r, 1)), header, vbTextCompare) = 0 Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
End Function

' Parameter value: workbook name first, then code lookup on the Params sheet
Public Function GetParam(ByVal code As String) As Variant
    Dim r As Range, ws As Worksheet, i As Long
    On Error Resume Next
    Set r = ThisWorkbook.Names(code).RefersToRange
    Err.Clear
    On Error GoTo 0
    If Not r Is Nothing Then
        GetParam = r.Cells(1, 1).Value
        Exit Function
    End If
    Set ws = GetSheet(SH_PARAMS)
    i = FindHeaderRow(ws, PRM_CODE_COL, code)
    If i > 0 Then
        GetParam = ws.Cells(i, PRM_VALUE_COL).Value
    Else
        GetParam = Empty
    End If
End Function

' Clear an output block (values, bold, fill) from firstRow down
Public Sub ClearOutputBlock(ByVal ws As Worksheet, ByVal firstRow As Long, _
                            ByVal nCols As Long)
    Dim lastR As Long
    lastR = UsedLastRow(ws)
    If lastR < firstRow Then Exit Sub
    With ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastR, nCols))
        .ClearContents
        .Font.Bold = False
        .Interior.ColorIndex = xlNone
    End With
End Sub

' Write a 2-D array (1-based) to a sheet starting at (firstRow, 1)
Public Sub WriteBlock(ByVal ws As Worksheet, ByVal firstRow As Long, _
                      ByRef arr() As Variant, ByVal nRows As Long, ByVal nCols As Long)
    If nRows <= 0 Then Exit Sub
    ws.Range(ws.Cells(firstRow, 1), ws.Cells(firstRow + nRows - 1, nCols)).Value = arr
End Sub

' Apply number formats (pipe separated, one per column, "" = leave) to rows
Public Sub FormatColumns(ByVal ws As Worksheet, ByVal firstRow As Long, _
                         ByVal nRows As Long, ByVal formats As String)
    Dim parts() As String, c As Long
    If nRows <= 0 Then Exit Sub
    parts = Split(formats, "|")
    For c = 0 To UBound(parts)
        If Len(parts(c)) > 0 Then
            ws.Range(ws.Cells(firstRow, c + 1), ws.Cells(firstRow + nRows - 1, c + 1)).NumberFormat = parts(c)
        End If
    Next c
End Sub
