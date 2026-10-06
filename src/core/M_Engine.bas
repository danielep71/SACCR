Attribute VB_Name = "M_Engine"
'==============================================================================
' Module   : M_Engine
' Purpose  : SA-CCR exposure at default per netting set
'
'   EAD = alpha * (RC + PFE)                                   [CRE52.1]
'   PFE = multiplier * AddOn_aggregate                         [CRE52.20]
'   AddOn_aggregate = sum of asset-class add-ons (IR,FX,CR,EQ,CO) [CRE52.25]
'   Trade effective notional = delta * d * MF
'
'   For margined netting sets the EAD is capped at the EAD of the same
'   netting set computed on an unmargined basis               [CRE52.2]
'
' Regime layer (default on Params, override per netting set):
'   BCBS : cap computed with C = VM + NICA (posted VM negative)    [CRE52.2]
'   CRR  : cap per Art. 274(3) with RC per Art. 275(1); NICA excludes
'          VM posted or received (EBA Q&A 2023_6962)
'   CRR  : 'other risks' category OT, SF 8%                    [Art. 280f]
'   CRR  : lambda for IR / commodity options per Del. Reg. 2021/931 Art. 5
'   SF-table rows can be restricted to one regime (column 'Regimes').
'
' Flow : LoadParams -> LoadSFTable -> LoadFXTable -> LoadNettingSets
'        -> ProcessTrades -> ComputeHedgingSets -> ComputeNettingSets
'        -> write TradeCalc / Buckets / HedgingSets / Results / Checks
'==============================================================================
Option Explicit
Option Private Module
Option Private Module

'--- Data structures -----------------------------------------------------------
Private Type tNettingSet
    ID As String
    Counterparty As String
    Margined As Boolean
    Cleared As Boolean
    RemarginBD As Double
    LargeOrIlliquid As Boolean
    Disputes As Boolean
    MPOR As Double              ' effective MPOR, business days
    MFMargined As Double
    VM As Double                ' net variation margin held (+) / posted (-)
    NICA As Double              ' net independent collateral amount
    TH As Double
    MTA As Double
    Alpha As Double
    Regime As String            ' BCBS or CRR
    IsCRR As Boolean
    V As Double                 ' sum of trade MtM (reporting ccy)
    Trades As Long
End Type

Private Type tSupervisory
    Key As String
    AssetClass As String
    Category As String
    SF As Double
    Corr As Double
    Vol As Double
    Group As String             ' commodity hedging set (ENERGY, METALS, ...)
    Regimes As String           ' BOTH (blank), BCBS or CRR
End Type

Private Type tBucket            ' IR maturity bucket / FX pair / entity / commodity type
    NSIdx As Long
    AC As Long
    HSKey As String
    SubKey As String
    SF As Double                ' after basis / volatility scaling
    Corr As Double
    ENU As Double               ' sum(delta*d*MF), unmargined MF
    ENM As Double               ' sum(delta*d*MF), margined MF
    Trades As Long
    HSIdx As Long
End Type

Private Type tHedgingSet
    NSIdx As Long
    AC As Long
    Key As String
    D1U As Double
    D2U As Double
    D3U As Double
    D1M As Double
    D2M As Double
    D3M As Double
    SysU As Double
    IdioU As Double
    SysM As Double
    IdioM As Double
    ENU As Double
    ENM As Double
    AddOnU As Double
    AddOnM As Double
    Trades As Long
End Type

'--- State ---------------------------------------------------------------------
Private mNS() As tNettingSet
Private mNSCount As Long
Private mNSIndex As Collection

Private mSF() As tSupervisory
Private mSFCount As Long
Private mSFIndex As Collection

Private mFXRate() As Double
Private mFXCount As Long
Private mFXIndex As Collection

Private mBk() As tBucket
Private mBkCount As Long
Private mBkIndex As Collection

Private mHS() As tHedgingSet
Private mHSCount As Long
Private mHSIndex As Collection

Private mAddOnU() As Double     ' (netting set, asset class)
Private mAddOnM() As Double

Private mLog() As Variant
Private mLogCount As Long
Private mErrCount As Long
Private mWarnCount As Long

Private mTradesRead As Long
Private mTradesUsed As Long
Private mTotalEAD As Double

'--- Parameters ----------------------------------------------------------------
Private pAsOf As Double
Private pRepCcy As String
Private pAlpha As Double
Private pFloor As Double
Private pDaysYear As Double
Private pBDYear As Double
Private pMinMatBD As Double
Private pSDFloorBD As Double
Private pMPORBil As Double
Private pMPORClr As Double
Private pMPORLarge As Double
Private pBasisF As Double
Private pVolF As Double
Private pRho12 As Double
Private pRho23 As Double
Private pRho13 As Double
Private pIRFull As Boolean
Private pRegime As String
Private pLamThrIR As Double
Private pLamThrCO As Double

'==============================================================================
' Public entry point. writeOutputs = False runs validation only.
' Returns True when the run completed (individual trades may still be
' excluded - see the Checks sheet).
'==============================================================================
Public Function Calculate(ByVal writeOutputs As Boolean) As Boolean
    Dim t0 As Double
    t0 = Timer
    ResetState

    If Not LoadParams() Then GoTo Finish
    If Not LoadSFTable() Then GoTo Finish
    If Not LoadFXTable() Then GoTo Finish
    If Not LoadNettingSets() Then GoTo Finish

    ProcessTrades writeOutputs
    ComputeHedgingSets
    ComputeNettingSets writeOutputs
    If writeOutputs Then
        WriteBuckets
        WriteHedgingSets
    End If
    Calculate = True

Finish:
    WriteChecks
    If writeOutputs Then WriteRunInfo Timer - t0, Calculate
End Function

Public Property Get ErrorCount() As Long
    ErrorCount = mErrCount
End Property

Public Property Get WarningCount() As Long
    WarningCount = mWarnCount
End Property

Public Property Get TotalEAD() As Double
    TotalEAD = mTotalEAD
End Property

Public Property Get TradesUsed() As Long
    TradesUsed = mTradesUsed
End Property

Public Property Get TradesRead() As Long
    TradesRead = mTradesRead
End Property

'==============================================================================
' Initialisation
'==============================================================================
Private Sub ResetState()
    mNSCount = 0: mSFCount = 0: mFXCount = 0: mBkCount = 0: mHSCount = 0
    mLogCount = 0: mErrCount = 0: mWarnCount = 0
    mTradesRead = 0: mTradesUsed = 0: mTotalEAD = 0#
    Set mNSIndex = New Collection
    Set mSFIndex = New Collection
    Set mFXIndex = New Collection
    Set mBkIndex = New Collection
    Set mHSIndex = New Collection
    ReDim mNS(1 To 16)
    ReDim mSF(1 To 32)
    ReDim mFXRate(1 To 32)
    ReDim mBk(1 To 64)
    ReDim mHS(1 To 32)
    ReDim mLog(1 To CK_NCOLS, 1 To 64)
End Sub

Private Sub LogMsg(ByVal severity As String, ByVal sheetName As String, _
                   ByVal ref As String, ByVal msg As String, Optional ByVal rowNum As Long = 0)
    mLogCount = mLogCount + 1
    If mLogCount > UBound(mLog, 2) Then
        ReDim Preserve mLog(1 To CK_NCOLS, 1 To mLogCount * 2)
    End If
    mLog(1, mLogCount) = severity
    mLog(2, mLogCount) = sheetName
    If rowNum > 0 Then mLog(3, mLogCount) = rowNum Else mLog(3, mLogCount) = Empty
    mLog(4, mLogCount) = ref
    mLog(5, mLogCount) = msg
    If severity = SEV_ERROR Then mErrCount = mErrCount + 1
    If severity = SEV_WARN Then mWarnCount = mWarnCount + 1
End Sub

Private Function NumParam(ByVal code As String, ByVal dflt As Double) As Double
    Dim v As Variant
    v = GetParam(code)
    If IsNum(v) Then
        NumParam = CDbl(v)
    Else
        NumParam = dflt
        LogMsg SEV_WARN, SH_PARAMS, code, "Parameter missing or not numeric - default " & CStr(dflt) & " used."
    End If
End Function

Private Function LoadParams() As Boolean
    On Error GoTo Fail
    pAsOf = ToSerial(GetParam(PRM_ASOF))
    If pAsOf <= 0# Then
        LogMsg SEV_ERROR, SH_PARAMS, PRM_ASOF, "Reporting date missing or invalid."
        Exit Function
    End If
    pRepCcy = UTxt(GetParam(PRM_REPCCY))
    If Len(pRepCcy) = 0 Then
        LogMsg SEV_ERROR, SH_PARAMS, PRM_REPCCY, "Reporting currency missing."
        Exit Function
    End If
    pAlpha = NumParam(PRM_ALPHA, 1.4)
    pFloor = NumParam(PRM_FLOOR, 0.05)
    pDaysYear = NumParam(PRM_DAYSYEAR, 365#)
    pBDYear = NumParam(PRM_BDYEAR, 250#)
    pMinMatBD = NumParam(PRM_MINMAT, 10#)
    pSDFloorBD = NumParam(PRM_SDFLOOR, 10#)
    pMPORBil = NumParam(PRM_MPOR_BIL, 10#)
    pMPORClr = NumParam(PRM_MPOR_CLR, 5#)
    pMPORLarge = NumParam(PRM_MPOR_LARGE, 20#)
    pBasisF = NumParam(PRM_BASIS, 0.5)
    pVolF = NumParam(PRM_VOLF, 5#)
    pRho12 = NumParam(PRM_RHO12, 0.7)
    pRho23 = NumParam(PRM_RHO23, 0.7)
    pRho13 = NumParam(PRM_RHO13, 0.3)
    pIRFull = ToBool(GetParam(PRM_IRFULL), True)
    pRegime = UTxt(GetParam(PRM_REGIME))
    If Len(pRegime) = 0 Then
        pRegime = RG_CRR
        LogMsg SEV_WARN, SH_PARAMS, PRM_REGIME, "Regime missing - CRR assumed."
    ElseIf pRegime <> RG_BCBS And pRegime <> RG_CRR Then
        LogMsg SEV_ERROR, SH_PARAMS, PRM_REGIME, "Regime must be BCBS or CRR."
        Exit Function
    End If
    pLamThrIR = NumParam(PRM_LAMIR, 0.001)
    pLamThrCO = NumParam(PRM_LAMCO, 0.1)
    If pDaysYear <= 0# Or pBDYear <= 0# Then
        LogMsg SEV_ERROR, SH_PARAMS, PRM_DAYSYEAR, "Day-count parameters must be positive."
        Exit Function
    End If
    LoadParams = True
    Exit Function
Fail:
    LogMsg SEV_ERROR, SH_PARAMS, "", "Cannot read parameters: " & Err.Description
End Function

Private Function LoadSFTable() As Boolean
    Dim ws As Worksheet, r As Long, r0 As Long, k As String
    Set ws = GetSheet(SH_PARAMS)
    r0 = FindHeaderRow(ws, 1, HDR_SF)
    If r0 = 0 Then
        LogMsg SEV_ERROR, SH_PARAMS, HDR_SF, "Supervisory factor table not found (header '" & HDR_SF & "' in column A)."
        Exit Function
    End If
    r = r0 + 1
    Do While Not IsBlankCell(ws.Cells(r, 1).Value)
        k = UTxt(ws.Cells(r, 1).Value)
        If KeyIndex(mSFIndex, k) > 0 Then
            LogMsg SEV_WARN, SH_PARAMS, k, "Duplicate supervisory factor key - first occurrence used.", r
        Else
            mSFCount = mSFCount + 1
            If mSFCount > UBound(mSF) Then
                ReDim Preserve mSF(1 To mSFCount * 2)
            End If
            mSF(mSFCount).Key = k
            mSF(mSFCount).AssetClass = UTxt(ws.Cells(r, 2).Value)
            mSF(mSFCount).Category = UTxt(ws.Cells(r, 3).Value)
            mSF(mSFCount).SF = ToDbl(ws.Cells(r, 4).Value)
            mSF(mSFCount).Corr = ToDbl(ws.Cells(r, 5).Value)
            mSF(mSFCount).Vol = ToDbl(ws.Cells(r, 6).Value)
            mSF(mSFCount).Group = UTxt(ws.Cells(r, 7).Value)
            mSF(mSFCount).Regimes = UTxt(ws.Cells(r, 8).Value)
            If mSF(mSFCount).Regimes = "BOTH" Then
                mSF(mSFCount).Regimes = ""
            End If
            KeyAdd mSFIndex, k, mSFCount
        End If
        r = r + 1
    Loop
    If mSFCount = 0 Then
        LogMsg SEV_ERROR, SH_PARAMS, HDR_SF, "Supervisory factor table is empty."
        Exit Function
    End If
    LoadSFTable = True
End Function

Private Function LoadFXTable() As Boolean
    Dim ws As Worksheet, r As Long, r0 As Long, k As String, rate As Double
    Set ws = GetSheet(SH_PARAMS)
    r0 = FindHeaderRow(ws, 1, HDR_FX)
    If r0 = 0 Then
        LogMsg SEV_ERROR, SH_PARAMS, HDR_FX, "FX table not found (header '" & HDR_FX & "' in column A)."
        Exit Function
    End If
    r = r0 + 1
    Do While Not IsBlankCell(ws.Cells(r, 1).Value)
        k = UTxt(ws.Cells(r, 1).Value)
        rate = ToDbl(ws.Cells(r, 3).Value, -1#)
        If rate <= 0# Then
            LogMsg SEV_WARN, SH_PARAMS, k, "FX rate missing or not positive - currency ignored.", r
        ElseIf KeyIndex(mFXIndex, k) = 0 Then
            mFXCount = mFXCount + 1
            If mFXCount > UBound(mFXRate) Then
                ReDim Preserve mFXRate(1 To mFXCount * 2)
            End If
            mFXRate(mFXCount) = rate
            KeyAdd mFXIndex, k, mFXCount
        End If
        r = r + 1
    Loop
    If KeyIndex(mFXIndex, pRepCcy) = 0 Then
        mFXCount = mFXCount + 1
        If mFXCount > UBound(mFXRate) Then
            ReDim Preserve mFXRate(1 To mFXCount * 2)
        End If
        mFXRate(mFXCount) = 1#
        KeyAdd mFXIndex, pRepCcy, mFXCount
    ElseIf Abs(mFXRate(KeyIndex(mFXIndex, pRepCcy)) - 1#) > 0.0000001 Then
        LogMsg SEV_WARN, SH_PARAMS, pRepCcy, "Rate of the reporting currency is not 1."
    End If
    LoadFXTable = True
End Function

' Units of reporting currency per unit of ccy, or -1 if unknown
Private Function FXRate(ByVal ccy As String) As Double
    Dim i As Long
    If Len(ccy) = 0 Then ccy = pRepCcy
    i = KeyIndex(mFXIndex, ccy)
    If i = 0 Then FXRate = -1# Else FXRate = mFXRate(i)
End Function

Private Function LoadNettingSets() As Boolean
    Dim ws As Worksheet, lastR As Long, n As Long, i As Long, rowNum As Long
    Dim data As Variant, id As String, floorBD As Double, mpor As Double, ovr As Double
    Dim rg As String
    Set ws = GetSheet(SH_NS)
    lastR = LastDataRow(ws, FIRST_DATA_ROW, NS_ID)
    If lastR < FIRST_DATA_ROW Then
        LogMsg SEV_ERROR, SH_NS, "", "No netting sets defined."
        Exit Function
    End If
    n = lastR - FIRST_DATA_ROW + 1
    data = ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(lastR, NS_NCOLS)).Value
    For i = 1 To n
        rowNum = FIRST_DATA_ROW + i - 1
        id = UTxt(data(i, NS_ID))
        If Len(id) = 0 Then GoTo NextRow
        If KeyIndex(mNSIndex, id) > 0 Then
            LogMsg SEV_ERROR, SH_NS, id, "Duplicate netting set ID - row ignored.", rowNum
            GoTo NextRow
        End If
        mNSCount = mNSCount + 1
        If mNSCount > UBound(mNS) Then
            ReDim Preserve mNS(1 To mNSCount * 2)
        End If
        With mNS(mNSCount)
            .ID = id
            .Counterparty = SafeStr(data(i, NS_CPTY))
            .Margined = ToBool(data(i, NS_MARGINED), False)
            .Cleared = ToBool(data(i, NS_CLEARED), False)
            .RemarginBD = ToDbl(data(i, NS_FREQ), 1#)
            If .RemarginBD < 1# Then .RemarginBD = 1#
            .LargeOrIlliquid = ToBool(data(i, NS_LARGE), False)
            .Disputes = ToBool(data(i, NS_DISPUTE), False)
            .VM = ToDbl(data(i, NS_VM))
            .NICA = ToDbl(data(i, NS_NICA))
            .TH = ToDbl(data(i, NS_TH))
            .MTA = ToDbl(data(i, NS_MTA))
            .Alpha = ToDbl(data(i, NS_ALPHA), pAlpha)
            If .Alpha <= 0# Then .Alpha = pAlpha
            rg = UTxt(data(i, NS_REGIME))
            If Len(rg) = 0 Then
                rg = pRegime
            ElseIf rg <> RG_BCBS And rg <> RG_CRR Then
                LogMsg SEV_WARN, SH_NS, id, "Regime override must be BCBS or CRR - default " & pRegime & " used.", rowNum
                rg = pRegime
            End If
            .Regime = rg
            .IsCRR = (rg = RG_CRR)
            If Not .Margined And Abs(.VM) > 0# Then
                LogMsg SEV_WARN, SH_NS, id, "VM entered on an unmargined netting set - treated as collateral C. " & _
                       "Under CRR collateral of an unmargined set belongs in NICA.", rowNum
            End If
            .V = 0#
            .Trades = 0
            .MPOR = 0#
            .MFMargined = 0#
            If .Margined Then
                ' MPOR floors [CRE52.50-52.52]; cleared floor per CRE54 (configurable)
                If .Cleared Then floorBD = pMPORClr Else floorBD = pMPORBil
                If .LargeOrIlliquid Then floorBD = Max2(floorBD, pMPORLarge)
                mpor = floorBD + .RemarginBD - 1#
                If .Disputes Then mpor = 2# * mpor
                ovr = ToDbl(data(i, NS_MPOR), 0#)
                If ovr > 0# Then
                    If ovr < mpor Then
                        LogMsg SEV_WARN, SH_NS, id, "MPOR override " & CStr(ovr) & _
                               " BD is below the regulatory floor - " & CStr(mpor) & " BD applied.", rowNum
                    Else
                        mpor = ovr
                    End If
                End If
                .MPOR = mpor
                .MFMargined = SACCR_MaturityFactor(0#, True, .MPOR, pMinMatBD, pBDYear)
            End If
        End With
        KeyAdd mNSIndex, id, mNSCount
NextRow:
    Next i
    If mNSCount = 0 Then
        LogMsg SEV_ERROR, SH_NS, "", "No valid netting sets."
        Exit Function
    End If
    ReDim mAddOnU(1 To mNSCount, 1 To AC_COUNT)
    ReDim mAddOnM(1 To mNSCount, 1 To AC_COUNT)
    LoadNettingSets = True
End Function

'==============================================================================
' Trade level calculation
'==============================================================================
Private Sub AddErr(ByRef ok As Boolean, ByRef msg As String, ByVal txt As String)
    ok = False
    If Len(msg) > 0 Then msg = msg & "; "
    msg = msg & txt
End Sub

Private Sub AddWarn(ByRef warn As String, ByVal txt As String)
    If Len(warn) > 0 Then warn = warn & "; "
    warn = warn & txt
End Sub

Private Sub ProcessTrades(ByVal writeOutputs As Boolean)
    Dim ws As Worksheet, lastR As Long, n As Long, i As Long, rowNum As Long
    Dim data As Variant, outArr() As Variant, nOut As Long
    Dim ok As Boolean, msg As String, warn As String
    Dim tid As String, nsId As String, acTxt As String, subCls As String, rf As String
    Dim instType As String, dirTxt As String, optTxt As String, nature As String, lbl As String
    Dim nsIdx As Long, ac As Long, sfKey As String, sfIdx As Long
    Dim isLong As Boolean, natFactor As Double, natTag As String
    Dim notional As Double, nccy As String, fxN As Double, notionalRep As Double
    Dim mtm As Double, mccy As String, fxM As Double, mtmRep As Double
    Dim dStart As Double, dEnd As Double, dMat As Double, dExp As Double
    Dim S As Double, E As Double, M As Double, T As Double
    Dim sd As Double, adjN As Double, delta As Double, vDelta As Variant
    Dim mfU As Double, mfM As Double, enU As Double, enM As Double
    Dim hsKey As String, subKey As String, bucketNo As Long
    Dim sfEff As Double, corrEff As Double, pairSign As Double
    Dim P As Double, K As Double, lam As Double, attA As Double, detD As Double
    Dim lamIn As Variant, lamOut As Variant

    Set ws = GetSheet(SH_TRADES)
    lastR = LastDataRow(ws, FIRST_DATA_ROW, TR_ID)
    If lastR < FIRST_DATA_ROW Then
        LogMsg SEV_ERROR, SH_TRADES, "", "No trades found."
        If writeOutputs Then ClearOutputBlock GetSheet(SH_TRADECALC), FIRST_DATA_ROW, TC_NCOLS
        Exit Sub
    End If
    n = lastR - FIRST_DATA_ROW + 1
    data = ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(lastR, TR_NCOLS)).Value
    ReDim outArr(1 To n, 1 To TC_NCOLS)
    nOut = 0

    For i = 1 To n
        rowNum = FIRST_DATA_ROW + i - 1
        tid = SafeStr(data(i, TR_ID))
        If Len(tid) = 0 Then GoTo NextTrade
        mTradesRead = mTradesRead + 1
        ok = True: msg = "": warn = ""
        S = 0#: E = 0#: M = 0#: T = 0#: sd = 0#: adjN = 0#: delta = 0#
        mfU = 0#: mfM = 0#: enU = 0#: enM = 0#: bucketNo = 0
        sfEff = 0#: corrEff = 0#: hsKey = "": subKey = "": sfIdx = 0
        notionalRep = 0#: mtmRep = 0#: pairSign = 1#
        lam = 0#: lamOut = Empty

        '--- classification ----------------------------------------------
        nsId = UTxt(data(i, TR_NS))
        nsIdx = KeyIndex(mNSIndex, nsId)
        If nsIdx = 0 Then AddErr ok, msg, "unknown netting set '" & nsId & "'"

        acTxt = UTxt(data(i, TR_AC))
        ac = ACIndex(acTxt)
        If ac = 0 Then AddErr ok, msg, "asset class must be IR, FX, CR, EQ, CO or OT"

        subCls = UTxt(data(i, TR_SUB))
        rf = UTxt(data(i, TR_RF))
        If Len(rf) = 0 Then AddErr ok, msg, "risk factor / reference missing"

        instType = UTxt(data(i, TR_INSTR))
        If Len(instType) = 0 Then instType = "LINEAR"
        If instType <> "LINEAR" And instType <> "OPTION" And instType <> "CDO" Then
            AddErr ok, msg, "instrument type must be Linear, Option or CDO"
        End If
        If instType = "CDO" And ac <> AC_CR Then AddErr ok, msg, "CDO tranches are credit (CR) trades"

        dirTxt = UTxt(data(i, TR_DIR))
        Select Case dirTxt
            Case "LONG", "BUY", "BOUGHT", "L", "B": isLong = True
            Case "SHORT", "SELL", "SOLD", "S": isLong = False
            Case Else: AddErr ok, msg, "direction must be Long or Short"
        End Select

        nature = UTxt(data(i, TR_NATURE))
        If Len(nature) = 0 Then nature = "STANDARD"
        lbl = UTxt(data(i, TR_LABEL))
        Select Case nature
            Case "STANDARD"
                natFactor = 1#: natTag = "STD"
            Case "BASIS"
                natFactor = pBasisF: natTag = "BASIS:" & lbl
                If Len(lbl) = 0 Then AddWarn warn, "basis trade without hedging-set label"
            Case "VOLATILITY"
                natFactor = pVolF: natTag = "VOL:" & lbl
                If Len(lbl) = 0 Then AddWarn warn, "volatility trade without hedging-set label"
            Case Else
                AddErr ok, msg, "nature must be Standard, Basis or Volatility"
        End Select

        '--- supervisory parameters ----------------------------------------
        If ac > 0 Then
            Select Case ac
                Case AC_IR: sfKey = "IR"
                Case AC_FX: sfKey = "FX"
                Case AC_OT: sfKey = "OT"
                Case Else: sfKey = acTxt & "_" & subCls
            End Select
            sfIdx = KeyIndex(mSFIndex, sfKey)
            If sfIdx = 0 Then AddErr ok, msg, "no supervisory factor for '" & sfKey & "' (check sub-class)"
        End If

        '--- regime-specific availability ------------------------------------
        If nsIdx > 0 And ac = AC_OT Then
            If Not mNS(nsIdx).IsCRR Then
                AddErr ok, msg, "asset class OT (other risks) exists only under CRR (Art. 277(1)(f), 280f); " & _
                       "under BCBS map the trade to the class of its primary risk driver (CRE52)"
            End If
        End If
        If nsIdx > 0 And sfIdx > 0 Then
            If Len(mSF(sfIdx).Regimes) > 0 And mSF(sfIdx).Regimes <> mNS(nsIdx).Regime Then
                AddErr ok, msg, "'" & sfKey & "' is defined only for regime " & mSF(sfIdx).Regimes & _
                       " (netting set uses " & mNS(nsIdx).Regime & ")"
            End If
        End If

        '--- amounts ---------------------------------------------------------
        If IsNum(data(i, TR_NOTIONAL)) Then
            notional = CDbl(data(i, TR_NOTIONAL))
            If notional < 0# Then AddErr ok, msg, "notional must be positive (use Direction for the sign)"
        Else
            AddErr ok, msg, "notional missing"
        End If
        nccy = UTxt(data(i, TR_NCCY))
        fxN = FXRate(nccy)
        If fxN < 0# Then AddErr ok, msg, "no FX rate for notional currency '" & nccy & "'"
        If IsNum(data(i, TR_MTM)) Then
            mtm = CDbl(data(i, TR_MTM))
        Else
            mtm = 0#
            AddWarn warn, "MtM missing - 0 assumed"
        End If
        mccy = UTxt(data(i, TR_MCCY))
        If Len(mccy) = 0 Then mccy = nccy
        fxM = FXRate(mccy)
        If fxM < 0# Then AddErr ok, msg, "no FX rate for MtM currency '" & mccy & "'"
        If ok Then
            notionalRep = notional * fxN
            mtmRep = mtm * fxM
        End If

        '--- dates -> year fractions ------------------------------------------
        dStart = ToSerial(data(i, TR_START))
        dEnd = ToSerial(data(i, TR_END))
        dMat = ToSerial(data(i, TR_MAT))
        dExp = ToSerial(data(i, TR_EXPIRY))
        If dMat < 0# Then dMat = dEnd
        If dMat < 0# Then dMat = dExp
        If dEnd < 0# Then dEnd = dMat
        If dMat < 0# Then
            AddErr ok, msg, "maturity / end date missing"
        Else
            M = (dMat - pAsOf) / pDaysYear
            E = (dEnd - pAsOf) / pDaysYear
            If dStart >= 0# Then S = Max2(0#, (dStart - pAsOf) / pDaysYear) Else S = 0#
            If M <= 0# Then
                AddErr ok, msg, "trade has matured (maturity <= reporting date)"
            ElseIf (ac = AC_IR Or ac = AC_CR) And E <= S Then
                AddErr ok, msg, "end date must be after start date"
            End If
        End If

        '--- supervisory delta ------------------------------------------------
        If ok Then
            Select Case instType
                Case "LINEAR"
                    If isLong Then delta = 1# Else delta = -1#
                Case "OPTION"
                    optTxt = UTxt(data(i, TR_OPT))
                    If optTxt <> "CALL" And optTxt <> "PUT" Then AddErr ok, msg, "option type must be Call or Put"
                    If dExp < 0# Then
                        AddErr ok, msg, "option expiry missing"
                    Else
                        T = (dExp - pAsOf) / pDaysYear
                        If T <= 0# Then AddErr ok, msg, "option has expired"
                    End If
                    If Not IsNum(data(i, TR_PRICE)) Or Not IsNum(data(i, TR_STRIKE)) Then
                        AddErr ok, msg, "option needs underlying price and strike"
                    End If
                    If ok Then
                        P = CDbl(data(i, TR_PRICE))
                        K = CDbl(data(i, TR_STRIKE))
                        lamIn = data(i, TR_LAMBDA)
                        If mNS(nsIdx).IsCRR And (ac = AC_IR Or ac = AC_CO) Then
                            ' Del. Reg. 2021/931 Art. 5 (as amended by 2025/855)
                            If ac = AC_IR Then
                                lam = SACCR_LambdaCRR(P, K, True, pLamThrIR)
                            Else
                                lam = SACCR_LambdaCRR(P, K, False, pLamThrCO)
                            End If
                            If IsNum(lamIn) Then
                                If Abs(CDbl(lamIn) - lam) > 0.000000001 Then
                                    AddWarn warn, "lambda input ignored under CRR - RTS lambda applied"
                                End If
                            End If
                        Else
                            lam = ToDbl(lamIn, 0#)
                        End If
                        lamOut = lam
                        If P + lam <= 0# Or K + lam <= 0# Then
                            AddErr ok, msg, "price and strike must be > 0 after lambda shift (BCBS: enter lambda per CRE52.40 FAQ)"
                        ElseIf mSF(sfIdx).Vol <= 0# Then
                            AddErr ok, msg, "supervisory volatility missing for '" & sfKey & "'"
                        Else
                            vDelta = SACCR_OptionDelta(P, K, T, mSF(sfIdx).Vol, (optTxt = "CALL"), isLong, lam)
                            delta = CDbl(vDelta)
                        End If
                    End If
                Case "CDO"
                    If Not IsNum(data(i, TR_ATTACH)) Or Not IsNum(data(i, TR_DETACH)) Then
                        AddErr ok, msg, "CDO tranche needs attachment and detachment points"
                    Else
                        attA = CDbl(data(i, TR_ATTACH))
                        detD = CDbl(data(i, TR_DETACH))
                        If attA < 0# Or detD > 1# Or attA >= detD Then
                            AddErr ok, msg, "attachment/detachment must satisfy 0 <= A < D <= 1"
                        Else
                            delta = CDbl(SACCR_CDODelta(attA, detD, isLong))
                        End If
                    End If
            End Select
        End If

        '--- hedging set, adjusted notional, maturity factor ----------------
        If ok Then
            Select Case ac
                Case AC_IR
                    sd = SACCR_SupervisoryDuration(S, E, pSDFloorBD / pBDYear)
                    adjN = notionalRep * sd
                    If E < 1# Then
                        bucketNo = 1
                        subKey = "<1Y"
                    ElseIf E <= 5# Then
                        bucketNo = 2
                        subKey = "1-5Y"
                    Else
                        bucketNo = 3
                        subKey = ">5Y"
                    End If
                    hsKey = "IR|" & rf & "|" & natTag
                Case AC_FX
                    rf = NormalizePair(rf, pairSign)
                    If pairSign = 0# Then
                        AddErr ok, msg, "FX risk factor must be a currency pair such as EUR/USD"
                    End If
                    delta = delta * pairSign
                    adjN = notionalRep
                    hsKey = "FX|" & rf & "|" & natTag
                    subKey = rf
                Case AC_CR
                    sd = SACCR_SupervisoryDuration(S, E, pSDFloorBD / pBDYear)
                    adjN = notionalRep * sd
                    hsKey = "CR|" & natTag
                    subKey = rf
                Case AC_EQ
                    adjN = notionalRep
                    hsKey = "EQ|" & natTag
                    subKey = rf
                Case AC_CO
                    adjN = notionalRep
                    If Len(mSF(sfIdx).Group) = 0 Then
                        AddErr ok, msg, "commodity hedging group missing for '" & sfKey & "'"
                    End If
                    hsKey = "CO|" & mSF(sfIdx).Group & "|" & natTag
                    subKey = rf
                Case AC_OT
                    ' Art. 277a: same hedging set only for an identical primary risk driver
                    adjN = notionalRep
                    hsKey = "OT|" & rf & "|" & natTag
                    subKey = rf
            End Select
        End If

        If ok Then
            sfEff = mSF(sfIdx).SF * natFactor
            corrEff = mSF(sfIdx).Corr
            mfU = SACCR_MaturityFactor(M, False, 0#, pMinMatBD, pBDYear)
            If mNS(nsIdx).Margined Then mfM = mNS(nsIdx).MFMargined Else mfM = mfU
            enU = delta * adjN * mfU
            enM = delta * adjN * mfM
            AddToBucket nsIdx, ac, hsKey, subKey, sfEff, corrEff, enU, enM, tid, rowNum
            mNS(nsIdx).V = mNS(nsIdx).V + mtmRep
            mNS(nsIdx).Trades = mNS(nsIdx).Trades + 1
            mTradesUsed = mTradesUsed + 1
        End If

        If Not ok Then LogMsg SEV_ERROR, SH_TRADES, tid, "Trade excluded: " & msg, rowNum
        If Len(warn) > 0 Then LogMsg SEV_WARN, SH_TRADES, tid, warn, rowNum

        '--- output row ------------------------------------------------------
        nOut = nOut + 1
        outArr(nOut, 1) = tid
        outArr(nOut, 2) = nsId
        outArr(nOut, 3) = acTxt
        If ok Then
            outArr(nOut, 4) = "OK"
            outArr(nOut, 5) = hsKey
            outArr(nOut, 6) = subKey
            outArr(nOut, 7) = sfEff
            outArr(nOut, 8) = notionalRep
            outArr(nOut, 9) = mtmRep
            outArr(nOut, 10) = S
            outArr(nOut, 11) = E
            outArr(nOut, 12) = M
            If instType = "OPTION" Then outArr(nOut, 13) = T
            If ac = AC_IR Or ac = AC_CR Then outArr(nOut, 14) = sd
            outArr(nOut, 15) = adjN
            outArr(nOut, 16) = delta
            outArr(nOut, 17) = mfU
            If mNS(nsIdx).Margined Then outArr(nOut, 18) = mfM
            outArr(nOut, 19) = enU
            If mNS(nsIdx).Margined Then outArr(nOut, 20) = enM
            If ac = AC_IR Then outArr(nOut, 21) = bucketNo
            outArr(nOut, 22) = sfEff * Abs(enU)
            outArr(nOut, 23) = lamOut
            outArr(nOut, 24) = mNS(nsIdx).Regime
            outArr(nOut, 25) = warn
        Else
            outArr(nOut, 4) = "EXCLUDED"
            If nsIdx > 0 Then
                outArr(nOut, 24) = mNS(nsIdx).Regime
            End If
            outArr(nOut, 25) = msg
        End If
NextTrade:
    Next i

    If writeOutputs Then
        Dim wsOut As Worksheet
        Set wsOut = GetSheet(SH_TRADECALC)
        ClearOutputBlock wsOut, FIRST_DATA_ROW, TC_NCOLS
        WriteBlock wsOut, FIRST_DATA_ROW, outArr, nOut, TC_NCOLS
        FormatColumns wsOut, FIRST_DATA_ROW, nOut, _
            "@|@|@|@|@|@|0.00%|#,##0|#,##0|0.0000|0.0000|0.0000|0.0000|0.0000|#,##0|0.0000|0.0000|0.0000|#,##0|#,##0|0|#,##0|0.0000%|@|@"
    End If
End Sub

' "EUR/USD", "EURUSD", "eur-usd" -> canonical "EUR/USD" (alphabetical order).
' pairSgn = +1 when kept, -1 when inverted, 0 when invalid.
Private Function NormalizePair(ByVal txt As String, ByRef pairSgn As Double) As String
    Dim s As String, c1 As String, c2 As String
    s = UCase$(Replace(Replace(Replace(Replace(txt, "/", ""), "-", ""), " ", ""), ".", ""))
    If Len(s) <> 6 Then
        pairSgn = 0#
        NormalizePair = txt
        Exit Function
    End If
    c1 = Left$(s, 3): c2 = Right$(s, 3)
    If c1 = c2 Then
        pairSgn = 0#
        NormalizePair = txt
    ElseIf c1 < c2 Then
        pairSgn = 1#
        NormalizePair = c1 & "/" & c2
    Else
        pairSgn = -1#
        NormalizePair = c2 & "/" & c1
    End If
End Function

Private Sub AddToBucket(ByVal nsIdx As Long, ByVal ac As Long, ByVal hsKey As String, _
                        ByVal subKey As String, ByVal sfEff As Double, ByVal corr As Double, _
                        ByVal enU As Double, ByVal enM As Double, ByVal tid As String, ByVal rowNum As Long)
    Dim key As String, b As Long
    key = CStr(nsIdx) & "#" & hsKey & "#" & subKey
    b = KeyIndex(mBkIndex, key)
    If b = 0 Then
        mBkCount = mBkCount + 1
        If mBkCount > UBound(mBk) Then
            ReDim Preserve mBk(1 To mBkCount * 2)
        End If
        b = mBkCount
        mBk(b).NSIdx = nsIdx
        mBk(b).AC = ac
        mBk(b).HSKey = hsKey
        mBk(b).SubKey = subKey
        mBk(b).SF = sfEff
        mBk(b).Corr = corr
        mBk(b).ENU = 0#
        mBk(b).ENM = 0#
        mBk(b).Trades = 0
        KeyAdd mBkIndex, key, b
    ElseIf (ac = AC_CR Or ac = AC_EQ Or ac = AC_CO) And Abs(mBk(b).SF - sfEff) > 0.0000000001 Then
        LogMsg SEV_WARN, SH_TRADES, tid, "Sub-class differs from earlier trades on '" & subKey & _
               "' - supervisory factor of the first trade used.", rowNum
    End If
    mBk(b).ENU = mBk(b).ENU + enU
    mBk(b).ENM = mBk(b).ENM + enM
    mBk(b).Trades = mBk(b).Trades + 1
End Sub

'==============================================================================
' Hedging-set and asset-class add-ons
'==============================================================================
Private Sub ComputeHedgingSets()
    Dim b As Long, h As Long, key As String, aU As Double, aM As Double
    ' 1) collect buckets into hedging sets
    For b = 1 To mBkCount
        key = CStr(mBk(b).NSIdx) & "#" & mBk(b).HSKey
        h = KeyIndex(mHSIndex, key)
        If h = 0 Then
            mHSCount = mHSCount + 1
            If mHSCount > UBound(mHS) Then
                ReDim Preserve mHS(1 To mHSCount * 2)
            End If
            h = mHSCount
            mHS(h).NSIdx = mBk(b).NSIdx
            mHS(h).AC = mBk(b).AC
            mHS(h).Key = mBk(b).HSKey
            KeyAdd mHSIndex, key, h
        End If
        mBk(b).HSIdx = h
        mHS(h).Trades = mHS(h).Trades + mBk(b).Trades
        Select Case mBk(b).AC
            Case AC_IR
                Select Case mBk(b).SubKey
                    Case "<1Y"
                        mHS(h).D1U = mHS(h).D1U + mBk(b).ENU
                        mHS(h).D1M = mHS(h).D1M + mBk(b).ENM
                    Case "1-5Y"
                        mHS(h).D2U = mHS(h).D2U + mBk(b).ENU
                        mHS(h).D2M = mHS(h).D2M + mBk(b).ENM
                    Case Else
                        mHS(h).D3U = mHS(h).D3U + mBk(b).ENU
                        mHS(h).D3M = mHS(h).D3M + mBk(b).ENM
                End Select
            Case AC_FX, AC_OT
                mHS(h).ENU = mHS(h).ENU + mBk(b).ENU
                mHS(h).ENM = mHS(h).ENM + mBk(b).ENM
            Case Else
                ' entity / commodity-type add-on, then single-factor aggregation
                aU = mBk(b).SF * mBk(b).ENU
                aM = mBk(b).SF * mBk(b).ENM
                mHS(h).SysU = mHS(h).SysU + mBk(b).Corr * aU
                mHS(h).IdioU = mHS(h).IdioU + (1# - mBk(b).Corr ^ 2) * aU * aU
                mHS(h).SysM = mHS(h).SysM + mBk(b).Corr * aM
                mHS(h).IdioM = mHS(h).IdioM + (1# - mBk(b).Corr ^ 2) * aM * aM
        End Select
    Next b

    ' 2) hedging-set add-ons
    Dim sfH As Double
    For h = 1 To mHSCount
        Select Case mHS(h).AC
            Case AC_IR
                sfH = BucketSF(h)
                If pIRFull Then
                    mHS(h).ENU = SACCR_IREffectiveNotional(mHS(h).D1U, mHS(h).D2U, mHS(h).D3U, pRho12, pRho23, pRho13)
                    mHS(h).ENM = SACCR_IREffectiveNotional(mHS(h).D1M, mHS(h).D2M, mHS(h).D3M, pRho12, pRho23, pRho13)
                Else
                    mHS(h).ENU = Abs(mHS(h).D1U) + Abs(mHS(h).D2U) + Abs(mHS(h).D3U)
                    mHS(h).ENM = Abs(mHS(h).D1M) + Abs(mHS(h).D2M) + Abs(mHS(h).D3M)
                End If
                mHS(h).AddOnU = sfH * mHS(h).ENU
                mHS(h).AddOnM = sfH * mHS(h).ENM
            Case AC_FX, AC_OT
                ' FX [CRE52.59]; other risks [Art. 280f]: SF x |effective notional|
                sfH = BucketSF(h)
                mHS(h).AddOnU = sfH * Abs(mHS(h).ENU)
                mHS(h).AddOnM = sfH * Abs(mHS(h).ENM)
            Case Else
                mHS(h).AddOnU = Sqr(mHS(h).SysU ^ 2 + mHS(h).IdioU)
                mHS(h).AddOnM = Sqr(mHS(h).SysM ^ 2 + mHS(h).IdioM)
        End Select
        mAddOnU(mHS(h).NSIdx, mHS(h).AC) = mAddOnU(mHS(h).NSIdx, mHS(h).AC) + mHS(h).AddOnU
        mAddOnM(mHS(h).NSIdx, mHS(h).AC) = mAddOnM(mHS(h).NSIdx, mHS(h).AC) + mHS(h).AddOnM
    Next h
End Sub

' Credit, equity and commodity use entity/type add-ons with single-factor aggregation
Private Function IsFactorClass(ByVal ac As Long) As Boolean
    IsFactorClass = (ac = AC_CR Or ac = AC_EQ Or ac = AC_CO)
End Function

' Supervisory factor of the first bucket of a hedging set (IR/FX/OT: uniform)
Private Function BucketSF(ByVal h As Long) As Double
    Dim b As Long
    For b = 1 To mBkCount
        If mBk(b).HSIdx = h Then
            BucketSF = mBk(b).SF
            Exit Function
        End If
    Next b
End Function

'==============================================================================
' Netting-set results
'==============================================================================
Private Sub ComputeNettingSets(ByVal writeOutputs As Boolean)
    ' Results layout (RS_NCOLS = 27):
    '  1 ID  2 Counterparty  3 Regime  4 Margined  5 Trades  6 MPOR  7 V  8 C  9 RC
    '  10-15 AddOn IR FX CR EQ CO OT  16 AddOn aggregate  17 Multiplier  18 PFE
    '  19 Alpha  20 EAD margined  21 EAD cap  22 EAD  23 Cap applied
    '  24 C (cap basis)  25 RC (cap basis)  26 AddOn unmargined  27 Multiplier (cap basis)
    Dim k As Long, a As Long, outArr() As Variant, nOut As Long
    Dim C As Double, cCap As Double, addU As Double, addM As Double
    Dim rcU As Double, rcM As Double, mU As Double, mM As Double
    Dim pfeU As Double, pfeM As Double, eadU As Double, eadM As Double, ead As Double
    Dim tot(1 To RS_NCOLS) As Double

    ReDim outArr(1 To mNSCount + 1, 1 To RS_NCOLS)
    nOut = 0
    For k = 1 To mNSCount
        If mNS(k).Trades = 0 Then
            LogMsg SEV_INFO, SH_NS, mNS(k).ID, "Netting set has no valid trades - not reported."
            GoTo NextNS
        End If
        With mNS(k)
            C = .VM + .NICA
            addU = 0#: addM = 0#
            For a = 1 To AC_COUNT
                addU = addU + mAddOnU(k, a)
                addM = addM + mAddOnM(k, a)
            Next a

            ' unmargined calculation = EAD of an unmargined netting set, or the cap
            If .Margined And .IsCRR Then
                cCap = .NICA            ' Art. 274(3), 275(1); EBA Q&A 2023_6962
            Else
                cCap = C                ' CRE52.2 (posted VM enters C with a negative sign)
            End If
            rcU = SACCR_ReplacementCost(.V, cCap, False)
            mU = SACCR_Multiplier(.V - cCap, addU, pFloor)
            pfeU = mU * addU
            eadU = .Alpha * (rcU + pfeU)

            If .Margined Then
                rcM = SACCR_ReplacementCost(.V, C, True, .TH, .MTA, .NICA)
                mM = SACCR_Multiplier(.V - C, addM, pFloor)
                pfeM = mM * addM
                eadM = .Alpha * (rcM + pfeM)
                ead = Min2(eadM, eadU)
            Else
                ead = eadU
            End If
            mTotalEAD = mTotalEAD + ead

            nOut = nOut + 1
            outArr(nOut, 1) = .ID
            outArr(nOut, 2) = .Counterparty
            outArr(nOut, 3) = .Regime
            outArr(nOut, 4) = IIf(.Margined, "Y", "N")
            outArr(nOut, 5) = .Trades
            If .Margined Then
                outArr(nOut, 6) = .MPOR
            End If
            outArr(nOut, 7) = .V
            outArr(nOut, 8) = C
            For a = 1 To AC_COUNT
                If .Margined Then
                    outArr(nOut, 9 + a) = mAddOnM(k, a)
                Else
                    outArr(nOut, 9 + a) = mAddOnU(k, a)
                End If
            Next a
            If .Margined Then
                outArr(nOut, 9) = rcM
                outArr(nOut, 16) = addM
                outArr(nOut, 17) = mM
                outArr(nOut, 18) = pfeM
                outArr(nOut, 20) = eadM
                outArr(nOut, 21) = eadU
                outArr(nOut, 23) = IIf(eadU < eadM, "Y", "N")
                outArr(nOut, 24) = cCap
                outArr(nOut, 25) = rcU
                outArr(nOut, 26) = addU
                outArr(nOut, 27) = mU
            Else
                outArr(nOut, 9) = rcU
                outArr(nOut, 16) = addU
                outArr(nOut, 17) = mU
                outArr(nOut, 18) = pfeU
                outArr(nOut, 23) = "n/a"
            End If
            outArr(nOut, 19) = .Alpha
            outArr(nOut, 22) = ead
            tot(7) = tot(7) + .V
            tot(8) = tot(8) + C
            For a = 9 To 16
                tot(a) = tot(a) + ToDbl(outArr(nOut, a))
            Next a
            tot(18) = tot(18) + ToDbl(outArr(nOut, 18))
            tot(22) = tot(22) + ead
        End With
NextNS:
    Next k

    If writeOutputs Then
        Dim ws As Worksheet
        Set ws = GetSheet(SH_RESULTS)
        ClearOutputBlock ws, FIRST_DATA_ROW, RS_NCOLS
        If nOut > 0 Then
            nOut = nOut + 1
            outArr(nOut, 1) = "TOTAL"
            For a = 7 To 18
                If a <> 17 Then
                    outArr(nOut, a) = tot(a)
                End If
            Next a
            outArr(nOut, 22) = tot(22)
            WriteBlock ws, FIRST_DATA_ROW, outArr, nOut, RS_NCOLS
            FormatColumns ws, FIRST_DATA_ROW, nOut, _
                "@|@|@|@|0|0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|0.0000|#,##0|0.00|#,##0|#,##0|#,##0|@|#,##0|#,##0|#,##0|0.0000"
            ws.Range(ws.Cells(FIRST_DATA_ROW + nOut - 1, 1), ws.Cells(FIRST_DATA_ROW + nOut - 1, RS_NCOLS)).Font.Bold = True
        End If
    End If
End Sub

'==============================================================================
' Output writers
'==============================================================================
Private Sub WriteBuckets()
    Dim ws As Worksheet, arr() As Variant, b As Long
    Set ws = GetSheet(SH_BUCKETS)
    ClearOutputBlock ws, FIRST_DATA_ROW, BK_NCOLS
    If mBkCount = 0 Then Exit Sub
    ReDim arr(1 To mBkCount, 1 To BK_NCOLS)
    For b = 1 To mBkCount
        arr(b, 1) = mNS(mBk(b).NSIdx).ID
        arr(b, 2) = ACCode(mBk(b).AC)
        arr(b, 3) = mBk(b).HSKey
        arr(b, 4) = mBk(b).SubKey
        arr(b, 5) = mBk(b).Trades
        arr(b, 6) = mBk(b).SF
        If IsFactorClass(mBk(b).AC) Then arr(b, 7) = mBk(b).Corr
        arr(b, 8) = mBk(b).ENU
        arr(b, 9) = IIf(mNS(mBk(b).NSIdx).Margined, mBk(b).ENM, Empty)
        If IsFactorClass(mBk(b).AC) Then
            arr(b, 10) = mBk(b).SF * mBk(b).ENU
            If mNS(mBk(b).NSIdx).Margined Then arr(b, 11) = mBk(b).SF * mBk(b).ENM
        End If
    Next b
    WriteBlock ws, FIRST_DATA_ROW, arr, mBkCount, BK_NCOLS
    FormatColumns ws, FIRST_DATA_ROW, mBkCount, "@|@|@|@|0|0.00%|0%|#,##0|#,##0|#,##0|#,##0"
End Sub

Private Sub WriteHedgingSets()
    Dim ws As Worksheet, arr() As Variant, h As Long, mg As Boolean
    Set ws = GetSheet(SH_HEDGING)
    ClearOutputBlock ws, FIRST_DATA_ROW, HS_NCOLS
    If mHSCount = 0 Then Exit Sub
    ReDim arr(1 To mHSCount, 1 To HS_NCOLS)
    For h = 1 To mHSCount
        mg = mNS(mHS(h).NSIdx).Margined
        arr(h, 1) = mNS(mHS(h).NSIdx).ID
        arr(h, 2) = ACCode(mHS(h).AC)
        arr(h, 3) = mHS(h).Key
        arr(h, 4) = mHS(h).Trades
        If mHS(h).AC = AC_IR Then
            If mg Then
                arr(h, 5) = mHS(h).D1M: arr(h, 6) = mHS(h).D2M: arr(h, 7) = mHS(h).D3M
            Else
                arr(h, 5) = mHS(h).D1U: arr(h, 6) = mHS(h).D2U: arr(h, 7) = mHS(h).D3U
            End If
        End If
        If Not IsFactorClass(mHS(h).AC) Then
            If mg Then arr(h, 8) = mHS(h).ENM Else arr(h, 8) = mHS(h).ENU
        Else
            If mg Then
                arr(h, 9) = mHS(h).SysM: arr(h, 10) = Sqr(mHS(h).IdioM)
            Else
                arr(h, 9) = mHS(h).SysU: arr(h, 10) = Sqr(mHS(h).IdioU)
            End If
        End If
        If mg Then arr(h, 11) = mHS(h).AddOnM Else arr(h, 11) = mHS(h).AddOnU
        arr(h, 12) = mHS(h).AddOnU
        arr(h, 13) = IIf(mg, "Margined", "Unmargined")
        If Not IsFactorClass(mHS(h).AC) Then
            arr(h, 14) = BucketSF(h)
        End If
    Next h
    WriteBlock ws, FIRST_DATA_ROW, arr, mHSCount, HS_NCOLS
    FormatColumns ws, FIRST_DATA_ROW, mHSCount, "@|@|@|0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|#,##0|@|0.00%"
End Sub

Private Sub WriteChecks()
    Dim ws As Worksheet, arr() As Variant, i As Long, j As Long
    On Error GoTo Done
    Set ws = GetSheet(SH_CHECKS)
    ClearOutputBlock ws, FIRST_DATA_ROW, CK_NCOLS
    If mLogCount = 0 Then
        ReDim arr(1 To 1, 1 To CK_NCOLS)
        arr(1, 1) = SEV_INFO
        arr(1, 5) = "No issues found."
        WriteBlock ws, FIRST_DATA_ROW, arr, 1, CK_NCOLS
        Exit Sub
    End If
    ReDim arr(1 To mLogCount, 1 To CK_NCOLS)
    For i = 1 To mLogCount
        For j = 1 To CK_NCOLS
            arr(i, j) = mLog(j, i)
        Next j
    Next i
    WriteBlock ws, FIRST_DATA_ROW, arr, mLogCount, CK_NCOLS
    For i = 1 To mLogCount
        If arr(i, 1) = SEV_ERROR Then
            ws.Cells(FIRST_DATA_ROW + i - 1, 1).Font.Bold = True
        End If
    Next i
Done:
End Sub

Private Sub WriteRunInfo(ByVal secs As Double, ByVal completed As Boolean)
    Dim ws As Worksheet, txt As String
    On Error Resume Next
    Set ws = GetSheet(SH_RESULTS)
    If completed Then
        txt = "Last run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
              " | reporting date " & Format$(CDate(pAsOf), "yyyy-mm-dd") & _
              " | ccy " & pRepCcy & _
              " | trades used " & mTradesUsed & " of " & mTradesRead & _
              " | errors " & mErrCount & ", warnings " & mWarnCount & _
              " | " & Format$(secs, "0.00") & " s"
    Else
        txt = "Last run " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
              " FAILED - see Checks sheet (" & mErrCount & " errors)"
    End If
    ws.Range(RUNINFO_CELL).Value = txt
End Sub
