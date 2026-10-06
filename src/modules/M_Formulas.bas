Attribute VB_Name = "M_Formulas"
'==============================================================================
' Module   : M_Formulas
' Purpose  : SA-CCR building blocks as public functions. They are used by the
'            engine and can also be called from worksheet cells (UDFs), which
'            makes every intermediate figure independently checkable.
' Reference: BCBS CRE52 (paragraph numbers in brackets); EU CRR Art. 274-280f
'            and Delegated Regulation (EU) 2021/931 as amended.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Standard normal cumulative distribution function.
' Hart (1968) double-precision algorithm as published by G. West (2005),
' absolute error < 1E-14. Implemented locally so that the engine does not
' depend on WorksheetFunction (version / locale independent).
'------------------------------------------------------------------------------
Public Function SACCR_NormCDF(ByVal x As Double) As Double
    Dim xAbs As Double, ex As Double, num As Double, den As Double, b As Double
    xAbs = Abs(x)
    If xAbs > 37# Then
        SACCR_NormCDF = 0#
    Else
        ex = Exp(-xAbs * xAbs / 2#)
        If xAbs < 7.07106781186547 Then
            num = 3.52624965998911E-02 * xAbs + 0.700383064443688
            num = num * xAbs + 6.37396220353165
            num = num * xAbs + 33.912866078383
            num = num * xAbs + 112.079291497871
            num = num * xAbs + 221.213596169931
            num = num * xAbs + 220.206867912376
            den = 8.83883476483184E-02 * xAbs + 1.75566716318264
            den = den * xAbs + 16.064177579207
            den = den * xAbs + 86.7807322029461
            den = den * xAbs + 296.564248779674
            den = den * xAbs + 637.333633378831
            den = den * xAbs + 793.826512519948
            den = den * xAbs + 440.413735824752
            SACCR_NormCDF = ex * num / den
        Else
            b = xAbs + 0.65
            b = xAbs + 4# / b
            b = xAbs + 3# / b
            b = xAbs + 2# / b
            b = xAbs + 1# / b
            SACCR_NormCDF = ex / b / 2.506628274631
        End If
    End If
    If x > 0# Then SACCR_NormCDF = 1# - SACCR_NormCDF
End Function

'------------------------------------------------------------------------------
' Supervisory duration for IR and credit trades [CRE52.34]:
'   SD = (exp(-0.05*S) - exp(-0.05*E)) / 0.05
' S, E in years from the reporting date (S floored at 0).
' FloorYears: floor on SD (CRE52.34 / CRR Art. 279b: 10 business days = 10/250).
'------------------------------------------------------------------------------
Public Function SACCR_SupervisoryDuration(ByVal S As Double, ByVal E As Double, _
                                          Optional ByVal FloorYears As Double = 0#) As Double
    Dim sd As Double
    If S < 0# Then S = 0#
    sd = (Exp(-0.05 * S) - Exp(-0.05 * E)) / 0.05
    If sd < FloorYears Then sd = FloorYears
    SACCR_SupervisoryDuration = sd
End Function

'------------------------------------------------------------------------------
' Maturity factor [CRE52.48 unmargined, CRE52.52 margined]
'   unmargined: MF = sqrt( min( max(M, MinBD/BDPerYear), 1 ) / 1 )
'   margined  : MF = 1.5 * sqrt( MPOR / BDPerYear )
'------------------------------------------------------------------------------
Public Function SACCR_MaturityFactor(ByVal M As Double, _
                                     Optional ByVal Margined As Boolean = False, _
                                     Optional ByVal MPOR_BD As Double = 10#, _
                                     Optional ByVal MinBD As Double = 10#, _
                                     Optional ByVal BDPerYear As Double = 250#) As Double
    Dim mEff As Double
    If Margined Then
        SACCR_MaturityFactor = 1.5 * Sqr(MPOR_BD / BDPerYear)
    Else
        mEff = M
        If mEff < MinBD / BDPerYear Then mEff = MinBD / BDPerYear
        If mEff > 1# Then mEff = 1#
        SACCR_MaturityFactor = Sqr(mEff)
    End If
End Function

'------------------------------------------------------------------------------
' Supervisory delta for options [CRE52.40]
'   d1 = ( ln((P+L)/(K+L)) + 0.5*vol^2*T ) / ( vol*sqrt(T) )
'   bought call: +N(d1)   sold call: -N(d1)
'   bought put : -N(-d1)  sold put : +N(-d1)
' Lambda (L) is the shift for negative prices/rates (CRR Art. 279a).
'------------------------------------------------------------------------------
Public Function SACCR_OptionDelta(ByVal P As Double, ByVal K As Double, ByVal T As Double, _
                                  ByVal Vol As Double, ByVal IsCall As Boolean, _
                                  ByVal IsBought As Boolean, _
                                  Optional ByVal Lambda As Double = 0#) As Variant
    Dim d1 As Double, dlt As Double
    If T <= 0# Or Vol <= 0# Or (P + Lambda) <= 0# Or (K + Lambda) <= 0# Then
        SACCR_OptionDelta = CVErr(xlErrNum)
        Exit Function
    End If
    d1 = (Log((P + Lambda) / (K + Lambda)) + 0.5 * Vol * Vol * T) / (Vol * Sqr(T))
    If IsCall Then
        dlt = SACCR_NormCDF(d1)
    Else
        dlt = -SACCR_NormCDF(-d1)
    End If
    If Not IsBought Then dlt = -dlt
    SACCR_OptionDelta = dlt
End Function

'------------------------------------------------------------------------------
' EU lambda shift for options mapped to IR or commodity (Commission Delegated
' Regulation (EU) 2021/931, Art. 5, as amended by Delegated Regulation
' (EU) 2025/855), determined per option j:
'   interest rate : lambda = max( threshold - min(P, K) ; 0 ),  threshold = 0.10%
'   commodity     : lambda = max( -(1 + threshold) * min(P, K) ; 0 ), threshold = 0.1
' Basel (CRE52.40 FAQ) leaves lambda to judgement: enter it on the trade.
'------------------------------------------------------------------------------
Public Function SACCR_LambdaCRR(ByVal P As Double, ByVal K As Double, _
                                ByVal IsInterestRate As Boolean, _
                                Optional ByVal Threshold As Variant) As Double
    Dim thr As Double, lowest As Double, lam As Double
    lowest = P
    If K < lowest Then
        lowest = K
    End If
    If IsInterestRate Then
        If IsMissing(Threshold) Then
            thr = 0.001
        Else
            thr = CDbl(Threshold)
        End If
        lam = thr - lowest
    Else
        If IsMissing(Threshold) Then
            thr = 0.1
        Else
            thr = CDbl(Threshold)
        End If
        lam = -(1# + thr) * lowest
    End If
    If lam < 0# Then
        lam = 0#
    End If
    SACCR_LambdaCRR = lam
End Function

'------------------------------------------------------------------------------
' Supervisory delta for CDO tranches [CRE52.41]
'   delta = +/- 15 / ((1 + 14*A) * (1 + 14*D))   (+ purchased protection)
'------------------------------------------------------------------------------
Public Function SACCR_CDODelta(ByVal A As Double, ByVal D As Double, _
                               ByVal IsLong As Boolean) As Variant
    If A < 0# Or D > 1# Or A >= D Then
        SACCR_CDODelta = CVErr(xlErrNum)
        Exit Function
    End If
    SACCR_CDODelta = IIf(IsLong, 1#, -1#) * 15# / ((1# + 14# * A) * (1# + 14# * D))
End Function

'------------------------------------------------------------------------------
' IR hedging-set effective notional across maturity buckets [CRE52.57]
'   EN = sqrt( D1^2 + D2^2 + D3^2 + 2*r12*D1*D2 + 2*r23*D2*D3 + 2*r13*D1*D3 )
' Defaults reproduce the regulatory 1.4 / 1.4 / 0.6 cross terms.
'------------------------------------------------------------------------------
Public Function SACCR_IREffectiveNotional(ByVal D1 As Double, ByVal D2 As Double, _
                                          ByVal D3 As Double, _
                                          Optional ByVal Rho12 As Double = 0.7, _
                                          Optional ByVal Rho23 As Double = 0.7, _
                                          Optional ByVal Rho13 As Double = 0.3) As Double
    Dim q As Double
    q = D1 * D1 + D2 * D2 + D3 * D3 _
        + 2# * Rho12 * D1 * D2 + 2# * Rho23 * D2 * D3 + 2# * Rho13 * D1 * D3
    If q < 0# Then q = 0#
    SACCR_IREffectiveNotional = Sqr(q)
End Function

'------------------------------------------------------------------------------
' PFE multiplier [CRE52.23]
'   m = min( 1 ; F + (1-F) * exp( (V-C) / (2*(1-F)*AddOn) ) ),  F = 5%
'------------------------------------------------------------------------------
Public Function SACCR_Multiplier(ByVal VminusC As Double, ByVal AddOn As Double, _
                                 Optional ByVal FloorPct As Double = 0.05) As Double
    Dim m As Double, z As Double
    If AddOn <= 0# Then
        SACCR_Multiplier = 1#
        Exit Function
    End If
    z = VminusC / (2# * (1# - FloorPct) * AddOn)
    If z > 0# Then
        m = 1#
    ElseIf z < -700# Then
        m = FloorPct
    Else
        m = FloorPct + (1# - FloorPct) * Exp(z)
    End If
    If m > 1# Then m = 1#
    SACCR_Multiplier = m
End Function

'------------------------------------------------------------------------------
' Replacement cost [CRE52.10 unmargined, CRE52.18 margined]
'   unmargined: RC = max(V - C, 0)
'   margined  : RC = max(V - C, TH + MTA - NICA, 0)
'------------------------------------------------------------------------------
Public Function SACCR_ReplacementCost(ByVal V As Double, ByVal C As Double, _
                                      Optional ByVal Margined As Boolean = False, _
                                      Optional ByVal TH As Double = 0#, _
                                      Optional ByVal MTA As Double = 0#, _
                                      Optional ByVal NICA As Double = 0#) As Double
    Dim rc As Double
    rc = V - C
    If Margined Then
        If TH + MTA - NICA > rc Then rc = TH + MTA - NICA
    End If
    If rc < 0# Then rc = 0#
    SACCR_ReplacementCost = rc
End Function

'------------------------------------------------------------------------------
' Single-factor aggregation used for credit, equity and commodity [CRE52.61,
' 52.66, 52.70]:  AddOn = sqrt( (sum rho_k*A_k)^2 + sum (1-rho_k^2)*A_k^2 )
' AddOns and Rhos are same-sized ranges or arrays.
'------------------------------------------------------------------------------
Public Function SACCR_FactorAggregation(ByVal AddOns As Variant, ByVal Rhos As Variant) As Variant
    Dim a As Variant, r As Variant, i As Long, sys As Double, idio As Double
    Dim av As Double, rv As Double
    a = AddOns
    r = Rhos
    If Not IsArray(a) Then
        av = CDbl(a): rv = CDbl(r)
        SACCR_FactorAggregation = Sqr((rv * av) ^ 2 + (1# - rv * rv) * av * av)
        Exit Function
    End If
    Dim va As Variant, vr As Variant
    ReDim va(1 To 1): ReDim vr(1 To 1)
    i = 0
    Dim x As Variant
    For Each x In a
        i = i + 1
        ReDim Preserve va(1 To i)
        va(i) = x
    Next x
    i = 0
    For Each x In r
        i = i + 1
        ReDim Preserve vr(1 To i)
        vr(i) = x
    Next x
    If UBound(va) <> UBound(vr) Then
        SACCR_FactorAggregation = CVErr(xlErrRef)
        Exit Function
    End If
    For i = 1 To UBound(va)
        If IsNumeric(va(i)) And Not IsEmpty(va(i)) Then
            av = CDbl(va(i)): rv = CDbl(vr(i))
            sys = sys + rv * av
            idio = idio + (1# - rv * rv) * av * av
        End If
    Next i
    SACCR_FactorAggregation = Sqr(sys * sys + idio)
End Function
