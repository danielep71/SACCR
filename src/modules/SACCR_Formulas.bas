Attribute VB_Name = "SACCR_Formulas"
'==============================================================================
' MODULE: SACCR_Formulas
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the SA-CCR building blocks as public functions. The engine calls
'   them, and they can also be entered in worksheet cells as user-defined
'   functions, so every intermediate figure can be checked independently.
'
' PUBLIC SURFACE
'   The ten SACCR_* functions, listed in docs/PUBLIC_API.txt. Their names,
'   arguments, defaults and results are a contract: changing one is a public
'   API change.
'
' DEPENDENCIES
'   VBA runtime only. SACCR_NormCDF is implemented here so that results do
'   not depend on WorksheetFunction, the Excel version or the locale.
'
' STATE OWNERSHIP
'   Stateless. Results depend only on the arguments.
'
' ERROR POLICY
'   Invalid inputs that have no meaningful result return a worksheet error
'   value (#NUM! or #REF!) instead of raising, so that a cell shows the error
'   and the rest of the sheet still calculates. Functions declared As Double
'   apply floors instead of rejecting inputs.
'
' REFERENCE
'   BCBS CRE52 (paragraph numbers in brackets); EU CRR Art. 274 to 280f and
'   Commission Delegated Regulation (EU) 2021/931 as amended.
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
    'Require explicit declarations. This module is public: its functions are
    'available to worksheet cells.
    Option Explicit


'
'------------------------------------------------------------------------------
'
'                             STATISTICAL FUNCTION
'
'------------------------------------------------------------------------------
'

Public Function SACCR_NormCDF( _
    ByVal x As Double) _
    As Double
'
'==============================================================================
'                                SACCR_NormCDF
'------------------------------------------------------------------------------
' PURPOSE
'   Standard normal cumulative distribution function N(x).
'
' INPUTS
'   x: any real number.
'
' RETURNS
'   N(x), between 0 and 1. Absolute error below 1E-14.
'
' REFERENCE
'   Hart (1968) double-precision algorithm, as published by G. West (2005),
'   "Better approximations to cumulative normal functions".
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim xAbs   As Double    '|x|
    Dim ex     As Double    'exp(-x^2 / 2)
    Dim num    As Double    'Numerator polynomial, central region
    Dim den    As Double    'Denominator polynomial, central region
    Dim b      As Double    'Continued fraction, tail region

'------------------------------------------------------------------------------
' LOWER-TAIL PROBABILITY OF |x|
'------------------------------------------------------------------------------
    'Compute N(-|x|). Beyond 37 it is below 1E-298 and is taken as 0. Up to
    '7.07 (10 / sqrt(2)) a rational approximation is used; beyond, a
    'continued fraction.
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

'------------------------------------------------------------------------------
' REFLECT FOR POSITIVE x
'------------------------------------------------------------------------------
    'By symmetry, N(x) = 1 - N(-x).
        If x > 0# Then
            SACCR_NormCDF = 1# - SACCR_NormCDF
        End If

End Function


'
'------------------------------------------------------------------------------
'
'                          TRADE-LEVEL BUILDING BLOCKS
'
'------------------------------------------------------------------------------
'

Public Function SACCR_SupervisoryDuration( _
    ByVal S As Double, _
    ByVal E As Double, _
    Optional ByVal FloorYears As Double = 0#) _
    As Double
'
'==============================================================================
'                          SACCR_SupervisoryDuration
'------------------------------------------------------------------------------
' PURPOSE
'   Supervisory duration of an interest-rate or credit trade:
'     SD = (exp(-0.05 * S) - exp(-0.05 * E)) / 0.05
'
' INPUTS
'   S: start of the trade in years from the reporting date; a negative value
'      (already started) is floored at 0.
'   E: end of the trade in years from the reporting date.
'   FloorYears: lower bound on SD in years; 0 (no floor) by default. Ten
'      business days is 10 / 250.
'
' RETURNS
'   SD in years, at least FloorYears.
'
' REFERENCE
'   CRE52.34; CRR Art. 279b.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim sd   As Double    'Supervisory duration, years

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        If S < 0# Then
            S = 0#
        End If
        sd = (Exp(-0.05 * S) - Exp(-0.05 * E)) / 0.05
        If sd < FloorYears Then
            sd = FloorYears
        End If
        SACCR_SupervisoryDuration = sd

End Function


Public Function SACCR_MaturityFactor( _
    ByVal M As Double, _
    Optional ByVal Margined As Boolean = False, _
    Optional ByVal MPOR_BD As Double = 10#, _
    Optional ByVal MinBD As Double = 10#, _
    Optional ByVal BDPerYear As Double = 250#) _
    As Double
'
'==============================================================================
'                             SACCR_MaturityFactor
'------------------------------------------------------------------------------
' PURPOSE
'   Maturity factor MF of a trade:
'     unmargined: MF = sqrt(min(max(M, MinBD / BDPerYear), 1))
'     margined:   MF = 1.5 * sqrt(MPOR_BD / BDPerYear)
'
' INPUTS
'   M: remaining maturity in years; ignored when Margined is True.
'   Margined: True for a margined netting set.
'   MPOR_BD: margin period of risk in business days; 10 by default.
'   MinBD: floor on M in business days; 10 by default.
'   BDPerYear: business days per year; 250 by default; must be > 0.
'
' RETURNS
'   The maturity factor, a positive number.
'
' REFERENCE
'   CRE52.48 (unmargined); CRE52.52 (margined).
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim mEff   As Double    'Maturity after the floor and the one-year cap

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        If Margined Then
            SACCR_MaturityFactor = 1.5 * Sqr(MPOR_BD / BDPerYear)
        Else
            mEff = M
            If mEff < MinBD / BDPerYear Then
                mEff = MinBD / BDPerYear
            End If
            If mEff > 1# Then
                mEff = 1#
            End If
            SACCR_MaturityFactor = Sqr(mEff)
        End If

End Function


Public Function SACCR_OptionDelta( _
    ByVal P As Double, _
    ByVal K As Double, _
    ByVal T As Double, _
    ByVal Vol As Double, _
    ByVal IsCall As Boolean, _
    ByVal IsBought As Boolean, _
    Optional ByVal Lambda As Double = 0#) _
    As Variant
'
'==============================================================================
'                              SACCR_OptionDelta
'------------------------------------------------------------------------------
' PURPOSE
'   Supervisory delta of an option:
'     d1 = (ln((P + L) / (K + L)) + 0.5 * Vol^2 * T) / (Vol * sqrt(T))
'     bought call +N(d1); sold call -N(d1); bought put -N(-d1); sold put +N(-d1)
'
' INPUTS
'   P: price of the underlying; P + Lambda > 0.
'   K: strike price; K + Lambda > 0.
'   T: time to the latest exercise date in years; T > 0.
'   Vol: supervisory option volatility as a decimal; Vol > 0.
'   IsCall: True for a call, False for a put.
'   IsBought: True for a bought (long) option, False for a sold one.
'   Lambda: shift L for negative prices or rates; 0 by default.
'
' RETURNS
'   The delta, between -1 and 1; #NUM! when an input is outside its domain.
'
' REFERENCE
'   CRE52.40; CRR Art. 279a; lambda per Delegated Regulation (EU) 2021/931.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim d1    As Double    'Black-Scholes d1 with the lambda shift
    Dim dlt   As Double    'Delta before the sign of the position

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'The logarithm and the square root need strictly positive arguments.
        If T <= 0# Or Vol <= 0# Or (P + Lambda) <= 0# Or (K + Lambda) <= 0# Then
            SACCR_OptionDelta = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Delta of a bought option first; a sold option takes the opposite sign.
        d1 = (Log((P + Lambda) / (K + Lambda)) + 0.5 * Vol * Vol * T) / (Vol * Sqr(T))
        If IsCall Then
            dlt = SACCR_NormCDF(d1)
        Else
            dlt = -SACCR_NormCDF(-d1)
        End If
        If Not IsBought Then
            dlt = -dlt
        End If
        SACCR_OptionDelta = dlt

End Function


Public Function SACCR_LambdaCRR( _
    ByVal P As Double, _
    ByVal K As Double, _
    ByVal IsInterestRate As Boolean, _
    Optional ByVal Threshold As Variant) _
    As Double
'
'==============================================================================
'                               SACCR_LambdaCRR
'------------------------------------------------------------------------------
' PURPOSE
'   Lambda shift of an option under the EU rules, determined per option:
'     interest rate: lambda = max(threshold - min(P, K), 0); threshold 0.10%
'     commodity:     lambda = max(-(1 + threshold) * min(P, K), 0); threshold 0.1
'   Under Basel the shift is left to judgement (CRE52.40 FAQ) and is entered
'   on the trade instead.
'
' INPUTS
'   P: price of the underlying.
'   K: strike price.
'   IsInterestRate: True for an option mapped to interest rate, False for one
'      mapped to commodity.
'   Threshold: optional threshold as a decimal; when omitted, 0.001 for
'      interest rate and 0.1 for commodity.
'
' RETURNS
'   Lambda, zero or positive.
'
' REFERENCE
'   Commission Delegated Regulation (EU) 2021/931 Art. 5, as amended by
'   Commission Delegated Regulation (EU) 2025/855.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim thr      As Double    'Threshold applied
    Dim lowest   As Double    'min(P, K)
    Dim lam      As Double    'Lambda before the floor at zero

'------------------------------------------------------------------------------
' TAKE THE LOWER OF PRICE AND STRIKE
'------------------------------------------------------------------------------
        lowest = P
        If K < lowest Then
            lowest = K
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
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


Public Function SACCR_CDODelta( _
    ByVal A As Double, _
    ByVal D As Double, _
    ByVal IsLong As Boolean) _
    As Variant
'
'==============================================================================
'                                SACCR_CDODelta
'------------------------------------------------------------------------------
' PURPOSE
'   Supervisory delta of a CDO tranche:
'     delta = +/- 15 / ((1 + 14 * A) * (1 + 14 * D))
'   positive for purchased protection (long), negative for sold protection.
'
' INPUTS
'   A: attachment point as a decimal; A >= 0.
'   D: detachment point as a decimal; A < D <= 1.
'   IsLong: True for purchased protection.
'
' RETURNS
'   The delta; #NUM! when the tranche is not valid.
'
' REFERENCE
'   CRE52.41.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        If A < 0# Or D > 1# Or A >= D Then
            SACCR_CDODelta = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        SACCR_CDODelta = IIf(IsLong, 1#, -1#) * 15# / ((1# + 14# * A) * (1# + 14# * D))

End Function


'
'------------------------------------------------------------------------------
'
'                        HEDGING-SET AND NETTING-SET FORMULAS
'
'------------------------------------------------------------------------------
'

Public Function SACCR_IREffectiveNotional( _
    ByVal D1 As Double, _
    ByVal D2 As Double, _
    ByVal D3 As Double, _
    Optional ByVal Rho12 As Double = 0.7, _
    Optional ByVal Rho23 As Double = 0.7, _
    Optional ByVal Rho13 As Double = 0.3) _
    As Double
'
'==============================================================================
'                          SACCR_IREffectiveNotional
'------------------------------------------------------------------------------
' PURPOSE
'   Effective notional of an interest-rate hedging set across the three
'   maturity buckets:
'     EN = sqrt(D1^2 + D2^2 + D3^2 + 2*r12*D1*D2 + 2*r23*D2*D3 + 2*r13*D1*D3)
'   The default correlations give the regulatory cross terms 1.4, 1.4, 0.6.
'
' INPUTS
'   D1, D2, D3: signed effective notionals of the buckets under one year,
'      one to five years and over five years.
'   Rho12, Rho23, Rho13: bucket correlations as decimals.
'
' RETURNS
'   The effective notional, zero or positive.
'
' REFERENCE
'   CRE52.57.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim q   As Double    'Quadratic form under the square root

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Correlations entered by the user can make the quadratic form negative;
    'it is floored at zero before the square root.
        q = D1 * D1 + D2 * D2 + D3 * D3 _
            + 2# * Rho12 * D1 * D2 + 2# * Rho23 * D2 * D3 + 2# * Rho13 * D1 * D3
        If q < 0# Then
            q = 0#
        End If
        SACCR_IREffectiveNotional = Sqr(q)

End Function


Public Function SACCR_Multiplier( _
    ByVal VminusC As Double, _
    ByVal AddOn As Double, _
    Optional ByVal FloorPct As Double = 0.05) _
    As Double
'
'==============================================================================
'                               SACCR_Multiplier
'------------------------------------------------------------------------------
' PURPOSE
'   PFE multiplier of a netting set:
'     m = min(1, F + (1 - F) * exp((V - C) / (2 * (1 - F) * AddOn)))
'
' INPUTS
'   VminusC: V - C, value of the trades less the collateral held.
'   AddOn: aggregate add-on of the netting set; when 0 or less, m is 1.
'   FloorPct: floor F as a decimal; 0.05 by default; F < 1.
'
' RETURNS
'   The multiplier, between FloorPct and 1.
'
' REFERENCE
'   CRE52.23.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim m   As Double    'Multiplier
    Dim z   As Double    'Exponent of the formula

'------------------------------------------------------------------------------
' HANDLE A ZERO ADD-ON
'------------------------------------------------------------------------------
        If AddOn <= 0# Then
            SACCR_Multiplier = 1#
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'A positive exponent always gives 1 after the cap. Below -700, Exp
    'underflows, so the result is the floor.
        z = VminusC / (2# * (1# - FloorPct) * AddOn)
        If z > 0# Then
            m = 1#
        ElseIf z < -700# Then
            m = FloorPct
        Else
            m = FloorPct + (1# - FloorPct) * Exp(z)
        End If
        If m > 1# Then
            m = 1#
        End If
        SACCR_Multiplier = m

End Function


Public Function SACCR_ReplacementCost( _
    ByVal V As Double, _
    ByVal C As Double, _
    Optional ByVal Margined As Boolean = False, _
    Optional ByVal TH As Double = 0#, _
    Optional ByVal MTA As Double = 0#, _
    Optional ByVal NICA As Double = 0#) _
    As Double
'
'==============================================================================
'                            SACCR_ReplacementCost
'------------------------------------------------------------------------------
' PURPOSE
'   Replacement cost RC of a netting set:
'     unmargined: RC = max(V - C, 0)
'     margined:   RC = max(V - C, TH + MTA - NICA, 0)
'
' INPUTS
'   V: value of the trades in the netting set.
'   C: haircut value of the net collateral held.
'   Margined: True for a margined netting set.
'   TH: threshold; MTA: minimum transfer amount; NICA: net independent
'      collateral amount. Used only when Margined is True.
'
' RETURNS
'   RC, zero or positive.
'
' REFERENCE
'   CRE52.10 (unmargined); CRE52.18 (margined).
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim rc   As Double    'Replacement cost

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        rc = V - C
        If Margined Then
            If TH + MTA - NICA > rc Then
                rc = TH + MTA - NICA
            End If
        End If
        If rc < 0# Then
            rc = 0#
        End If
        SACCR_ReplacementCost = rc

End Function


Public Function SACCR_FactorAggregation( _
    ByVal AddOns As Variant, _
    ByVal Rhos As Variant) _
    As Variant
'
'==============================================================================
'                           SACCR_FactorAggregation
'------------------------------------------------------------------------------
' PURPOSE
'   Single-factor aggregation used for credit, equity and commodity:
'     AddOn = sqrt((sum rho_k * A_k)^2 + sum (1 - rho_k^2) * A_k^2)
'
' INPUTS
'   AddOns: the entity or commodity-type add-ons A_k; a range, an array or a
'      single number.
'   Rhos: the correlations rho_k as decimals; same size as AddOns.
'
' RETURNS
'   The aggregated add-on; #REF! when AddOns and Rhos differ in size. Blank
'   or non-numeric add-ons are skipped.
'
' REFERENCE
'   CRE52.61 (credit), CRE52.66 (equity), CRE52.70 (commodity).
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim a      As Variant    'Copy of AddOns
    Dim r      As Variant    'Copy of Rhos
    Dim i      As Long       'Element counter
    Dim sys    As Double     'Systematic part: sum rho_k * A_k
    Dim idio   As Double     'Idiosyncratic part: sum (1 - rho_k^2) * A_k^2
    Dim av     As Double     'Add-on A_k
    Dim rv     As Double     'Correlation rho_k
    Dim va     As Variant    'Add-ons flattened to a 1-based list
    Dim vr     As Variant    'Correlations flattened to a 1-based list
    Dim x      As Variant    'Element being copied

'------------------------------------------------------------------------------
' HANDLE A SINGLE VALUE
'------------------------------------------------------------------------------
        a = AddOns
        r = Rhos
        If Not IsArray(a) Then
            av = CDbl(a)
            rv = CDbl(r)
            SACCR_FactorAggregation = Sqr((rv * av) ^ 2 + (1# - rv * rv) * av * av)
            Exit Function
        End If

'------------------------------------------------------------------------------
' FLATTEN BOTH INPUTS
'------------------------------------------------------------------------------
    'Ranges arrive as 2-D arrays of any shape. For Each visits every element,
    'so both inputs become 1-based lists in the same order.
        ReDim va(1 To 1)
        ReDim vr(1 To 1)
        i = 0
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

'------------------------------------------------------------------------------
' AGGREGATE
'------------------------------------------------------------------------------
        For i = 1 To UBound(va)
            If IsNumeric(va(i)) And Not IsEmpty(va(i)) Then
                av = CDbl(va(i))
                rv = CDbl(vr(i))
                sys = sys + rv * av
                idio = idio + (1# - rv * rv) * av * av
            End If
        Next i
        SACCR_FactorAggregation = Sqr(sys * sys + idio)

End Function
