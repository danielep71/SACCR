Attribute VB_Name = "M_Config"
'==============================================================================
' Module   : M_Config
' Purpose  : Layout constants (sheet names, column positions, parameter codes)
'            shared by the SA-CCR engine. Change layout here only.
'==============================================================================
Option Explicit
Option Private Module
Option Private Module

'--- Sheet names --------------------------------------------------------------
Public Const SH_README As String = "README"
Public Const SH_PARAMS As String = "Params"
Public Const SH_NS As String = "NettingSets"
Public Const SH_TRADES As String = "Trades"
Public Const SH_RESULTS As String = "Results"
Public Const SH_HEDGING As String = "HedgingSets"
Public Const SH_BUCKETS As String = "Buckets"
Public Const SH_TRADECALC As String = "TradeCalc"
Public Const SH_CHECKS As String = "Checks"

'--- Common layout ------------------------------------------------------------
Public Const HEADER_ROW As Long = 4
Public Const FIRST_DATA_ROW As Long = 5
Public Const RUNINFO_CELL As String = "A2"

'--- Params sheet: code in column A, value in column C ------------------------
Public Const PRM_CODE_COL As Long = 1
Public Const PRM_VALUE_COL As Long = 3
Public Const PRM_ASOF As String = "AsOfDate"
Public Const PRM_REPCCY As String = "ReportingCcy"
Public Const PRM_ALPHA As String = "Alpha"
Public Const PRM_FLOOR As String = "MultiplierFloor"
Public Const PRM_DAYSYEAR As String = "DaysPerYear"
Public Const PRM_BDYEAR As String = "BusinessDaysPerYear"
Public Const PRM_MINMAT As String = "MinMaturityBD"
Public Const PRM_SDFLOOR As String = "SDFloorBD"
Public Const PRM_MPOR_BIL As String = "MPORFloorBilateral"
Public Const PRM_MPOR_CLR As String = "MPORFloorCleared"
Public Const PRM_MPOR_LARGE As String = "MPORFloorLarge"
Public Const PRM_BASIS As String = "BasisFactor"
Public Const PRM_VOLF As String = "VolatilityFactor"
Public Const PRM_RHO12 As String = "IRCorrBucket12"
Public Const PRM_RHO23 As String = "IRCorrBucket23"
Public Const PRM_RHO13 As String = "IRCorrBucket13"
Public Const PRM_IRFULL As String = "IRBucketOffset"
Public Const PRM_REGIME As String = "Regime"
Public Const PRM_LAMIR As String = "LambdaThresholdIR"
Public Const PRM_LAMCO As String = "LambdaThresholdCO"

'--- Regimes -------------------------------------------------------------------
Public Const RG_BCBS As String = "BCBS"
Public Const RG_CRR As String = "CRR"

'--- Table headers on Params (searched in column A) ---------------------------
Public Const HDR_SF As String = "SF_Key"
Public Const HDR_FX As String = "FX_Ccy"

'--- NettingSets input columns ------------------------------------------------
Public Const NS_ID As Long = 1
Public Const NS_CPTY As Long = 2
Public Const NS_MARGINED As Long = 3
Public Const NS_CLEARED As Long = 4
Public Const NS_FREQ As Long = 5
Public Const NS_LARGE As Long = 6
Public Const NS_DISPUTE As Long = 7
Public Const NS_MPOR As Long = 8
Public Const NS_VM As Long = 9
Public Const NS_NICA As Long = 10
Public Const NS_TH As Long = 11
Public Const NS_MTA As Long = 12
Public Const NS_ALPHA As Long = 13
Public Const NS_REGIME As Long = 14
Public Const NS_NCOLS As Long = 14

'--- Trades input columns -----------------------------------------------------
Public Const TR_ID As Long = 1
Public Const TR_NS As Long = 2
Public Const TR_AC As Long = 3
Public Const TR_SUB As Long = 4
Public Const TR_RF As Long = 5
Public Const TR_INSTR As Long = 6
Public Const TR_DIR As Long = 7
Public Const TR_OPT As Long = 8
Public Const TR_NATURE As Long = 9
Public Const TR_LABEL As Long = 10
Public Const TR_NOTIONAL As Long = 11
Public Const TR_NCCY As Long = 12
Public Const TR_MTM As Long = 13
Public Const TR_MCCY As Long = 14
Public Const TR_START As Long = 15
Public Const TR_END As Long = 16
Public Const TR_MAT As Long = 17
Public Const TR_EXPIRY As Long = 18
Public Const TR_PRICE As Long = 19
Public Const TR_STRIKE As Long = 20
Public Const TR_LAMBDA As Long = 21
Public Const TR_ATTACH As Long = 22
Public Const TR_DETACH As Long = 23
Public Const TR_COMMENT As Long = 24
Public Const TR_NCOLS As Long = 24

'--- Output widths -------------------------------------------------------------
Public Const TC_NCOLS As Long = 25
Public Const BK_NCOLS As Long = 11
Public Const HS_NCOLS As Long = 14
Public Const RS_NCOLS As Long = 27
Public Const CK_NCOLS As Long = 5

'--- Asset classes -------------------------------------------------------------
Public Const AC_IR As Long = 1
Public Const AC_FX As Long = 2
Public Const AC_CR As Long = 3
Public Const AC_EQ As Long = 4
Public Const AC_CO As Long = 5
Public Const AC_OT As Long = 6          ' other risks - CRR Art. 277(1)(f) / 280f only
Public Const AC_COUNT As Long = 6

'--- Severity labels -----------------------------------------------------------
Public Const SEV_ERROR As String = "ERROR"
Public Const SEV_WARN As String = "WARNING"
Public Const SEV_INFO As String = "INFO"
