VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shResults"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Activate()
'
'==============================================================================
'                              Worksheet_Activate
'------------------------------------------------------------------------------
' PURPOSE
'   Mark the results as out of date when the inputs changed after the last
'   run (#36). M_Main.FlagStaleResults never raises.
'
' UPDATED
'   2026-10-07
'==============================================================================
'
        M_Main.FlagStaleResults

End Sub
