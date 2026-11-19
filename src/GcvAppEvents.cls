VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GcvAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Excel-wide sheet and workbook switches: only used to enable or grey out the
' student buttons. There is deliberately no SheetChange handler - the lookup
' runs on demand, so typing keeps Ctrl+Z and never sends requests.

Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Public Sub Attach()
    Set App = Application
End Sub

Public Sub Detach()
    Set App = Nothing
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    LookupUi.RefreshRibbon
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    LookupUi.RefreshRibbon
End Sub

Private Sub App_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    LookupUi.RefreshRibbon
End Sub
