VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StudentLookupResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Result of one lookup request.

Public Outcome As LookupOutcome
Public HttpStatus As Long
Public Body As String
Public ErrorDescription As String
