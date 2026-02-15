Option Compare Database
Option Explicit

Public BL As Integer

Public TEA As String
Public SPT As Long
Public SES As String

Public dum As Variant
Public v As Variant


Public DB As Database

Sub initialize()

If DB Is Nothing Then Set DB = CurrentDb

If BL = 0 Then BL = SQ.sel("SELECT BL FROM SETT WHERE LUP='ACTIVE'")
If Len(SES) = 0 Then SES = SQ.sel("SELECT TXT FROM SETT WHERE LUP='ACTIVE'")


End Sub