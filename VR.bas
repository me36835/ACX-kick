Option Compare Database
Option Explicit

Sub distriNew() ' $00:00007
Dim RST As DAO.Recordset
Dim tmp As String
Dim htm As String
Dim SQL As String
Dim s() As String
Dim i As Long

GL.envi "reset"

VRSI = Right("0000" & (Val(VRSI) + 1), 4)

dum = SQ.UPD("UPDATE SETT SET TXT = 'Code Version " & VRSI & "', DAT = NOW() WHERE LUP = 'Version'")

If dum = 0 Then
    dum = SQ.UPD("INSERT INTO SETT (TXT, DAT, LUP) VALUES ('Code Version " & VRSI & "', NOW(), 'Version')")
End If

DB.Properties("Versionsnummer") = VRSI
DB.Properties.Refresh

VR.AlleFormulareSpeichern
VR.AlleModuleSpeichern

Call GT.write2git ' IDX00051

End Sub


Sub AlleFormulareSpeichern()
Dim frm As Form
For Each frm In Forms ' err
    On Error Resume Next ' Fehlerbehandlung, falls ein Formular nicht speicherbar ist
    frm.Save
    On Error GoTo 0
Next frm
End Sub

Sub AlleModuleSpeichern()
Dim obj As AccessObject

On Error Resume Next ' Fehler ignorieren, falls ein Modul aus irgendeinem Grund nicht gespeichert werden kann
    
' Schleife durch alle Module im aktuellen Projekt
For Each obj In Application.CurrentProject.AllModules
    DoCmd.Save acModule, obj.Name
Next obj

For Each obj In Application.CurrentProject.AllClasses
    DoCmd.Save acModule, obj.Name ' acModule funktioniert auch f?r Klassenmodule
Next obj

On Error GoTo 0

End Sub