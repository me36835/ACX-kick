Option Compare Database
Option Explicit

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