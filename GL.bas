Option Compare Database
Option Explicit

Public HEUTE As Date
Public REPO As String
Public SUP As String
Public VRSI As String
Public Const IDX As String = "$00:"

' Library Scripting
' C:\WINDOWS\system32\scrrun.dll
' C:\Windows\sysWOW64\scrrun.dll
' Microsoft Scripting Runtime

Public FSO As New Scripting.FileSystemObject
Public FOL As Scripting.Folder
Public FIL As Scripting.File
Public FST As Scripting.TextStream
Public LOG As Scripting.TextStream
 
Sub test_envi()
envi
End Sub

Sub envi(Optional ByVal action As String) ' IDX00079

HEUTE = Date

Set DB = CurrentDb

On Error Resume Next

Do
    Err.Clear
    VRSI = DB.Properties("Versionsnummer")
    If Err = 3270 Then
        DB.Properties.Append DB.CreateProperty("Versionsnummer", dbText, "0000")
        DB.Properties.Refresh
    End If
Loop Until Err = 0

On Error GoTo 0


End Sub