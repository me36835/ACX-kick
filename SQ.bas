Option Compare Database
Option Explicit

Function UPD(SQL As String) As Long

Dim c As Long

If DB Is Nothing Then GL.envi

DB.Execute SQL, dbFailOnError

c = DB.RecordsAffected

'If c Then STX.makestats SQL, c
'UPD = c

End Function


Function sel(SQL As String) As Variant
Dim RST As DAO.Recordset

Set RST = SQ.RQS(SQL)

On Error Resume Next

    sel = RST.Fields(0).Value
    If Err Then sel = Null
    RST.Close: Set RST = Nothing
    
On Error GoTo 0
End Function

Function RQS(SQL As String) As DAO.Recordset

If DB Is Nothing Then PB.initialize

On Error Resume Next

RQS.Close
Set RQS = Nothing

On Error GoTo 0

SQL = Replace(SQL, "{AND BL}", "AND BL = " & BL)

On Error Resume Next

Set RQS = DB.OpenRecordset(SQL, dbOpenSnapshot)

If Err Then Set RQS = Nothing

On Error GoTo 0

End Function

Sub BAT(SQL As String)
Dim v As Variant
Dim tmp As String

If DB Is Nothing Then PB.initialize

For Each v In Split(SQL, ";")

    tmp = CStr(v)

    DB.Execute tmp, dbFailOnError

    dum = DB.RecordsAffected

Next v


End Sub