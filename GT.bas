Option Compare Database
Option Explicit

Sub calculateIDX() 'IDX00084
Dim SQL As String
Dim RST As DAO.Recordset
Dim MIX As Long
Dim MSG As String

SQL = "SELECT MAX(IDX) FROM GIT WHERE IDX < 99999"
dum = SQ.sel(SQL)
MIX = CLng("0" & dum)

SQL = "SELECT * FROM GIT WHERE IDX = 99999 AND ( CPL = TRUE OR LIN > 0 )"
Set RST = DB.OpenRecordset(SQL, dbOpenDynaset)
Do While Not RST.EOF
            
    MIX = MIX + 1
    RST.Edit
    RST!IDX = MIX
    RST.Update
    
    MSG = "Es wurden Indexwerte ersetzt. Bitte noch einmal Nacharbeiten."
    
RST.MoveNext: Loop: RST.Close: Set RST = Nothing

If Len(MSG) Then MsgBox MSG

End Sub

Sub resetandtest()

GL.envi

Call GT.calculateIDX

dum = SQ.UPD("UPDATE GIT SET VRS = 0")

VR.AlleFormulareSpeichern
VR.AlleModuleSpeichern

GT.write2git
GT.temporary

' IDX00075 switch DEBUGS off
' dum = SQ.UPD("UPDATE SET SETT NUM = 0 WHERE GRP = 'DEBUGS' AND TXT LIKE 'GT.*'")

MsgBox "Fertig"

End Sub

Sub write2git() ' $00:00003

Dim obj As Object 'Deklariere obj als Object
Dim tmpCode As String
Dim i As Long
Dim v() As String
Dim such As String
Dim SQL As String
Dim tmp As String
Dim BAS As String
Dim MDL As String
Dim gitDT As Date
Dim tip As Byte
Dim ext As String: ext = "::.frm:.rpt::.bas"
Dim RST As DAO.Recordset
  
GL.envi

Call GT.calculateIDX

SQL = "SELECT TXT FROM SETT WHERE LUP = 'GIT Repository'"
REPO = CStr(SQ.sel(SQL))

If Not FSO.FolderExists(REPO) Then Exit Sub

dum = SQ.sel("SELECT COUNT(0) FROM GIT WHERE VRS = 0 AND CPL = TRUE")
If dum = 0 Then Exit Sub

SQL = "SELECT DAT FROM SETT WHERE LUP = 'GIT Repository'"
dum = SQ.sel(SQL)
If IsNull(dum) Then dum = #1/1/1950#
gitDT = CDate(dum)
If gitDT > Date Then gitDT = Date

SQL = "UPDATE SETT SET DAT = NOW(), NUM = " & VRSI & " WHERE LUP = 'GIT Repository'"
dum = SQ.UPD(SQL)

SQL = "UPDATE GIT SET VRS=0 WHERE Int(DAT) = Int(Now())"
dum = SQ.UPD(SQL)

tmp = REPO & "dummy.txt"

Set RST = CurrentDb.OpenRecordset("SYSINFO", dbOpenSnapshot)

Do While Not RST.EOF

    MDL = RST!Name
    
    Select Case RST!Type.Value
    Case -32761: tip = acModule
    Case -32764: tip = acReport
    Case -32768: tip = acForm
    Case Else
        tip = -1
        Stop
    End Select
    
    BAS = REPO & MDL & Split(ext, ":")(tip)
    Application.SaveAsText tip, MDL, tmp
        
    If tip <> acModule Then
        
        tmpCode = FSO.OpenTextFile(tmp, ForReading, False, -1).ReadAll()
        
        If InStr(tmpCode, "Option Compare Database") Then
            v = Split(tmpCode, "Option Compare Database")
            v(1) = "Option Compare Database" & v(1)
            If FSO.FileExists(tmp) Then FSO.DeleteFile tmp, True
            FSO.CreateTextFile(tmp, True, False).Write v(0)
            Call GT.savecompare(tmp, BAS, MDL)
            tip = acModule
            If FSO.FileExists(tmp) Then FSO.DeleteFile tmp, True
            FSO.CreateTextFile(tmp, True, False).Write v(1)
            BAS = REPO & RST!Name & Split(ext, ":")(tip)
        Else
            If FSO.FileExists(tmp) Then FSO.DeleteFile tmp, True
            FSO.CreateTextFile(tmp, True, False).Write tmpCode
        End If
    End If
    
    Call GT.savecompare(tmp, BAS, MDL)
    
RST.MoveNext: Loop: RST.Close: Set RST = Nothing

If FSO.FileExists(tmp) Then FSO.DeleteFile tmp, True

GT.ExportDatabaseSchema

' IDX00040
SQ.UPD ("UPDATE GIT SET VRS = " & VRSI & ", DAT = NOW() WHERE VRS = 0 AND CPL = TRUE")

Call GT.writeChangelogMD

Call GT.GitManager ' IDX00037 gitCommit

End Sub

Private Function Pad(ByVal text As Variant, ByVal laenge As Integer) As String ' IDX00047

If laenge > 0 Then
    Pad = Left(CStr(Nz(text, "")) & Space(laenge), laenge)
Else
    laenge = Abs(laenge)
    Pad = Right(Space(laenge) & CStr(Nz(text, "")), laenge)
End If

End Function

Public Sub GitManager(Optional ByVal strFilePath As String = "") ' IDX00048
' Aufbauend auf https://gemini.google.com/app/9c2a42cc2cb7654e
Dim shell As Object
Dim fileName As String
Dim SQL As String
Dim gitCmd As String
Dim exec As Object
Dim result As String

Set shell = CreateObject("WScript.Shell")

If Len(REPO) = 0 Then Stop
    
' 1. Wechsel in das Verzeichnis (immer notwendig)
gitCmd = "cmd.exe /c cd /d " & Chr(34) & REPO & Chr(34) & " && "

If strFilePath <> "" Then ' IDX00003 git add

    fileName = FSO.GetFileName(strFilePath)
    gitCmd = gitCmd & "git add " & Chr(34) & fileName & Chr(34)
    
Else ' IDX00004 git commit & push
    
    gitCmd = gitCmd & "git commit -m " & Chr(34) & "V" & VRSI & Chr(34) & " && git push"
    
End If

' IDX00033 shell feedback auffangen und auswerten
' IDX00067 Wenn man den Return auff?ngt, blitzt der Schirm kurz auf. - Ist so

Set exec = shell.exec(gitCmd)
result = exec.StdOut.ReadAll & exec.StdErr.ReadAll

If Len(result) Then
    MsgBox result
End If

Set exec = Nothing
Set shell = Nothing

End Sub

Sub savecompare(tmp As String, BAS As String, MDL As String, Optional SETT As String = "")
'  $00:00012

Dim strCode As String
Dim tmpCode As String
Dim such As String
Dim v() As String
Dim i As Long
Dim SQL As String
Dim RST As DAO.Recordset

tmpCode = FSO.OpenTextFile(tmp, ForReading).ReadAll()

If FSO.FileExists(BAS) Then
    strCode = FSO.OpenTextFile(BAS, ForReading).ReadAll()
    
    If strCode = tmpCode Then
        FSO.DeleteFile tmp, True
    Else
        FSO.DeleteFile BAS, True
    End If

End If
        
If FSO.FileExists(tmp) Then

    'If DBUG("GT.savecompare:NEW") Then Stop

    FSO.MoveFile tmp, BAS
    
    Call GT.GitManager(BAS) ' IDX00049 gitAdd

    If Len(SUP) Then
        
        SQL = "SELECT COUNT(*) FROM GIT WHERE MDL = '" & MDL & "' AND SUB ='" & SETT & "'"
        SQL = SQL & " AND CPL = TRUE ORDER BY IDX ASC"
        
        dum = "0" & SQ.sel(SQL)
        
        If Val(dum) = 0 Then
        
            i = DLookup("max(IDX)", "GIT", "IDX < 99999") + 1
            
            SQL = "INSERT INTO GIT (IDX, VRS, MDL, DAT, DSC, CPL, SUB"
            SQL = SQL & ") VALUES (" & i & ", 0, '" & MDL & "'"
            SQL = SQL & ", #1/1/1900#, 'Init', true, '" & SETT & "'"
            SQL = SQL & ")"
            
            dum = SQ.UPD(SQL)
        
        End If
        
        Exit Sub
        
    End If
        
' IDX00035 reset before count
    
    SQL = "UPDATE GIT SET LIN = -1 WHERE MDL = '" & MDL & "'"
    dum = SQ.UPD(SQL)
    
    SQL = "SELECT * FROM GIT WHERE MDL = '" & MDL & "'"
    SQL = SQL & " AND CPL = TRUE ORDER BY IDX ASC"
    Set RST = DB.OpenRecordset(SQL, dbOpenDynaset)
    
    If RST.EOF Then ' IDX00010 add new items to GIT table
    
        i = DLookup("max(IDX)", "GIT", "IDX < 99999") + 1
   
        SQL = "INSERT INTO GIT (IDX, VRS, MDL, DAT, DSC, CPL"
        SQL = SQL & ") VALUES (" & i & ", 280, '" & MDL & "', #1/1/1900#, 'added by GT.savecompare', true"
        SQL = SQL & ")"
        dum = SQ.UPD(SQL)
        
    Else ' IDX00002 update LIN Numbers on Tabel GIT


        Do While Not RST.EOF
        
            Stop
            
            such = IDX & Format(RST!IDX, "00000")
            
            v() = Split(tmpCode & such, such)
            
            If UBound(v) > 1 Then
        
                RST.Edit
                RST!LIN = UBound(Split(v(0), vbCrLf)) + 1
                RST.Update
                
            End If
            
        RST.MoveNext: Loop: RST.Close: Set RST = Nothing
        
    End If
    
End If

End Sub

Sub writeChangelogMD() ' IDX00036
Dim tmp As String
Dim SQL As String
Dim RST As DAO.Recordset

Dim wMdl As Integer: wMdl = 25
Dim wLin As Integer: wLin = -5
Dim wDsc As Integer: wDsc = 60

GL.envi

'If DBUG("GT.writeChangelogMD") Then Stop

If Len(REPO) = 0 Then
    SQL = "SELECT TXT FROM SETT WHERE LUP = 'GIT Repository'"
    REPO = CStr(SQ.sel(SQL))
End If

tmp = REPO & "Changelog.md"

If FSO.FileExists(tmp) Then FSO.DeleteFile tmp, True

Set FST = FSO.CreateTextFile(tmp, True, False)

' Header f?r das Markdown-File
FST.WriteLine "# Changelog"
FST.WriteLine "Generiert am: " & Now

' 2. SQL definieren (Du sagtest, das baust du selbst, hier ein Beispiel)
SQL = "SELECT * FROM GIT WHERE CPL = TRUE ORDER BY VRS DESC, MDL, LIN ASC"
Set RST = DB.OpenRecordset(SQL, dbOpenSnapshot)
 
Do While Not RST.EOF
    
    If dum <> RST!VRS Then
        dum = RST!VRS
        FST.WriteLine ""
        FST.WriteLine "## Version " & dum & " Date: " & Format(RST!DAT, "dd-Mmm-yyyy")
        FST.WriteLine ""
        FST.WriteLine "| Index | " & Pad("Modul.Prozedur", wMdl) & " | " & Pad("Zeile", Abs(wLin)) & " | " & Pad("Beschreibung", wDsc) & " |"
        FST.WriteLine "| ----- | " & String(wMdl, "-") & " | " & String(Abs(wLin), "-") & " | " & String(Abs(wDsc), "-") & " |"
    End If

    If Len(RST!MDL) * Len(RST!Sub) Then
        tmp = RST!MDL & "." & RST!Sub
    Else
        tmp = RST!MDL & RST!Sub
    End If
    
    If IsNull(RST!LIN) Then
        FST.WriteLine "| " & Pad(RST!IDX, -5) & " | " & Pad(tmp, wMdl) & " | " & Pad(" ", wLin) & " | " & Pad(RST!DSC, wDsc) & " |"
    Else
        If Val(RST!LIN) < 1 Then ' IDX00006
            FST.WriteLine "| " & Pad(RST!IDX, -5) & " | " & Pad(tmp, wMdl) & " | " & Pad(" ", wLin) & " | " & Pad(RST!DSC, wDsc) & " |"
        Else
            FST.WriteLine "| " & Pad(RST!IDX, -5) & " | " & Pad(tmp, wMdl) & " | " & Pad(RST!LIN, wLin) & " | " & Pad(RST!DSC, wDsc) & " |"
        End If
    End If
    
RST.MoveNext: Loop: RST.Close: Set RST = Nothing

FST.Close

' IDX00038 gitAdd
Call GT.GitManager(REPO & "Changelog.md")

End Sub

Sub temporary() ' IDX00066
Dim MDL As String
Dim tmp As String
Dim neu As String
Dim strCode As String
Dim tmpCode As String
Dim such As String
Dim v() As String
Dim i As Long
Dim SQL As String
Dim RST As DAO.Recordset

GL.envi

'If DBUG("GT.temporary") Then Stop

SQL = "SELECT TXT FROM SETT WHERE LUP = 'GIT Repository'"
REPO = CStr(SQ.sel(SQL))

If Not FSO.FolderExists(REPO) Then Exit Sub

SQL = "UPDATE GIT SET LIN = -1"
dum = SQ.UPD(SQL)

SQL = "SELECT * FROM GIT WHERE LIN = -1 AND CPL = TRUE AND LEN(MDL) > 0   ORDER BY MDL, IDX ASC"
Set RST = DB.OpenRecordset(SQL, dbOpenDynaset)

Do While Not RST.EOF

    neu = REPO & RST!MDL & ".bas"

    If FSO.FileExists(neu) Then
        
        If neu <> tmp Then
            'If DBUG("GT.temporary:Dateiwechsel", 0) Then Stop
            If FSO.FileExists(REPO & "dummy.txt") Then FSO.DeleteFile REPO & "dummy.txt", True
            FSO.CopyFile neu, REPO & "dummy.txt"
            tmp = neu
            tmpCode = FSO.OpenTextFile(neu, ForReading).ReadAll()
        End If
                
        such = "IDX" & Format(RST!IDX, "00000")
        v() = Split(tmpCode & such, such)
                
        If UBound(v) > 1 Then
        
            RST.Edit
            RST!LIN = UBound(Split(v(0), vbCrLf)) + 1
            RST.Update
        
        End If
        
    End If
            
RST.MoveNext: Loop: RST.Close: Set RST = Nothing

If FSO.FileExists(REPO & "dummy.txt") Then FSO.DeleteFile REPO & "dummy.txt", True

End Sub

Public Sub ExportDatabaseSchema() ' IDX00085
Dim tdf As DAO.TableDef
Dim qdf As DAO.QueryDef
Dim fld As DAO.Field
Dim tmp As String
Dim BAS As String
Dim SQL As String
Dim intFile As Integer
    
If Len(REPO) = 0 Then
    SQL = "SELECT TXT FROM SETT WHERE LUP = 'GIT Repository'"
    REPO = CStr(SQ.sel(SQL))
End If

tmp = REPO & "dummy.txt"
        
' Tabellen exportieren
For Each tdf In DB.TableDefs ' IDX00086
    ' Systemtabellen ?berspringen
    If Left(tdf.Name, 4) <> "MSys" And Left(tdf.Name, 1) <> "~" Then
     
        BAS = REPO & "Create Table " & tdf.Name & ".sql"
    
        SQL = BuildCreateTableScript(tdf)
        
        If FSO.FileExists(tmp) Then FSO.DeleteFile tmp, True
        
        FSO.CreateTextFile(tmp, True, False).Write SQL
        
        Call GT.savecompare(tmp, BAS, "TABLE", tdf.Name)
        
    End If
Next tdf

' Views (gespeicherte Abfragen) exportieren
For Each qdf In DB.QueryDefs ' IDX00087
    If Left(qdf.Name, 1) <> "~" Then
        BAS = REPO & "Create Query " & qdf.Name & ".sql"
         
        SQL = "CREATE QUERY [" & qdf.Name & "] AS"
        SQL = SQL & vbCrLf & qdf.SQL & vbCrLf
        
        If FSO.FileExists(tmp) Then FSO.DeleteFile tmp, True
        
        FSO.CreateTextFile(tmp, True, False).Write SQL
        
        Call GT.savecompare(tmp, BAS, "QUERY", qdf.Name)

    End If
Next qdf

Set tdf = Nothing
Set qdf = Nothing
Set DB = Nothing
End Sub

Private Function BuildCreateTableScript(tdf As DAO.TableDef) As String ' IDX00088
Dim fld As DAO.Field
Dim aox As DAO.Index
Dim SQL As String
Dim strFields As String
Dim strPK As String

SQL = "CREATE TABLE [" & tdf.Name & "] (" & vbCrLf

' Felder durchgehen
For Each fld In tdf.Fields
    strFields = strFields & "  [" & fld.Name & "] " & _
                GetAccessDataType(fld) & _
                IIf(fld.Required, " NOT NULL", "") & "," & vbCrLf
Next fld

' Prim?rschl?ssel finden
For Each aox In tdf.Indexes
    If aox.Primary Then
        strPK = "  CONSTRAINT [PK_" & tdf.Name & "] PRIMARY KEY ("
        Dim idxFld As DAO.Field
        For Each idxFld In aox.Fields
            strPK = strPK & "[" & idxFld.Name & "],"
        Next idxFld
        strPK = Left(strPK, Len(strPK) - 1) & ")" & vbCrLf
        Exit For
    End If
Next aox

SQL = SQL & strFields
If strPK <> "" Then SQL = SQL & strPK
SQL = Left(SQL, Len(SQL) - 3) & vbCrLf & ");"

BuildCreateTableScript = SQL
End Function

Private Function GetAccessDataType(fld As DAO.Field) As String 'IDX00089
Select Case fld.Type
    Case dbLong: GetAccessDataType = "LONG"
    Case dbText: GetAccessDataType = "TEXT(" & fld.Size & ")"
    Case dbMemo: GetAccessDataType = "MEMO"
    Case dbDate: GetAccessDataType = "DATETIME"
    Case dbBoolean: GetAccessDataType = "YESNO"
    Case dbDouble: GetAccessDataType = "DOUBLE"
    Case dbSingle: GetAccessDataType = "SINGLE"
    Case dbCurrency: GetAccessDataType = "CURRENCY"
    Case dbInteger: GetAccessDataType = "SHORT"
    Case dbByte: GetAccessDataType = "BYTE"
    Case Else: GetAccessDataType = "VARIANT"
End Select
End Function