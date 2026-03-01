Option Compare Database
Option Explicit

Sub start() ' $00:00004
Dim URL As String
Dim DML As MSHTML.HTMLDocument
Dim tblAll As MSHTML.IHTMLElementCollection
Dim t As Long
Dim tbl As MSHTML.HTMLTable
Dim rw As Long
Dim cl As Long
Dim tmp As String
Dim HEA As String
Dim SQL As String

If BL = 0 Then PB.initialize

URL = SQ.sel("SELECT TXT FROM SETT WHERE LUP='HEUTE' {AND BL}")

Set DML = WB.GetHtml(URL)

' ### Paarungen

tmp = DML.documentElement.innerHTML

MN.bookPaarungen tmp

Stop

' ### Tabelle

Set tblAll = DML.getElementsByTagName("table")

For t = 0 To tblAll.length - 1          '0-basiert
    Set tbl = tblAll(t)
    
    tmp = tbl.innerText
    
    If tmp Like "Pxl. Team Sp. Diff. Pkt.*" Then
        
        For rw = 1 To tbl.rows.length - 1
        
            SQL = "DELETE * FROM TABL WHERE TEAM='{Team}' AND SP={Sp.};INSERT INTO TABL VALUES ({BL},'{Team}',{Sp.},{TOR},{Diff.},{Pkt.})"
        
            SQL = Replace(SQL, "{BL}", BL)
        
            For cl = 0 To tbl.rows(rw).cells.length - 1
                
                HEA = Trim(tbl.rows(0).cells(cl).innerText)
                tmp = Trim(tbl.rows(rw).cells(cl).innerText)
                
                Do While HEA = "Team"
                    
                    dum = SQ.sel("SELECT KEY FROM SETT WHERE LUP = 'MATCH' AND TXT = '" & tmp & "'")
                
                    If IsNull(dum) Then
                
                        SQ.BAT "INSERT INTO SETT (BL, LUP, TXT) VALUES (0, 'MATCH', '" & tmp & "')"
                        Stop
                    Else
                    
                        tmp = CStr(dum)
                        Exit Do
                    End If
                    
                Loop
                
                
                
                HEA = "{" & HEA & "}"
                
                SQL = Replace(SQL, HEA, tmp)

            Next cl
                        
            SQL = Replace(SQL, "{TOR}", "Null")
            
            SQ.BAT SQL
            
        Next rw
        
        Exit For
    
    End If
    
Next t

End Sub

Sub bookPaarungen(innerHTML As String)
Dim tmp As String
Dim t() As String
Dim u As Long

tmp = innerHTML

tmp = Replace(tmp, vbLf, "")
tmp = Replace(tmp, vbCr, "")
tmp = Replace(tmp, "SPAN", "DIV")
tmp = Replace(tmp, Chr(34), "")

tmp = Replace(tmp, "<H2 class=kick__site-headline>", vbCrLf)

For Each v In Split(tmp, vbCrLf)

    If Split(v, "</H2>")(0) = "Begegnungen" Then
        tmp = Split(v, "</H2>")(1)
        Exit For
    End If
Next v

tmp = Replace(tmp, "<DIV class=", vbCrLf)
 
innerHTML = Replace(tmp, "</DIV>", "")

CB.StringToClipboard innerHTML

For Each v In Split(innerHTML, vbCrLf)
    tmp = Trim(CStr(v))
    t = Split(tmp, ">")
    u = UBound(t)
    
    If u <> 1 Then GoTo NEXTV
    
    Stop



NEXTV:
Next v


End Sub