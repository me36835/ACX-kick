Option Compare Database
Option Explicit

Function GetHtml(ByVal URL As String) As MSHTML.HTMLDocument
    Dim xhr As MSXML2.XMLHTTP60
    Dim html As MSHTML.HTMLDocument
    Set xhr = New MSXML2.XMLHTTP60
    xhr.Open "GET", URL, False
    xhr.send
    
    Set html = New MSHTML.HTMLDocument
    html.body.innerHTML = xhr.responseText      'ganzes Dokument rein
    Set GetHtml = html
End Function