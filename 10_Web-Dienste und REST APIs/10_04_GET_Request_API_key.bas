' ---------------------------------------------------------------------------
' läuft in MS Excel

' Benötigte Verweise:
' - Microsoft XML, v6.0
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 10.4  Code-Beispiel Spritpreise-Umkreissuche mit API-Key
' ---------------------------------------------------------------------------
Const SPRIT_URL = _
  "https://creativecommons.tankerkoenig.de/json/list.php?"

Const APIKEY = ""           ' hier einsetzen

Sub umkreissuche()
Dim req As String
Dim param As Integer

Dim XMLHttp As New MSXML2.ServerXMLHTTP60

req = SPRIT_URL
With Worksheets("Spritpreise")
  For param = 2 To 6
10 Web-Dienste und REST APIs
    req = req & Cells(param, 1).Value & "=" _
      & Cells(param, 3).Value & "&"
  Next param
  req = req & "apikey=" & APIKEY
End With

XMLHttp.Open "GET", req, False
XMLHttp.send
If (XMLHttp.Status = 200) Then
  MsgBox "Status: " & XMLHttp.Status
  MsgBox "responseText: " & XMLHttp.responseText
End If
End Sub
