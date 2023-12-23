' ---------------------------------------------------------------------------
' läuft in MS Excel

' Benötigte Verweise:
' - Microsoft XML, v6.0
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 10.5  Code-Beispiel Stimmungserkennung mit POST Request
' ---------------------------------------------------------------------------
Const APIKEY As String = "XXXXxxxxXXXXxxxxXXXXxxxx"   ' hier ersetzen
Const SENT_URL As String = 
  "https://api.apilayer.com/sentiment/analysis"

Sub SentimentPost()
Dim XMLHttp As New MSXML2.ServerXMLHTTP60
Dim requestbody As String

requestbody = "Alles hat super geklappt. Das ist ein tolles " _
  & " Produkt. Danke! Ich freue mich sehr."

XMLHttp.Open "POST", SENT_URL, False

XMLHttp.setRequestHeader "apikey", APIKEY
XMLHttp.setRequestHeader "Content-Type", "text/plain"

XMLHttp.send requestbody
MsgBox "Status: " & XMLHttp.Status
MsgBox "responseText: " & XMLHttp.responseText
End Sub