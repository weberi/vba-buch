' ---------------------------------------------------------------------------
' läuft in MS Excel

' Benötigte Verweise:
' - Microsoft XML, v6.0
' - Microsoft Scripting Runtime
' - Microsoft Script Control 1.0
'- Modul:
' - VBA-JSON
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 10.2.3  Mit VBA HTTP GET Request senden und Response empfangen
' ---------------------------------------------------------------------------
Const ZIP_URL As String = "https://api.zippopotam.us/DE/"
Sub zipcode()
' Einfacher GET Request ohne Authentifizierung
Dim zip As String
Dim req As String
Dim XMLHttp As New MSXML2.ServerXMLHTTP60
zip = InputBox("Bitte eine PLZ in Deutschland eingeben:")
req = ZIP_URL & zip
' Request absenden
XMLHttp.Open "GET", req, False
XMLHttp.send
' Rückgabewerte ausgeben
MsgBox "Status: " & XMLHttp.Status
MsgBox "responseText: " & XMLHttp.responseText
If (XMLHttp.Status = 200) Then
  JsonAnalysieren (XMLHttp.responseText)
End If
End Sub


' ---------------------------------------------------------------------------
' 10.2.4  JSON auswerten mit VBA-JSON
' ---------------------------------------------------------------------------
Sub JsonAnalysieren(responseText As String)
Dim responseObjekt As Object
Set responseObjekt = JsonConverter.ParseJson(responseText)
Cells(2, 2).Value = responseObjekt("post code")
Cells(4, 2).Value = responseObjekt("places")(1)("place name")
Cells(5, 2).Value = responseObjekt("places")(1)("state")
Cells(6, 2).Value = responseObjekt("country")
End Sub

' ---------------------------------------------------------------------------
' 10.2.5  JSON auswerten mit MS Script Control
' ---------------------------------------------------------------------------
Sub JsonAnalysieren2(responseText As String)
Dim scrContr As New MSScriptControl.ScriptControl
scrContr.Language = "JScript"

Dim responseObjekt As Object
Dim place1Objekt As Object
Set responseObjekt = scrContr.Eval("(" & responseText & ")")
Set place1Objekt = CallByName(responseObjekt.places, 0, VbGet)

Cells(2, 2).Value = CallByName(responseObjekt, "post code", VbGet)
Cells(4, 2).Value = CallByName(place1Objekt, "place name", VbGet)
Cells(5, 2).Value = CallByName(place1Objekt, "state", VbGet)
Cells(6, 2).Value = CallByName(responseObjekt, "country", VbGet)
End Sub