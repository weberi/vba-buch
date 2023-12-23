' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
' keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
13.3.4 Umgebungsvariablen nutzen
' ---------------------------------------------------------------------------


Sub UmgebungsvariablenDemo()
Dim i As Integer
For i = 1 To 55
  Debug.Print i & ":  " & Environ$(i)
Next i
Debug.Print Environ$("USERNAME")
Debug.Print Environ$("HOMEPATH")
Debug.Print Environ$("APPDATA")
End Sub