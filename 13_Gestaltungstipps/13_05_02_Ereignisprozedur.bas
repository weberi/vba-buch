' ---------------------------------------------------------------------------
' läuft in Word im Modul ThisDocument
' 
' Benötigte Verweise:
' keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 13.5.2 Ereignisbehandlungsroutinen entwickeln
' ---------------------------------------------------------------------------

Private Sub Document_Open()
  MsgBox "Tagomat - Automatische Verschlagwortung " _
    & "von technischen Dokumenten " _
    & Chr(13) & "Version 1.2 (2022)" _
    & Chr(13) & Chr(13) _
    & "Kontakt: Anna@hier.com", vbOKOnly, "Tagomat"
End Sub