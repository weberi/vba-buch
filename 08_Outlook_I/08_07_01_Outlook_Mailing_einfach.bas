' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'    Outlook
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 8.7.1 Sub EinfacheMailVersenden
' ---------------------------------------------------------------------------
Sub EinfacheMailVersenden()
Dim outlookApp As New Outlook.Application
Dim mailItem As Outlook.mailItem
' Abbruch, wenn Outlook nicht läuft
On Error GoTo Outlook_Fehler                  
' Test, ob Outlook läuft:
Set outlookApp = GetObject(, "Outlook.Application") 
' Fehlerabfangen wieder ausschalten
On Error GoTo 0 
Set mailItem = outlookApp.CreateItem(0)
With mailItem
  .To = "jemand@sonstwo.de"
  .Subject = "Test"
  .Body = "Hallo, dies ist ein Test."
End With

mailItem.Display
If (vbYes = MsgBox("Mails versenden?", vbYesNo)) Then
  mailItem.Send
Else
  mailItem.Close (olDiscard)  ' Schließen ohne Speichern
End If
Exit Sub
Outlook_Fehler:
MsgBox ("Abbruch - Bitte Outlook starten und Skript neu ausführen.")
End Sub