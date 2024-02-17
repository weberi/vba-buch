' ---------------------------------------------------------------------------
' läuft in Outlook
' 
' Benötigte Verweise:
'    keine
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 9.3.2  AnhangAuswaehlen v1 - Deklarationen
' ---------------------------------------------------------------------------

Const verzeichnispfad As String = "C:\Users\...\...\Anhaenge\"
Const anhang1 As String = "Imagebroschüre.pdf"
Const anhang2 As String = "AGB.pdf"
Const anhang3 As String = "Preisliste.pdf"

Public istAnhang1 As Boolean
Public istAnhang2 As Boolean
Public istAnhang3 As Boolean

' ---------------------------------------------------------------------------
' 9.3.2  AnhangAuswaehlen v1 - Sub AnhangAuswaehlen
' ---------------------------------------------------------------------------
Sub AnhangAuswaehlen()
Dim auswahlForm As AnhaengeAuswahlForm
Dim nachricht As MailItem
' auf die Nachricht zugreifen
If TypeName(ActiveWindow) = "Explorer" Then
  If Not ActiveExplorer.ActiveInlineResponse Is Nothing Then
    Set nachricht = ActiveExplorer.ActiveInlineResponse
  End If
ElseIf TypeName(ActiveWindow) = "Inspector" Then
  If TypeOf ActiveInspector.CurrentItem Is Outlook.MailItem Then
    Set nachricht = ActiveInspector.CurrentItem
  End If 
End If  

If nachricht Is Nothing Then     ' sollte nie vorkommen
  MsgBox ("Funktion Anhänge ist hier nicht möglich.")
  Exit Sub
End If

' Auswahlform anzeigen
Set auswahlForm = New AnhaengeAuswahlForm
auswahlForm.Show
' Benutzer bedient die UserForm ...
'  ... jetzt sind in der UserForm die Variablen istAnhang_* gesetzt
If istAnhang1 Then
  nachricht.Attachments.Add verzeichnispfad & anhang1
End If
If istAnhang2 Then
  nachricht.Attachments.Add verzeichnispfad & anhang2
End If
If istAnhang3 Then
  nachricht.Attachments.Add verzeichnispfad & anhang3
End If
Set nachricht = Nothing
Set auswahlForm = Nothing
End Sub
