' ---------------------------------------------------------------------------
' läuft in Outlook
' 
' Benötigte Verweise:
'    Microsoft Scripting Runtime
'    Microsoft Forms
'    Microsoft Office
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 9.6.5  Modul AnhaengeAuswaehlen v3
' ---------------------------------------------------------------------------
Sub AnhangAuswaehlen()
Dim auswahlForm As AnhaengeDynForm

Dim nachricht As mailItem

Dim dat As Variant
' auf die Nachricht zugreifen
If TypeName(ActiveWindow) = "Explorer" Then
  If Not ActiveExplorer.ActiveInlineResponse Is Nothing Then
    Set nachricht = ActiveExplorer.ActiveInlineResponse
    End If
ElseIf TypeName(ActiveWindow) = "Inspector" Then
  If TypeOf ActiveInspector.CurrentItem Is Outlook.mailItem Then
    Set nachricht = ActiveInspector.CurrentItem
  End If
End If

If nachricht Is Nothing Then     ' sollte nie vorkommen
  MsgBox ("Funktion Anhänge ist hier nicht möglich.")
  Exit Sub
End If

' Auswahlform anzeigen
Set auswahlForm = New AnhaengeDynForm
auswahlForm.Show

For Each dat In auswahlForm.colDateien
  nachricht.Attachments.Add auswahlForm.verzeichnispfad & dat
Next dat
End Sub