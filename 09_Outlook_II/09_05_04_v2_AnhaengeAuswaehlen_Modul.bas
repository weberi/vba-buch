' ---------------------------------------------------------------------------
' läuft in Outlook
' 
' Benötigte Verweise:
'    keine
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 9.5.4  AnhangAuswaehlen v2 - Deklarationen
' ---------------------------------------------------------------------------

' nicht mehr noetig in version 2:
' Const verzeichnispfad As String = "C:\Users\...\...\Anhaenge\"
' unverändert aus Version 1
Const anhang1 As String = "Imagebroschüre.pdf"
Const anhang2 As String = "AGB.pdf"
Const anhang3 As String = "Preisliste.pdf"

Public istAnhang1 As Boolean
Public istAnhang2 As Boolean
Public istAnhang3 As Boolean

' ---------------------------------------------------------------------------
' 9.5.4  AnhangAuswaehlen v2 - Sub AnhangAuswaehlen
' ---------------------------------------------------------------------------
Sub AnhangAuswaehlen ()
Dim auswahlForm As AnhaengeAuswahlForm
Dim nachricht As mailItem
' zwei neue Variablen in Version 2:
Dim speicher As StorageItem
Dim verzeichnispfad As String

' auf die Nachricht zugreifen, unverändert aus Version 1
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
  MsgBox ("Funktion AnhängeAuswählen ist hier nicht möglich.")
  Exit Sub
End If

' neu in Version 2:
Set speicher = Session.GetDefaultFolder(olFolderDrafts) _
  .GetStorage(AKONFIG, olIdentifyBySubject)
If speicher.Size = 0 Then
  AnhangKonfigurieren
  Set speicher = Session.GetDefaultFolder(olFolderDrafts) _
      .GetStorage(AKONFIG, olIdentifyBySubject)
End If
verzeichnispfad = speicher.UserProperties(VPFAD).Value
If Right(verzeichnispfad, 1) <> "\" Then
  verzeichnispfad = verzeichnispfad & "\"
End If

' Auswahlform anzeigen, unverändert aus Version 1
Set auswahlForm = New AnhaengeAuswahlForm
auswahlForm.Show

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