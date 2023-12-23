' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'    Outlook
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 8.7.2 Konstanten
' ---------------------------------------------------------------------------
Const VORLAGE as String = "vorlage.docx"

' ---------------------------------------------------------------------------
' 8.7.2  Formatierte E-Mails versenden 
' ---------------------------------------------------------------------------
Sub MailFormatiert()
Dim outlookApp As New Outlook.Application
Dim wordApp As New Word.Application
Dim wordDoc As Word.Document
Dim mailDoc As Word.Document
Dim mailItm As Outlook.mailItem
Dim addressrow As Integer
Dim istBestaetigt As Boolean
On Error GoTo Aufraeumen
Set wordDoc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & VORLAGE)
wordDoc.Content.Copy
addressrow
=
2
With Worksheets("Adressaten")
  Do While .Cells(addressrow, 1) <> ""
    Set mailItm = outlookApp.CreateItem(0)
    mailItm.BodyFormat = olFormatRichText
    mailItm.To = .Cells(addressrow, 1).Value
    mailItm.subject = .Cells(addressrow, 2).Value
    Set mailDoc = mailItm.GetInspector.WordEditor
    mailDoc.Content.Paste    ' MailItem enthält jetzt die Vorlage 
    mailDoc.Bookmarks("teilnehmername").Range.Text = _
      .Cells(addressrow, 3)
    mailDoc.Bookmarks("kursname").Range.Text = .Cells(addressrow, 4)

    If Not istBestaetigt Then
      mailItm.Display
      If (vbYes = MsgBox("Mails versenden?", vbYesNo)) Then
        istBestaetigt = True
      Else
        mailItm.Close (olDiscard)  ' ohne Speichern schliessen
            Exit Do
      End If
    End If
    mailItm.Display
    mailItm.Send
    addressrow = addressrow + 1
  Loop
End With

Aufraeumen:
If Err.Number <> 0 Then
  MsgBox (Err.Number & Chr(13) & Err.Description)
End If
On Error Resume Next   ' Word auf alle Fälle wieder schließen
wordDoc.Close (False) 
wordApp.Quit
Set wordApp = Nothing
End Sub
