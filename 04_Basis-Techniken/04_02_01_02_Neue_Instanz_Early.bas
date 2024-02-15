' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'  Word
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 4.2.1.2  Neue Instanz von Word starten und Early Binding
' ---------------------------------------------------------------------------
Sub WordNewEarly()
Dim wordApp As Word.Application
Dim wordDoc As Word.Document
Set wordApp = New Word.Application

On Error GoTo Beenden
' wordApp.Visible = True
' wordApp.Visible = False

Set wordDoc = wordApp.Documents.Add
wordDoc.Content.InsertBefore ("Hallo Welt")
wordDoc.SaveAs (ThisWorkbook.Path & "\" & "halloWelt")
wordDoc.Close

Beenden:
wordApp.Quit
Set wordApp = Nothing
End Sub
