' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'  Word
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 4.2.1.3  Laufende Instanz von Word ansprechen und Late Binding
' ---------------------------------------------------------------------------
Sub GetOrStartWordLate()
Dim wordApp As Object   '<-  Late Binding
Dim wordDoc As Object   '<-

'versuche ein laufendes Word zu nutzen
On Error Resume Next
Set wordApp = GetObject(, "Word.Application")

If Err.Number <> 0 Then
  ' lösche den Fehler, erzeuge ein neues Word
  Err.Clear
  Set wordApp = CreateObject("Word.Application")
End If

wordApp.Visible = True

If wordApp.Documents.Count < 1 Then
  Set wordDoc = wordApp.Documents.Add
Else
  Set wordDoc = wordApp.Documents(1)
End If
WordDoc.Content.InsertAfter ("Hallo")
wordDoc.Save
End Sub