' ---------------------------------------------------------------------------
' läuft in Word

' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 6.3  Textmarken - Sub TMAktualisieren
' ---------------------------------------------------------------------------
Sub TMAktualisieren(dok As Document, tmName As String, neu As String)
  Dim TMMerk As Range
  Set TMMerk = dok.Bookmarks(tmName).Range
  TMMerk.text = neu
  dok.Bookmarks.Add tmName, TMMerk
End Sub

' ---------------------------------------------------------------------------
' 6.3  Textmarken - Sub TextmarkeTest
' ---------------------------------------------------------------------------
Sub TextmarkeTest()
Documents.Add().Activate
With ActiveDocument
  .Range.InsertAfter ("Einfach einige Wörter schreiben")
  .Bookmarks.Add "test", .Words(2)      
  .Bookmarks("test").Range.Bold = True
  TMAktualisieren ActiveDocument, "test", "neue "
  .Bookmarks("test").Range.Italic = True
End With
End Sub