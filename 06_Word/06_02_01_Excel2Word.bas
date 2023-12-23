' ---------------------------------------------------------------------------
' läuft in Excel

' Benötigte Verweise:
'   Word
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 6.2.1  Teil „Excel2Word“ der Demo-Anwendung
' ---------------------------------------------------------------------------
Sub Excel2Word()
Dim blatt As Worksheet
Dim zeile As Long

Dim wApp As New Word.Application
Dim dok As Word.Document
Set blatt = Worksheets("Modelle")
wApp.Visible = True
Set dok = wApp.Documents.Add()
On Error GoTo Fehler
zeile = 1
With dok.Paragraphs
  Do While blatt.Cells(zeile, 2) <> ""
    .Last.Range.InsertBefore blatt.Cells(zeile, 2).Value
    If Cells(zeile, 3).Value = "" Then
      .Last.Format.Style = "Überschrift 1"
      .Add
    Else
      .Last.Format.Style = "Überschrift 2"
      .Add
      .Last.Format.Style = "Standard"
            .Last.Range.InsertBefore blatt.Cells(zeile, 3).Value
      .Add
    End If
    zeile = zeile + 1
  Loop
End With

With ThisWorkbook
  If MsgBox("Speichern?", vbYesNo, "Excel -> Word") Then
    dok.SaveAs Left(.FullName, Len(.FullName) - 5) & ".docx"
  End If
End With
GoTo Aufraeumen

Fehler:
  Debug.Print Err.Source, Err.Description

Aufraeumen:
dok.Close SaveChanges:=False
wApp.Quit
Set wApp = Nothing
End Sub