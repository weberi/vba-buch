' ---------------------------------------------------------------------------
' läuft in Word

' Benötigte Verweise:
'   Excel
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 6.2.2  Teil „Word2Excel“ der Demo-Anwendung - Sub Word2Excel
' ---------------------------------------------------------------------------
Sub Word2Excel ()
Dim exApp As New Excel.Application
Dim dok As Word.Document
Dim blatt As Worksheet
Dim mappe As Workbook
Dim sammel As String
Dim zeile As Long
Dim index As Long
Dim absatz As Word.Paragraph

exApp.Visible = True
Set mappe = exApp.Workbooks.Add
Set blatt = mappe.Worksheets(1)
Set dok = ThisDocument

blatt.Columns(2).ColumnWidth = 39
blatt.Columns(3).ColumnWidth = 50
blatt.Columns(3).WrapText = True
zeile = 1
sammel = ""

On Error GoTo Fehler
For index = 1 To dok.Paragraphs.Count - 1
  Set absatz = dok.Paragraphs(index)
  If absatz.Format.Style = "Überschrift 1" Then
    blatt.Cells(zeile, 2).Value = AbsEntfernen(absatz.Range.text)
    blatt.Cells(zeile, 2).Interior.Color = wdColorAqua
    blatt.Cells(zeile, 3).Interior.Color = wdColorAqua

     zeile = zeile + 1
  ElseIf absatz.Format.Style = "Überschrift 2" Then
    blatt.Cells(zeile, 2).Value = AbsEntfernen(absatz.Range.text)
  ElseIf absatz.Next.Format.Style = "Überschrift 1" _
      Or absatz.Next.Format.Style = "Überschrift 2" _
      Or index = dok.Paragraphs.Count - 1 _
      Then
    sammel = sammel & absatz.Range.text
    blatt.Cells(zeile, 3).Value = AbsEntfernen(sammel)
    sammel = ""
    zeile = zeile + 1
  Else
    sammel = sammel & absatz.Range.text
  End If
Next
exApp.DisplayAlerts = False

If (vbYes = MsgBox("Speichern?", vbYesNo)) Then
  mappe.SaveAs Left(dok.FullName, Len(dok.FullName) - 5) & ".xlsx"
End If

GoTo Aufraeumen

Fehler:
Debug.Print Err.Source, Err.Description

Aufraeumen:
mappe.Close SaveChanges:=False
exApp.Quit
Set exApp = Nothing
End Sub

' ---------------------------------------------------------------------------
' 6.2.2  Teil „Word2Excel“ der Demo-Anwendung - Function AbsEntfernen
' ---------------------------------------------------------------------------
Function AbsEntfernen(text As String) As String
If text = "" Then
  AbsEntfernen = text
Else
  If Right(text, 1) = vbLf Or Right(text, 1) = vbCr _
  Or Right(text, 1) = vbCrLf Then
    AbsEntfernen = AbsEntfernen(Left(text, Len(text) - 1))
  Else
    AbsEntfernen = text
  End If
End If
End Function