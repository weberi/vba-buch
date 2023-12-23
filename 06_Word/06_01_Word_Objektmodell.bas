' ---------------------------------------------------------------------------
' läuft in Word
' 
' Benötigte Verweise:
' keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 6.1.1  Die Collections Documents
' ---------------------------------------------------------------------------
Const DATEINAME = "test"

Sub DokumentErzeugen()
Dim dok As Document
Set dok = Documents.Add
dok.Content.InsertAfter "VBA-Objekte in Word"
dok.SaveAs2 DATEINAME & ".docx"
dok.ExportAsFixedFormat outputfilename:=DATEINAME & ".pdf", _
  exportformat:=wdExportFormatPDF
MsgBox dok.FullName
dok.Close
End Sub

' ---------------------------------------------------------------------------
' 6.1.4  Text ersetzen, einfügen und löschen
' ---------------------------------------------------------------------------
Sub RangeDemo()
Dim dok As New Document
Dim bereich As Range
Set bereich = dok.Range(0, 0)
Debug.Print dok.Content.Start   ' 0
Debug.Print dok.Content.End     ' 1
With bereich
  .InsertAfter "Objekte"
  .InsertBefore "VBA-"
  .InsertAfter " in Word"
  .Words(4).Text = "mit "     
  .Collapse wdCollapseEnd
  .Delete wdWord, -2
  .SetRange 4, 11
  .Italic = True
  .Delete
End With
dok.Close
End Sub