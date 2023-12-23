' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'  keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 4.1.5  Konstanten
' ---------------------------------------------------------------------------
Const dateiname As String = "daten.xlsx"
Const blattname As String = "Datum"


' ---------------------------------------------------------------------------
' 4.1.5  Sub InsFolgejahr
' ---------------------------------------------------------------------------
Sub InsFolgejahr()
Dim wkbDaten As Workbook
Dim datumspalte As Range
Dim zelle As Range

On Error GoTo Abbruch
Set wkbDaten = Workbooks.Open(ThisWorkbook.Path & "\" & dateiname)
Set datumspalte = wkbDaten.Worksheets(blattname).UsedRange.Columns(1)

On Error Resume Next
For Each zelle In datumspalte.Cells
  zelle.Value = zelle.Value + 365
Next zelle
Exit Sub

Abbruch:
Select Case Err.Number
  Case 1004
    MsgBox ("Fehler beim Zugriff auf " & dateiname & "." & _
      Chr(13) & "Bitte die Datei hier ablegen: " & ThisWorkbook.Path)
  Case 9
    MsgBox ("Fehler in Datei " & dateiname & "." & _
      Chr(13) & "Es gibt kein Arbeitsblatt " & blattname & ".")
End Select
End Sub