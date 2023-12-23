' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'  keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 4.1.3  Konstante
' ---------------------------------------------------------------------------
Const dateiname As String = "daten.xlsx"


' ---------------------------------------------------------------------------
' 4.1.3  Sub Pruefen
' ---------------------------------------------------------------------------
Sub Pruefen()
On Error GoTo Abbruch
Dim wbDaten As Workbook
Set wbDaten = Workbooks.Open(ThisWorkbook.Path & "\" & dateiname)
MsgBox (dateiname & " enthält " & _  
  wbDaten.Worksheets(1).UsedRange.Rows.Count & " Einträge.")
Exit Sub
Abbruch:
MsgBox ("Fehler beim Zugriff auf " & dateiname & "." & _
  Chr(13) & "Bitte die Datei hier ablegen: " & ThisWorkbook.Path)
End Sub