' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'  keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 4.1.4  Konstanten
' ---------------------------------------------------------------------------
Const dateiname As String = "daten.xlsx"
Const blattname As String = "Datum"


' ---------------------------------------------------------------------------
' 4.1.4  Sub Pruefen2
' ---------------------------------------------------------------------------
Sub Pruefen2()
On Error GoTo Abbruch
Dim wkbDaten As Workbook
Set wkbDaten = Workbooks.Open(ThisWorkbook.Path & "\" & dateiname)
MsgBox (dateiname & " enthält " & _
  wkbDaten.Worksheets(blattname).UsedRange.Rows.Count & " Einträge.")
Exit Sub
Abbruch:
                          ' während der Entwicklung
MsgBox (Err.Number & " " & Err.Description & " " & Err.Source) 
                          ' danach nur dieses:
Select Case Err.Number
  Case 1004
    MsgBox ("Fehler beim Zugriff auf " & dateiname & "." & _
      Chr(13) & "Bitte die Datei hier ablegen: " & ThisWorkbook.Path)
  Case 9
    MsgBox ("Fehler in Datei " & dateiname & "." & _
    Chr(13) & "Es gibt kein Arbeitsblatt " & blattname & ".")
End Select
End Sub