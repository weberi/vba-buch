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
' 4.1.5  Sub Pruefen3
' ---------------------------------------------------------------------------
Sub Pruefen3()
Dim wkbDaten As Workbook

On Error Resume Next
Err.Clear

Set wkbDaten = Workbooks.Open(ThisWorkbook.Path & "\" & dateiname)

If Err.Number <> 0 Then
  MsgBox ("Fehler beim Zugriff auf " & dateiname & "." & _
     Chr(13) & "Bitte die Datei hier ablegen: " & ThisWorkbook.Path)
  Exit Sub
End If

MsgBox (dateiname & " enthält " & _
  wkbDaten.Worksheets(blattname).UsedRange.Rows.Count & " Einträge.")
If Err.Number <> 0 Then
  MsgBox ("Fehler in Datei " & dateiname & "." & _
    Chr(13) & "Es gibt kein Arbeitsblatt " & blattname & ".")
    Exit Sub
End If

On Error GoTo 0
' hier kann weiterer Code folgen ...
End Sub