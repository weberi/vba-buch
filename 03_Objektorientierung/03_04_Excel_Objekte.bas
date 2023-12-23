' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 3.4.4  Befehle zum Umgang mit Excel-Objekten
' ---------------------------------------------------------------------------
Sub ArbeitsblattDemo()
Dim NeuesBlatt As Worksheet

Set NeuesBlatt = ThisWorkbook.Worksheets.Add
NeuesBlatt.Name = "Demo"
NeuesBlatt.Cells(1, 1).Value = "Hallo Welt"
NeuesBlatt.Cells(2, 1).Value = 7
NeuesBlatt.Cells(2, 2).Value = 6
NeuesBlatt.Cells(2, 3).Formula = "=A2*B2"
MsgBox NeuesBlatt.Cells(2, 3).Value
MsgBox NeuesBlatt.Name & " wird wieder gelöscht...", vbOKOnly, "Demo"
NeuesBlatt.Delete
End Sub

' ---------------------------------------------------------------------------
' 3.4.5  Kurzformen und Voreinstellungen im Excel-Objektmodell
' ---------------------------------------------------------------------------
Sub ArbeitsblattDemoKurz()
Worksheets.Add
ActiveSheet.Name = "Demo"
Cells(1, 1) = "Hallo Welt"
Cells(2, 1) = 7
Cells(2, 2) = 6
Cells(2, 3).Formula = "=A2*B2"
MsgBox Cells(2, 3)
MsgBox ActiveSheet.Name & " wird wieder gelöscht...", vbOKOnly, "Demo"
ActiveSheet.Delete
End Sub