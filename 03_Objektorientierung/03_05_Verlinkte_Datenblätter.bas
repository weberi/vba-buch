' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 3.5.1 Deklarationen 
' ---------------------------------------------------------------------------
Public Const START_NAME As String = "Start"
Public Const LISTE_NAME As String = "sportgruppe"

' ---------------------------------------------------------------------------
' 3.5.1 Sub Main
' ---------------------------------------------------------------------------
Sub Main()
Application.ScreenUpdating = False
BlaetterEinrichten
StartEinrichten
Application.ScreenUpdating = True
ThisWorkbook.SaveAs LISTE_NAME, xlOpenXMLWorkbookMacroEnabled
MsgBox "Erfolgreich gespeichert: " & ThisWorkbook.FullName
End Sub

' ---------------------------------------------------------------------------
' 3.5.2Sub BlaetterEinrichten 
' ---------------------------------------------------------------------------
Sub BlaetterEinrichten()
Dim listeBlatt As Worksheet
Dim blatt As Worksheet
Dim zellenAdr As String
Dim z As Long
Set listeBlatt = Worksheets(LISTE_NAME)
zellenAdr = "'" & START_NAME & "'!A1"
z = 2
Do While listeBlatt.Cells(z, 1) <> ""
  Set blatt = Worksheets.Add(after:=Worksheets(Worksheets.Count))
  With blatt
    .Name = listeBlatt.Cells(z, 2)
                  ' Werte in Zellen schreiben
    .Cells(1, 1) = listeBlatt.Cells(z, 1) & " " _
      & listeBlatt.Cells(z, 2)
    .Cells(1, 3) = listeBlatt.Cells(z, 6)
    .Cells(2, 1) = "Geburtstag"
    .Cells(2, 2) = listeBlatt.Cells(z, 3)
    .Cells(3, 1) = "Größe"
    .Cells(3, 2) = listeBlatt.Cells(z, 4)
    .Cells(4, 1) = "Gewicht"
    .Cells(4, 2) = listeBlatt.Cells(z, 5)
                 ' Zellen formatieren
    .Cells(1, 1).Style = "Überschrift 3"
    .Cells(1, 3).Interior.Color = RGB(100, 250, 100)
    .Cells(1, 3).Font.Name = "Courier New"
    .Cells(1, 3).Font.Size = 16
    .Cells(1, 3).Font.Bold = True
    .Cells(1, 3).Borders(xlEdgeBottom).Color = RGB(200, 50, 50)
    .Cells(1, 3).Borders(xlEdgeBottom).Weight = xlMedium
    .Cells(2, 2).NumberFormat = "dd.mm.yyyy"
    .Cells(4, 2).NumberFormat = "#0.00"
    .Hyperlinks.Add Anchor:=.Cells(1, 5), _
      Address:="", subaddress:=zellenAdr, _
      ScreenTip:="Gehe zu Start", _
      TextToDisplay:=START_NAME
    .Columns.AutoFit
  End With   
  z = z + 1 
Loop
End Sub

' ---------------------------------------------------------------------------
3.5.3 Sub StartEinrichten
' ---------------------------------------------------------------------------
Sub StartEinrichten()
Dim startBlatt As Worksheet
Dim blatt As Worksheet
Dim zellenAdr As String
Dim zelle As Range
Dim btn As Button
Dim z As Integer
z = 2
Set startBlatt = Worksheets.Add(after:=Worksheets(1))
startBlatt.Name = START_NAME

For Each blatt In Worksheets
  With blatt
    If .Name <> START_NAME And .Name <> LISTE_NAME Then
      zellenAdr = "'" & .Name & "'!A1"
      blatt.Hyperlinks.Add Anchor:=startBlatt.Cells(z, 2), _
        Address:="", _
        subaddress:=zellenAdr, _
        ScreenTip:="Zu " & .Name, _
        TextToDisplay:=.Name
      z = z + 1
    End If
  End With
Next blatt
                          ' Arbeitsblatt formatieren
                          ' einheitliche Spaltenbreite festlegen
startBlatt.Columns(1).ColumnWidth = 6  
                          ' Arbeitsblattregister einfärben
With startBlatt.Tab
  .Color = vbCyan         ' oder vbMagenta, vbRed
  .TintAndShade = 0       ' oder -0.5
End With
                          ' Schaltfläche einfügen
Set zelle = startBlatt.Cells(2, 3)
With zelle
  Set btn = startBlatt.Buttons.Add(.Left + 10, .Top, 75, 30)
End With
With btn
  .OnAction = "Aufraeumen"
  .Caption = "Aufräumen "
End With
                          ' Startblatt aktivieren, Gitterlinien ausschalten

startBlatt.Activate
ActiveWindow.DisplayGridlines = False
End Sub

' ---------------------------------------------------------------------------
' 3.5.4Sub Aufraeumen
' ---------------------------------------------------------------------------
Sub Aufraeumen()
Dim blatt As Worksheet
Application.DisplayAlerts = False
For Each blatt In Worksheets
  If blatt.Name <> LISTE_NAME Then
    blatt.Delete
  End If
Next blatt
Application.DisplayAlerts = True
End Sub
