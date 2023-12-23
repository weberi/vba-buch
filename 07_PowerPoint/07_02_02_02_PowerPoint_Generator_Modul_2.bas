' ---------------------------------------------------------------------------
' läuft in Excel

' Benötigte Verweise:
'   PowerPoint 
'   Microsoft Scripting Runtime
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' .7.2.2.2 Modul Main - Deklarationen
' ---------------------------------------------------------------------------
Public FSO As FileSystemObject
Private PPApp As PowerPoint.Application
Private musterpraes As PowerPoint.Presentation

' ---------------------------------------------------------------------------
' .7.2.2.2 Modul Main - Sub Main
' ---------------------------------------------------------------------------
Sub Main()
Dim einstell As EinstellungT

On Error GoTo Fehler
Set FSO = CreateObject("Scripting.FileSystemObject")
Set PPApp = CreateObject("Powerpoint.Application")
einstell = Einstellung(Eingabe())
Set musterpraes = PPApp.Presentations.Open(einstell.musterpfad)

PraesentationErzeugen einstell
MsgBox ("Präsentation erzeugt.")

GoTo Aufraeumen

Fehler:
Select Case Err.Number
Case 1004
  MsgBox (Err.Number & ": " & Err.Description & Chr(13) _
    & "Bitte Einstellungen prüfen.")
Case 1008
  MsgBox (Err.Number & ": " & Err.Description & Chr(13) _
    & "Das Ausgabeverzeichnis für " & einst.ausgabedat & _
    " lässt sich nicht einrichten. Bitte Einstellungen ändern.")
Case Is < -200000
  MsgBox (Err.Number & ": " & Err.Description & Chr(13) _
    & "Sind Verzeichnis und Musterpräsentation " _
    & einst.musterpfad & " vorhanden?" & Chr(13) _
    & "Die Musterpräsentation darf nicht offen sein.")
Case Else
  MsgBox (Err.Number & ": " & Err.Description & Chr(13) _
    & "Unvorhergesehener Fehler.")
End Select

Aufraeumen:

On Error Resume Next
musterpraes.Close
Set musterpraes = Nothing
If PPApp.Presentations.Count = 0 Then
  PPApp.Quit
  Set PPApp = Nothing
End If
Set FSO = Nothing
MsgBox ("Programm beendet.")
End Sub

' ---------------------------------------------------------------------------
' .7.2.2.2 Modul Main - Sub PraesentationErzeugen
' ---------------------------------------------------------------------------
Sub PraesentationErzeugen(einst As EinstellungT)
Dim folie As PowerPoint.Slide
Dim diagrammPlatzhalter As PowerPoint.Shape
Set folie = musterpraes.Slides(1)
With folie.Shapes("Titelplatzhalter")
  .TextFrame.TextRange.Text = einst.titelbereich
End With
With folie.Shapes("Textplatzhalter")
  .TextFrame.TextRange.Text = einst.textbereich
End With
Set diagrammPlatzhalter = folie.Shapes("Diagrammplatzhalter")
einst.diagramm.SetSourceData Source:=einst.daten
einst.diagramm.ChartArea.Copy
diagrammPlatzhalter.Select msoTrue
musterpraes.Windows(1).View.PasteSpecial (ppPasteShape)
musterpraes.SaveCopyAs (einst.ausgabedat)
musterpraes.ExportAsFixedFormat einst.ausgabedat _
       & ".pdf", ppFixedFormatTypePDF, ppFixedFormatIntentPrint
End Sub