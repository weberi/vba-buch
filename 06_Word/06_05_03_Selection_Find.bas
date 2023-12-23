' ---------------------------------------------------------------------------
' läuft in Word

' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 6.5.3 Wörter innerhalb der Auswahl formatieren - Konstante 
' ---------------------------------------------------------------------------
Const SUCHTEXT As String = "VBA"


' ---------------------------------------------------------------------------
' 6.5.3  Wörter innerhalb der Auswahl formatieren - Sub Formatieren
' ---------------------------------------------------------------------------
Sub Formatieren()
Dim bereich As Range
Set bereich = Selection.Range.Duplicate   ' Find ändert Range
With bereich.Find
  .Text = SUCHTEXT
  .MatchCase = True         ' Groß-Kleinschreibung wie im Suchtext
  .MatchWholeWord = True    ' nur ganze Wörter
  .MatchWildcards = False
  .MatchSoundsLike = False
  .MatchAllWordForms = False
  Do While .Execute = True And bereich.Start <= Selection.Range.End
    If bereich.InRange(Selection.Range) And _
        ActiveDocument.Characters(bereich.End + 1).Text <> "-" Then
      bereich.Font.Color = wdColorAqua
      bereich.Font.Name = "Arial"
      bereich.Font.Bold = True
      bereich.Collapse wdCollapseEnd
    End If
  Loop
  End With
End Sub