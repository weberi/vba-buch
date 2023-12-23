' ---------------------------------------------------------------------------
' läuft in Word

' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 6.5.4  Wörter mit bestimmtem Absatzformat - Konstanten
' ---------------------------------------------------------------------------
Const ABSATZFORMAT = "Standard"
Const SUCHTEXT = "VBA"


' ---------------------------------------------------------------------------
' 6.5.4  Wörter mit bestimmtem Absatzformat - Sub VBAFormatieren
' ---------------------------------------------------------------------------
Sub VBAFormatieren()
Dim bereich As Range
Set bereich = ActiveDocument.Range
With bereich.Find
  .Text = SUCHTEXT
  .MatchCase = True
  .MatchWholeWord = True
  .MatchWildcards = False
  .MatchSoundsLike = False
  .MatchAllWordForms = False
  Do While .Execute = True
    If bereich.ParagraphFormat.Style = ABSATZFORMAT Then
      bereich.Font.Color = wdColorAqua
      bereich.Font.Name = "Arial"
      bereich.Font.Bold = True
      bereich.Collapse wdCollapseEnd
    End If
  Loop
  End With
End Sub