' ---------------------------------------------------------------------------
' läuft in PowerPoint
' 
' Benötigte Verweise:
' keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 7.1.2  Der Objekttyp Shape
' ---------------------------------------------------------------------------

Sub ShapesAuflisten()
Dim elem As Shape
For Each elem In ActivePresentation.Slides(1).Shapes
  Debug.Print
  Debug.Print "*** " & elem.Name
  Debug.Print "ist Platzhalter? " & (elem.Type = msoPlaceholder)
  Debug.Print "ist Bild? " & (elem.Type = msoPicture)
  Debug.Print "hat Text? " & elem.HasTextFrame
  Debug.Print "hat Tabelle? " & elem.HasTable
  Debug.Print "hat Diagramm? " & elem.HasChart
Next elem
End Sub

' ---------------------------------------------------------------------------
' 7.1.3 Shape-Objekte bearbeiten - TextelementAnsprechen
' ---------------------------------------------------------------------------
Sub TextelementAnsprechen()
Dim elem As Shape
Set elem = ActivePresentation.Slides(1). _
  Shapes("Inhaltsplatzhalter 2")
If elem.HasTextFrame Then
  If elem.TextFrame.HasText Then
    With elem.TextFrame.TextRange
      Debug.Print Left(.Text, 10) & "..."   ' Dies ist e...
      Debug.Print .Length                   '   97
      Debug.Print .Words.Count              '   19 
      .Words(19).Font.Size = 44
      .Words(19).Font.Name = "Times New Roman"
      .Words(19).Font.Italic = msoTrue
      .Words(19).Font.Color.RGB = RGB(200, 80, 80)
    End With
  End If
End If
End Sub

' ---------------------------------------------------------------------------
' 7.1.3 Shape-Objekte bearbeiten - TabelleAnsprechen
' ---------------------------------------------------------------------------
Sub TabelleAnsprechen()
Dim elem As Shape
Dim tabelle As Table
Set elem = ActivePresentation.Slides(1). _
  Shapes("Inhaltsplatzhalter 4")
If elem.HasTable Then
  Set tabelle = elem.Table
  tabelle.Cell(2, 2).Shape.TextFrame2.TextRange.text = "999"
  tabelle.Cell(2, 2).Shape.TextFrame2.TextRange.Font.Size = "32"
  tabelle.Cell(2, 2).Shape.Fill.ForeColor.RGB = RGB(200, 0, 200)
End If
End Sub

' ---------------------------------------------------------------------------
' 7.1.3 Shape-Objekte bearbeiten - TabelleAnsprechen
' ---------------------------------------------------------------------------
Sub GrafikAnsprechen()
Dim elem As Shape
Set elem = ActivePresentation.Slides(1).Shapes("Grafik 5")
elem.Height = 200
elem.Width = 200
If elem.Type = msoPicture Then
  elem.PictureFormat.ColorType = msoPictureGrayscale
End If
End Sub
