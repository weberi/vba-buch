' ---------------------------------------------------------------------------
' läuft in Word

' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 6.4  Produktblätter - Type produkt
' ---------------------------------------------------------------------------
Private Type Produkt
  nummer As String
  name As String
  kategorie As String
  hersteller As String
  hoehe As Double
  massstab As String
  beschreibung As String
End Type


' ---------------------------------------------------------------------------
' 6.4  Produktblätter - Sub DatenblattDemo
' ---------------------------------------------------------------------------
Sub DatenblattDemo()
Dim dok As Document
Dim p As Produkt
Dim nr As Integer
nr = InputBox("Welches Produkt? Bitte 1 oder 2 eingeben.")
p = Produktdaten(nr)

Set dok = Documents.Open(ThisDocument.Path & "\datenblatt.docx")
dok.Bookmarks("name").Range.text = p.name
dok.Bookmarks("beschreibung").Range.text = p.beschreibung
With dok.Tables (1)
  .Cell(1, 2).Range.text = p.nummer
  .Cell(2, 2).Range.text = p.kategorie
  .Cell(3, 2).Range.text = p.hersteller
  .Cell(4, 2).Range.text = p.hoehe & " / " & p.hoehe * 2.54
  .Cell(5, 2).Range.text = p.massstab
End With

If vbYes = MsgBox("Speichern?", vbYesNo) Then
  dok.SaveAs2 ThisDocument.Path & "\prodblatt" & nr & ".docx"
End If
dok.Close SaveChanges:=False
End Sub

' ---------------------------------------------------------------------------
' 6.4  Produktblätter - Function Produktdaten
' ---------------------------------------------------------------------------
Function Produktdaten(nr As Integer) As Produkt
Dim p As Produkt
If nr = 1 Then
  p.nummer = "S10_1678"
  p.name = "1969 Harley Davidson Ultimate Chopper"
  p.kategorie = "Motorräder"
  p.hersteller = "Min Lin Diecast"
  p.hoehe = 1.77
  p.massstab = "1:10"
  p.beschreibung = "Dieser Nachbau hat einen funktionierenden " _
  & "Ständer, Vorderradaufhängung, Schalthebel, Fußbremshebel, " _
  & "Antriebskette, Räder und Lenkung."
Else
  p.nummer = "S10_1678"
  p.name = "1996 Moto Guzzi 1100i"
  p.kategorie = "Motorräder"
  p.hersteller = "Highway 66 Mini Classics"
  p.hoehe = 0.68
  p.massstab = "1:10"
  p.beschreibung = "Offizielle Moto Guzzi Logos und Insignien " _
  & " sowie Satteltaschen an der Seite. Ein detailreicher Motor, " _
  & "eine funktionstüchtige Lenkung und Federung, drehbare Räder " _
  & "und ein funktionierender Ständer begeistern. Das Modell " _
  & "hat zwei Ledersitze, zwei Auspuffrohre und eine Lenkertasche."
End If
Produktdaten = p
End Function