' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 2.7.3  Funktion definieren: Kreisflaeche
' ---------------------------------------------------------------------------
Function Kreisflaeche(radius As Double) As Double
Kreisflaeche = radius ^ 2 * 3.1416
End Function

' ---------------------------------------------------------------------------
' 2.7.3  Funktion definieren: Kreisumfang
' ---------------------------------------------------------------------------
Function Kreisumfang(radius As Double) As Double
Kreisumfang = 2 * radius * 3.1416
End Function

' ---------------------------------------------------------------------------
' 2.7.3  Funktion definieren: Rechteckflaeche
' ---------------------------------------------------------------------------
Function Rechteckflaeche(a As Double, b As Double) As Double
Rechteckflaeche = a * b
End Function

' ---------------------------------------------------------------------------
' 2.7.3  Funktion definieren und aufrufen: Zylinderoberflaeche
' ---------------------------------------------------------------------------
Function Zylinderoberflaeche(hoehe As Double, radius As Double) As
Double
Dim deckel As Double
Dim mantelflaeche As Double
deckel = Kreisflaeche(radius)
mantelflaeche = Rechteckflaeche(hoehe, Kreisumfang(radius))
Zylinderoberflaeche = 2 * deckel + mantelflaeche
End Function

' ---------------------------------------------------------------------------
' 2.7.3  Funktion aufrufen: FunktionenDemo
' ---------------------------------------------------------------------------
Sub FunktionenDemo()
Debug.Print Zylinderoberflaeche(3, 4)
Debug.Print Zylinderoberflaeche(3, 5.5)
End Sub