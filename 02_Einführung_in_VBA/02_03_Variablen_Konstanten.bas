' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 2.3.2 Wertzuweisungen
' ---------------------------------------------------------------------------

Sub VariablenDemo()
Dim a As Long ' Variablendeklarationen
Dim b As Long
Dim c As Long
Dim d As Long
Dim erg as String
a = 50     ' Wertzuweisung
b = 5
c = b      ' c erhaelt den Wert von b
c = c + 1  ' der Wert von c wird um 1 erhoeht
d = a + b  ' Addition mit 2 Variablen und Wertzuweisung
erg = "c = " & c & " und d = " & d ' baut einen String und speichert 
                                  ' ihn in eine Variable
Debug.Print erg                   ' c = 6 und d = 55
End Sub

' ---------------------------------------------------------------------------
' 2.3.4 Arrays - Sub ArrayDemo1
' ---------------------------------------------------------------------------
Sub ArrayDemo1()
Dim vektor(1 To 5) As Integer
vektor(1) = 100
vektor(2) = 150
vektor(3) = 170
vektor(4) = 200
vektor(5) = 201
Debug.Print vektor(2) + vektor(3)
vektor(0) = 3 ' Index-Fehler!!
End Sub

' ---------------------------------------------------------------------------
' 2.3.4 Arrays - Sub ArrayDemo2
' ---------------------------------------------------------------------------
Sub ArrayDemo2()
Dim matrix(0 To 1, 0 To 2) As Double
matrix(0, 0) = 13.05
matrix(0, 1) = 1.002
matrix(0, 2) = 33
matrix(1, 0) = 65.88
matrix(1, 1) = 43.6
matrix(1, 2) = 32.2

Debug.Print matrix(1, 1)
End Sub