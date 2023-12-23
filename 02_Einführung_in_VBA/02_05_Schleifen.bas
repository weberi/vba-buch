' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 2.5.1 ForNext-Schleife
' ---------------------------------------------------------------------------
Sub ForNextDemo()
Dim vektor(1 To 5) As Integer
Dim i As Integer

For i = 1 To 5
  vektor(i) = i * 10
  Debug.Print i & ": " & vektor(i)
Next i
Debug.Print i & " Fertig!!"
End Sub

' ---------------------------------------------------------------------------
' 2.5.2 DoLoop-Schleifen - DoWhileLoopDemo
' ---------------------------------------------------------------------------
Sub DoWhileLoopDemo()
Dim vektor(1 To 5) As Integer
Dim i As Integer

i = 1
Do While i <= 5
  vektor(i) = i * 10
  Debug.Print i & ": " & vektor(i)
  i = i + 1
Loop
Debug.Print i & " Fertig!!"
End Sub


' ---------------------------------------------------------------------------
'  2.5.2 DoLoop-Schleifen - DoUntilLoopDemo
' ---------------------------------------------------------------------------
Sub DoUntilLoopDemo()
Dim abbruch As Integer

abbruch = vbNo
Do Until abbruch = vbYes
  abbruch = MsgBox("Aufhören?", vbYesNo)
  Debug.Print abbruch
Loop
Debug.Print "Fertig!!"
End Sub

' ---------------------------------------------------------------------------
'  2.5.2 DoLoop-Schleifen - DoUntilLoopDemo2
' ---------------------------------------------------------------------------
Sub DoLoopUntilDemo2()
Dim abbruch As Integer
' abbruch = vbNo    hier nicht nötig
Do
  abbruch = MsgBox("Aufhören?", vbYesNo)
  Debug.Print abbruch
Loop Until abbruch = vbYes
Debug.Print "Fertig!!"
End Sub

' ---------------------------------------------------------------------------
'  2.5.3 ForEach-Schleifen
' ---------------------------------------------------------------------------
Sub ForEachDemo()
Dim gruppe(1 To 26) As String
Dim kind As Variant

gruppe(1) = "Anna"
gruppe(2) = "Benno"
gruppe(3) = "Chris"
gruppe(4) = "Dani"
gruppe(5) = "Emrah"
gruppe(6) = "Finja"
gruppe(26) = "Zora"

For Each kind In gruppe
  Debug.Print kind
Next kind
End Sub

' ---------------------------------------------------------------------------
'  2.5.4 Geschachtelte Schleifen
' ---------------------------------------------------------------------------
Sub ForNextInForNextDemo()
Const anzZeilen As Integer = 3
Const anzSpalten As Integer = 4
Dim matrix(1 To anzZeilen, 1 To anzSpalten) As String

Dim zeile As Integer
Dim spalte As Integer
Dim wert As Integer

For zeile = 1 To anzZeilen
    For spalte = 1 To anzSpalten
        matrix(zeile, spalte) = wert
        wert = wert + 1
    Next spalte
Next zeile

Debug.Print matrix(1, 1)
Debug.Print matrix(2, 2)
Debug.Print matrix(3, 3)
Debug.Print matrix(3, 4)
End Sub