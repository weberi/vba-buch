' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 2.6.1 IfThenElse-Verzweigung
' ---------------------------------------------------------------------------
Sub IfThenElseDemo()
Dim a As Integer
Dim b As Integer
Dim abstand As Integer

a = 60
b = 90

If a < b Then
    abstand = b - a
Else
    abstand = a - b
End If
Debug.Print abstand
End Sub

' ---------------------------------------------------------------------------
' 2.6.2 IfThen-Verzweigung
' ---------------------------------------------------------------------------
Sub IfThenDemo()
Dim a As Integer
Dim b As Integer
Dim abstand As Integer

a = 60
b = 90
abstand = a - b

If abstand < 0 Then
    abstand = abstand * (-1)
End If
Debug.Print abstand
End Sub


' ---------------------------------------------------------------------------
'  2.6.3 IfThenElsifElse-Verzweigung
' ---------------------------------------------------------------------------
Sub IfThenElsifElseDemo()
Dim land As String
land = "Schweiz"   ' Schweden Italien

If land = "Schweiz" Then
  Debug.Print "CHF"
ElseIf land = "Schweden" Then
  Debug.Print "SEK"
ElseIf land = "Großbritannien" Then
  Debug.Print "GBP"
Else
   Debug.Print "EUR"
End If
End Sub 

' ---------------------------------------------------------------------------
'  2.6.4 SelectCase-Verzweigung
' ---------------------------------------------------------------------------
Sub SelectCaseDemo()
Dim temperatur As Single
temperatur = 3.1

Select Case temperatur
  Case Is < 4
    Debug.Print "eisig"
    Debug.Print "Handschuhe erforderlich"
  Case 4 To 8
    Debug.Print "kalt"
  Case 8 To 15
    Debug.Print "kühl"
  Case Else
    Debug.Print "warm"
End Select
End Sub