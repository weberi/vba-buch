' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 2.8.2  Sichtbarkeit von Variablen und Konstanten  - ein Modul
' ---------------------------------------------------------------------------
Sub ModulVarSetzen()
a1 = "Modulebene"
Debug.Print a1
End Sub

Sub ProzedurVarSetzen()
Dim a1 As String
a1 = "Prozedurebene"
Debug.Print a1
End Sub

Sub ModulVarLesen()
Debug.Print a1
End Sub

Sub SichtbarDemo()      ' Ausgabe:
Debug.Print a1c          ' beim 1. Ausführen "", danach "Modulebene"
ModulVarSetzen          ' "Modulebene"
ProzedurVarSetzen       ' "Prozedurebene"
ModulVarLesen           ' "Modulebene"
Debug.Print a1          ' "Modulebene"
End Sub


