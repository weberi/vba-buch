' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 2.7.1  Sub definieren 
' ---------------------------------------------------------------------------
Sub Bestaetigen(name As String, pers As Integer)
Dim text As String

text = "Anmeldung von " & name & " mit " & pers & _
  " Personen bestätigt."
Debug.Print text
End Sub

' ---------------------------------------------------------------------------
'  2.7.1  Sub aufrufen
' ---------------------------------------------------------------------------
Sub ProzedurDemo()
Bestaetigen "Hans", 4
Bestaetigen "Franz", 2
Bestaetigen "Silke", 3
Bestaetigen "Carol", 5
End Sub