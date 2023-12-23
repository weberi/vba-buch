' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 2.4.2  Verkettungsoperatoren und Stringfunktionen - KonkatDemo
' ---------------------------------------------------------------------------
Sub KonkatDemo()
Dim konkat As String
Dim anfang As String
Dim ende As String
anfang = "Anfang "
ende = " und Ende"
konkat = anfang & ende
Debug.Print konkat
End Sub

' ---------------------------------------------------------------------------
' 2.4.2  Verkettungsoperatoren und Stringfunktionen - StringFunktionen
' ---------------------------------------------------------------------------
Sub StringFunktionen()
Dim text1 As String
text1 = "Otto Müller, Kempten"
Debug.Print text1
Debug.Print Len(text1)          ' Länge des Strings
Debug.Print InStr(text1, "te")  ' Startposition eines Teilstrings im
                                  ' String
Debug.Print Mid(text1, 5, 6)    ' Teilstring ab Position 5, 6 Zeichen
                                  ' lang
Debug.Print Left(text1, 4)      ' Linker Teilstring, 4 Zeichen lang
Debug.Print Right(text1, 8)     ' Rechter Teilstring, 8 Zeichen lang
Debug.Print StrReverse(text1)   ' String umgedreht
Debug.Print Replace(text1, "t", "!!")  ' Zeichen t durch !! ersetzt                  
End Sub


' ---------------------------------------------------------------------------
' 2.4.3 Vergleichsoperatoren
' ---------------------------------------------------------------------------
Sub VergleicheDemo()
Dim i1 As Integer: i1 = 10
Dim i2 As Integer: i2 = 2
Dim s1 As String: s1 = "10"
Dim s2 As String: s2 = "2"
Dim s3 As String: s3 = "eins"
Dim erg As Boolean

Debug.Print "zwei Zahlen vergleichen:"
erg = i1 < i2
Debug.Print erg            ' Falsch  
Debug.Print i1 <> i2       ' Wahr

Debug.Print "zwei Strings vergleichen:"
Debug.Print s1 < s2        ' Wahr
Debug.Print s2 < s3        ' Wahr

Debug.Print "String mit Zahl vergleichen:"
Debug.Print i1 = s1        ' Wahr
Debug.Print i1 < s3        ' Fehler: Typen unverträglich!
End Sub

' ---------------------------------------------------------------------------
' 2.4.3 Vergleichsoperatoren
' ---------------------------------------------------------------------------
Const MAX_TEILNEHMER As Integer = 5
Sub BooleDemo()
Dim anmeldungen As Integer
anmeldungen = 5  ' oder -1 oder 200 ...

If anmeldungen > MAX_TEILNEHMER _
    Or anmeldungen < 0 Then
  MsgBox "Wert unzulässig!"
End If
End Sub