' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 2.7.4  Optionale und benannte Argumente: OptArgDemo
' ---------------------------------------------------------------------------
Sub OptArgDemo()
Dim orig As String: orig = "foobarbaz"

Debug.Print Mid(orig, 4, 6)                         ' barbaz
Debug.Print Mid(orig, 4)                            ' barbaz
Debug.Print Mid(String:=orig, Start:=4, Length:=3)  ' bar
End Sub

' ---------------------------------------------------------------------------
' 2.7.4  Optionale und benannte Argumente: OptArgDemo2
' ---------------------------------------------------------------------------
Sub OptArgDemo2()
Dim orig As String
orig = "Hallo"

MsgBox Prompt:=orig, Title:="Demo 1"
MsgBox orig, Title:="Demo 2"
MsgBox orig, , "Demo 3"
End Sub

