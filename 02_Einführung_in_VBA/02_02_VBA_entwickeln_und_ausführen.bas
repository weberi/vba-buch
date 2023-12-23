' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 2.2.1 Die Entwicklungsumgebung
' ---------------------------------------------------------------------------

Sub HalloWelt()

MsgBox "Hallo Welt!"
Debug.Print "Hallo Welt!"

End Sub


' ---------------------------------------------------------------------------
' 2.2.2 VBA-Code schreiben
' ---------------------------------------------------------------------------
Sub SyntaxDemo()

Dim i As Integer: i = 1
Dim info As String

info = "ich bin ein ziemlich " _
  & "langer Infotext."

' Debug.Print info   ' auskommentiert!
Debug.Print info     ' nicht auskommentiert!
End Sub

