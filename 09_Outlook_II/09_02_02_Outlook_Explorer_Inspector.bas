x' ---------------------------------------------------------------------------
' läuft in Outlook
' 
' Benötigte Verweise:
'    keine
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 9.2.2  Die Objekte ActiveWindow, ActiveExplorer und
' ActiveInspector
' ---------------------------------------------------------------------------
Sub OffenesElementZugreifen()
Dim elem As Object
If TypeName(ActiveWindow) = "Explorer" Then
  If Not ActiveExplorer.Selection.Count = 0 Then
    Set elem = ActiveExplorer.Selection.Item(1)
  Else
    Debug.Print "Aktuell ist kein Element aktiv"
    Exit Sub
  End If
ElseIf TypeName(ActiveWindow) = "Inspector" Then
  Set elem = ActiveInspector.CurrentItem
End If

Debug.Print "Aktuelles Element: ' " & elem.Subject _
  & " (Typ: " & elem.Class & ")"
End Sub
