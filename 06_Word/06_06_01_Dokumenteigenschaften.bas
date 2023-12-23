' ---------------------------------------------------------------------------
' läuft in Word

' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 6.6.1  Der Objekttyp DocumentProperty
' ---------------------------------------------------------------------------
Sub DemoDocProps()
Dim p As DocumentProperty
Dim props As DocumentProperties

Set props = ThisDocument.BuiltInDocumentProperties
' Set props = ThisWorkbook.BuiltInDocumentProperties   ' für Excel
On Error GoTo Info
For Each p In props
  Debug.Print p.Name & ":"
  Debug.Print " > " & p.Value
Next

Debug.Print "Author:  " & props("Author").Value
Debug.Print "Erste Property: " & props(1).Name
Debug.Print "Property Gibtsnicht " & props("Gibtsnicht").Value
Debug.Print "Number of notes: " & props("Number of notes ").Value
Exit Sub

Info:
Debug.Print " ! "    ' einen Fehler anzeigen
Resume Next          ' weiterarbeiten
End Sub