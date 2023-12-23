' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   Microsoft Visual Basic for Applications Extensibility 5.3
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 4.2.3.1  Referenzierte Bibliotheken auflisten
' ---------------------------------------------------------------------------
Sub VerweiseAuflisten()
Dim z As Integer
With ThisWorkbook.VBProject.References
  For z = 1 To .Count
    Cells(z, 1) = .Item(z).GUID
    Cells(z, 2) = .Item(z).Description
    Cells(z, 3) = "'" & .Item(z).major & "." & .Item(z).minor
    Cells(z, 4) = .Item(z).FullPath
    Cells(z, 5) = .Item(z).Name
  Next
End With
End Sub


' ---------------------------------------------------------------------------
' 4.2.3.2. Bibliothek aus VBA referenzieren
' ---------------------------------------------------------------------------
Sub Referenzieren()
Dim verw As VBIDE.Reference
Dim verws As VBIDE.References
Dim IstAktiv As Boolean

IstAktiv = False
Set verws = ThisWorkbook.VBProject.References
For Each verw In verws
  If verw.Name = "Scripting" Then
    IstAktiv = True
    Exit For
  End If
Next verw

If Not IstAktiv Then
' verws.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0
  verws.AddFromFile ("C:\Windows\SysWOW64\scrrun.dll")
End If
End Sub