' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   Microsoft Scripting Runtime
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 5.2.2  Auf eine Datei zugreifen und Informationen anzeigen
' ---------------------------------------------------------------------------
Sub DateiInfo()
Const DATEI_NAME As String = "Datei.docx"
Dim FSO As New FileSystemObject
Dim datei As File
Dim dateipfad As String

dateipfad = ThisWorkbook.Path & "\" & DATEI_NAME
Debug.Print dateipfad
If FSO.FileExists(dateipfad) Then
  Set datei = FSO.GetFile(dateipfad)
  With datei
    Debug.Print .Name
    Debug.Print .DateCreated
    Debug.Print .Type
    Debug.Print .Path
    Debug.Print FSO.GetBaseName(.Path)
    Debug.Print FSO.GetExtensionName(.Path)
 End With
Else
  Debug.Print dateipfad & " ist nicht da"
End If
End Sub