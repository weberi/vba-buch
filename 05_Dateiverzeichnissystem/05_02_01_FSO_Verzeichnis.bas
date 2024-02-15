' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   Microsoft Scripting Runtime
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 5.2.1  Auf Laufwerk, Verzeichnis und Datei zugreifen
' ---------------------------------------------------------------------------
Sub FSODemo()
Dim FSO As New FileSystemObject
Dim laufw As Drive
Dim dat As File
Dim verz As Folder

Debug.Print "Aktuelles Verzeichnis: " & FSO.GetFolder(".").Path
Debug.Print "Hinterlegter Pfad: " & ThisWorkbook.Path
Debug.Print

Set laufw = FSO.Drives("c")
Debug.Print "Wurzelvezeichnis: " & laufw.RootFolder.Path
Debug.Print "** Anzahl Verzeichnisse: " & _ 
  laufw.RootFolder.SubFolders.Count
For Each verz In laufw.RootFolder.SubFolders
  Debug.Print verz.Name
Next verz

Debug.Print vbLf & "** Anzahl Dateien: " & _
  laufw.RootFolder.Files.Count
For Each dat In laufw.RootFolder.Files
  Debug.Print dat.Name
Next dat
End Sub
