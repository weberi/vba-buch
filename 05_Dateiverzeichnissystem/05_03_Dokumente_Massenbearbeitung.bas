' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   Microsoft Scripting Runtime
'   Microsoft Word (v2)
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 5.3 Dokumente-Massenbearbeitung - Konstanten
' ---------------------------------------------------------------------------
Const ZIELPFAD As String = "Worddateien"
Const ARBEITSPFAD As String = "FSO_Beispiele"
' Const NEUER_STIL = "Linien (markant)"             ' v2
Dim FSO As New Scripting.FileSystemObject
' Dim WordApp As New Word.Application               ' v2

' ---------------------------------------------------------------------------
' 5.3 Dokumente-Massenbearbeitung - Sub DateienSammeln
' ---------------------------------------------------------------------------
Sub DateienSammeln()
Dim arbeitsverz As Folder
Dim zielverz As Folder
Dim dat As File
Dim pfad As String

Err.Clear
On Error GoTo Ende
pfad = ThisWorkbook.Path & "\"

If FSO.FolderExists(pfad & ZIELPFAD) Then
  MsgBox ("ZIELPFAD existiert schon. Abbruch.")
Else
  Set arbeitsverz = FSO.GetFolder(pfad & ARBEITSPFAD)
  Set zielverz = FSO.CreateFolder(pfad & ZIELPFAD)
  VerzeichnisBearbeiten zielverz.Path & "\", arbeitsverz
End If

Ende:
If Err.Number <> 0 Then
  Debug.Print Err.Description
End If
' WordApp.Quit SaveChanges:=WdSaveOptions.wdDoNotSaveChanges ' v2
' Set WordApp = Nothing                                      ' v2
End Sub

' ---------------------------------------------------------------------------
' 5.3 Dokumente-Massenbearbeitung - Sub VerzeichnisBearbeiten
' ---------------------------------------------------------------------------
Sub VerzeichnisBearbeiten(praefix As String, verzeichnis As Folder)
Dim dat As File
' Dim dok As Document                                   ' v2
Dim unterverz As Folder
Dim praefixneu As String
Dim datpfad As String

praefixneu = praefix & verzeichnis.Name & "_"
For Each dat In verzeichnis.Files
  datpfad = praefixneu & dat.Name
  Debug.Print datpfad
  If dat.Type = "Microsoft Word-Dokument" Then
    dat.Copy (datpfad)                                   ' v1
'    Set dok = WordApp.Documents.Open(dat.Path)          ' v2
'    dok.ApplyQuickStyleSet2 (NEUER_STIL)                ' v2
'    dok.SaveAs (datpfad)                                ' v2
'    dok.Close                                           ' v2
  End If
Next dat
For Each unterverz In verzeichnis.SubFolders
  VerzeichnisBearbeiten praefixneu, unterverz
Next unterverz
End Sub
