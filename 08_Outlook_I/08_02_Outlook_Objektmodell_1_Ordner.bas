' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'    Outlook
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 8.2.3 Zugriff auf Outlook-Ordner
' ---------------------------------------------------------------------------
Sub AufOrdnerZugreifen()
Dim olApp As New Outlook.Application
Dim nsp As Outlook.Namespace
Dim ordner As Outlook.Folder
Dim besitzer As Outlook.Recipient
Set nsp = olApp.GetNamespace("MAPI")
Set ordner = nsp.GetDefaultFolder(olFolderInbox)            ' (1)
Debug.Print ordner.Items.Count
Set besitzer = nsp.CreateRecipient("Anna Beta")
Set ordner = nsp.GetSharedDefaultFolder(besitzer, olFolderCalendar) _
    .Folders("Termine")
Debug.Print ordner.Name & ordner.Items.Count                ' (2)
Set ordner = _
     nsp.GetDefaultFolder(olPublicFoldersAllPublicFolders) _
    .Folders("Meine Organisation").Folders("Termine")         '(3)
Debug.Print ordner.Name & ordner.Items.Count
End Sub

' ---------------------------------------------------------------------------
' 8.2.3 Navigation durch die Ordnerstruktur mit
' Parent und Folders
' ---------------------------------------------------------------------------
Sub OrdnerNavigieren()
Dim olApp As New Outlook.Application
Dim posteingang As Outlook.Folder
Dim ordner As Outlook.Folder
Set posteingang = olApp.Session.GetDefaultFolder(olFolderInbox)
With posteingang.Parent
  Debug.Print "Überordner: " & .Name
  For Each ordner In .Folders
    Debug.Print .Name & ": " & ordner.Name
  Next ordner
End With
For Each ordner In posteingang.Folders
  Debug.Print posteingang.Name & ": " & ordner.Name
Next ordner
End Sub


' ---------------------------------------------------------------------------
' 8.2.5  Direkter Zugriff auf Ordner mit der EntryID  - EntryIDZeigen
' ---------------------------------------------------------------------------
Sub EntryIDZeigen()
Dim olApp As Object
Dim ordner As Outlook.Folder
Dim namensraum As Outlook.Namespace
Set olApp = CreateObject("Outlook.Application")
Set namensraum = olApp.GetNamespace("MAPI")
Set ordner = namensraum.PickFolder  ' Ordner im Outlook-Dialog wählen
Debug.Print ordner.EntryID
olApp.Quit
Set olApp = Nothing
End Sub

' ---------------------------------------------------------------------------
' 8.2.5  Direkter Zugriff auf Ordner mit der EntryID  - PerEntryIDZugreifen
' ---------------------------------------------------------------------------
Sub PerEntryIDZugreifen()
Dim olApp As Outlook.Application
Dim ordner As Outlook.Folder
Dim namensraum As Outlook.Namespace
Set olApp = CreateObject("Outlook.Application")
Set namensraum = olApp.GetNamespace("MAPI")
Set ordner = olApp.GetNamespace("MAPI").GetDefaultFolder( _
    olPublicFoldersAllPublicFolders)      ' geeignet initialisieren
Set ordner = olApp.GetNamespace("MAPI").GetFolderFromID("00…000")
Debug.Print ordner.Name
Set olApp = Nothing
End Sub 