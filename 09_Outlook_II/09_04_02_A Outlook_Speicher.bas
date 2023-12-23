' ---------------------------------------------------------------------------
' läuft in Outlook
' 
' Benötigte Verweise:
'    keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 9.4.2  StorageDemo
' ---------------------------------------------------------------------------
Sub StorageDemo()
Dim entwuerfe As Folder
Dim speicher As StorageItem
Dim info As UserProperty
Set entwuerfe = Application.Session.GetDefaultFolder(olFolderDrafts)
Set speicher = entwuerfe.GetStorage("Test", olIdentifyBySubject)
Debug.Print "Size: " & speicher.Size                   ' Size: 0
Set info = speicher.UserProperties.Add("tzahl", olNumber)
info.Value = 9009
speicher.UserProperties.Add("ttext", olText).Value = "Zum Testen"
speicher.Save
With speicher
  Debug.Print .Subject & ": Anzahl " & .UserProperties.Count
                                               ' Test: Anzahl 2
  Debug.Print .UserProperties(2).Name              ' tzahl
  Debug.Print .UserProperties("tzahl").Value       ' 9009
  Debug.Print .UserProperties("ttext").Value       ' Zum Testen
  .UserProperties("ttext").Value = "Geändert"
  Debug.Print .UserProperties("ttext").Value       ' Geändert
  ' Debug.Print .UserProperties("null").Value      ' Laufzeitfehler
End
With
speicher.Delete
End
Sub