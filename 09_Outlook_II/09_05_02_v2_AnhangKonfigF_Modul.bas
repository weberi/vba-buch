' ---------------------------------------------------------------------------
' läuft in Outlook
' 
' Benötigte Verweise:
'    keine
' ---------------------------------------------------------------------------
' ---------------------------------------------------------------------------
' 9.5.2  Konstanten und Deklarationen - AnhangKonfigurieren v2
' ---------------------------------------------------------------------------

Public Const AKONFIG As String = "AnhangKonfig"
Public Const VPFAD As String = "AnhangKonfig"
Public pfad As String

' ---------------------------------------------------------------------------
' 9.5.2  Sub AnhangKonfigurieren - v2
' ---------------------------------------------------------------------------
Sub AnhangKonfigurieren()
Dim konfigForm As AnhangKonfigForm
Dim ordner As Folder
Dim konfig As StorageItem
Dim info As UserProperty

Set ordner = Application.Session.GetDefaultFolder(olFolderDrafts)
Set konfig = ordner.GetStorage(AKONFIG, olIdentifyBySubject)

If konfig.Size = 0 Then
  ' beim 1. Aufruf
  Set info = konfig.UserProperties.Add(VPFAD, olText)
Else
  Set info = konfig.UserProperties(VPFAD)
End If

'vorhandenen Wert lesen
pfad = info.Value

Set konfigForm = New AnhangKonfigForm
konfigForm.Show

' neuen Pfad speichern
info.Value = pfad
konfig.Save
End Sub

