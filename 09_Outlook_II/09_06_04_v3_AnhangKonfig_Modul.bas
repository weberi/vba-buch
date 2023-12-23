' ---------------------------------------------------------------------------
' läuft in Outlook
' 
' Benötigte Verweise:
'    Microsoft Scripting Runtime
'    Microsoft Forms
'    Microsoft Office
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 9.6.4 Modul AnhangKonfigPicker v3 - Deklarationen
' ---------------------------------------------------------------------------
Public Const AKONFIG As String = "AnhangKonfig"
Public Const VPFAD As String = "AnhangKonfig"

' ---------------------------------------------------------------------------
' 9.6.4 Modul AnhangKonfigPicker v3 - Sub AnhangKonfigurierenPicker
' ---------------------------------------------------------------------------
Sub AnhangKonfigurierenPicker()
Dim ordner As Folder
Dim konfig As StorageItem
Dim info As UserProperty
Dim pfad As String

Dim xlApp As Object
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False

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

Dim fd As Office.FileDialog
Set fd = xlApp.Application.FileDialog(msoFileDialogFolderPicker)

fd.AllowMultiSelect = False
If pfad <> vbNullString Then
  fd.InitialFileName = Left(pfad, InStrRev(pfad, "\"))
End If

Dim selectedItem As Variant
If fd.Show = -1 Then
  pfad = fd.SelectedItems(1)
End If

Set fd = Nothing
  xlApp.Quit
Set xlApp = Nothing

' neuen Pfad speichern
info.Value = pfad
konfig.Save
End Sub
