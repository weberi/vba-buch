' ---------------------------------------------------------------------------
' läuft in Outlook
' 
' Benötigte Verweise:
'    Microsoft Scripting Runtime
'    Microsoft Forms
'    Microsoft Office
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 9.6.3.1  Code der UserForm AnhaengeDynForm v3 - Deklarationen
' ---------------------------------------------------------------------------
Public verzeichnispfad As String
Public colDateien As New Collection


' ---------------------------------------------------------------------------
' 9.6.3.1  Code der UserForm AnhaengeDynForm v3 - UserForm_Initialize
' ---------------------------------------------------------------------------
Private Sub UserForm_Initialize()
Dim FSO  As New Scripting.FileSystemObject
Dim verzeichnis As Scripting.Folder
Dim dat As Scripting.File
Dim chkbox As MSForms.CheckBox
Dim i As Integer
Dim gap As Integer: gap = 18
Dim top_offset As Integer: top_offset = 20

CheckboxesEntfernen
verzeichnispfad = PfadLesen

On Error GoTo PfadNichtGefunden
' falls der konfigurierte Pfad nicht (mehr) existiert
Set verzeichnis = FSO.GetFolder(Me.verzeichnispfad)
On Error GoTo 0

Me.lblPfad.Caption = verzeichnispfad
i = 1
For Each dat In verzeichnis.Files
  Set chkbox = _
    Me.Controls.Add("Forms.Checkbox.1", "cbdat" & i, True)
  With chkbox
    .Top = i * gap + top_offset
    .Left = 20
    .Caption = dat.Name
    .Width = 190
    .Value = False
  End With
  i = i + 1
Next dat
btnOK.Top = (i + 1) * gap
btnKonfig.Top = (i + 1) * gap
Me.Height = btnOK.Top + btnOK.Height + gap + top_offset
Exit Sub

PfadNichtGefunden:
Me.lblPfad.Caption = verzeichnispfad _
  & " - existiert nicht. Bitte konfigurieren!"
End Sub

' ---------------------------------------------------------------------------
' 9.6.3.1  Code der UserForm AnhaengeDynForm v3 - CheckboxesEntfernen
' ---------------------------------------------------------------------------
Private Sub CheckboxesEntfernen()
Dim ctrl As MSForms.Control
For Each ctrl In Me.Controls
  If Left(ctrl.Name, 5) = "cbdat" Then
    Me.Controls.Remove ctrl.Name
  End If
Next ctrl
End Sub


' ---------------------------------------------------------------------------
' 9.6.3.1  Code der UserForm AnhaengeDynForm v3 - PfadLesen
' ---------------------------------------------------------------------------
Private Function PfadLesen() As String
Dim speicher As StorageItem
Dim verzeichnispfad As String
Set speicher = Session.GetDefaultFolder(olFolderDrafts) _
  .GetStorage(AKONFIG, olIdentifyBySubject)

If speicher.Size = 0 Then
  AnhangKonfigurierenPicker
  ' jetzt ist ein Pfad im Speicher ...
  Set speicher = Session.GetDefaultFolder(olFolderDrafts) _
    .GetStorage(AKONFIG, olIdentifyBySubject)
End If

verzeichnispfad = speicher.UserProperties(VPFAD).Value
If Right(verzeichnispfad, 1) <> "\" Then
  verzeichnispfad = verzeichnispfad & "\"
End If

PfadLesen = verzeichnispfad
End Function

' ---------------------------------------------------------------------------
' 9.6.3.2  Code des „Ok“ -Button mit Verwendung einer Collection
' ---------------------------------------------------------------------------
Private Sub btnOK_Click()
Dim ctrl As MSForms.Control
Dim cbox As MSForms.CheckBox

Set colDateien = Nothing          ' Collection leeren, siehe [2]
For Each ctrl In Me.Controls
  If Left(ctrl.Name, 5) = "cbdat" Then
    Set cbox = ctrl
    If cbox.Value Then
     colDateien.Add ctrl.Caption
    End If
  End If
Next ctrl
Me.Hide
End Sub 

' ---------------------------------------------------------------------------
' 9.6.3.3  Code des Konfig-Button
' ---------------------------------------------------------------------------

Private Sub btnKonfig_Click()
    Me.lblPfad = "Konfiguration startet gleich ..."
    AnhangKonfigurierenPicker   ' statt AnhangKonfigurieren
    UserForm_Initialize
End Sub