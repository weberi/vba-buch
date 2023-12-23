' ---------------------------------------------------------------------------
' läuft in Outlook
' 
' Benötigte Verweise:
'    keine
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 9.5.1 Code-Modul der UserForm AnhangKonfigForm v2
' ---------------------------------------------------------------------------
Private Sub UserForm_Initialize()
Me.TextBox1.Value = pfad
End Sub

Private Sub btnFertig_Click()
pfad = Me.TextBox1.Value
Me.Hide
End Sub

Private Sub btnCancel_Click()
Me.Hide
End Sub


