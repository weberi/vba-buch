' ---------------------------------------------------------------------------
' läuft in Outlook
' 
' Benötigte Verweise:
'    keine
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 9.5.3 Code-Modul der UserForm AnhaengeAuswahl v2
' ---------------------------------------------------------------------------
Private Sub UserForm_Initialize()
Me.CheckBox1.Value = False
Me.CheckBox2.Value = False
Me.CheckBox3.Value = False
End Sub

Private Sub btnFertig_Click()
istAnhang1 = Me.CheckBox1.Value
istAnhang2 = Me.CheckBox2.Value
istAnhang3 = Me.CheckBox3.Value
Me.Hide
End Sub

Private Sub btnKonfig_Click()
  AnhangKonfigurieren
End Sub