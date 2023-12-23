' ---------------------------------------------------------------------------
' läuft in Outlook
' 
' Benötigte Verweise:
'    keine
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 9.3.1 Code-Modul der UserForm AnhaengeAuswahl v1
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


