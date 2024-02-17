' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'    Outlook
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 8.5.3 Sub AnhaengeBearbeiten
' ---------------------------------------------------------------------------
Sub AnhaengeBearbeiten()
Dim protPfad As String: protPfad = ThisWorkbook.Path
Dim olApp As New Outlook.Application
Dim eingang As Outlook.MAPIFolder
Dim protokolle As Items
Dim mailAttachments As Outlook.Attachments

Dim filter1 As String, filter2 As String
Dim filter3 As String, filter4 As String
Dim filterKomplett As String
Dim mailitm As Object
On Error GoTo Fehler

filter1 = _
    """http://schemas.microsoft.com/mapi/proptag/0x001a001e"" & _
        = 'ipm.Note'"
filter2 = _
    """urn:schemas:httpmail:subject"" LIKE  '%Protokoll%'"
filter3 = _
    """urn:schemas:httpmail:datereceived"" >= " & DreiMonateZurueck
filter4 = """urn:schemas:httpmail:hasattachment"" = 1 "
filterKomplett = "@SQL= " & filter1 & " AND " & filter2 _
    & " AND " & filter3 & " AND " & filter4

Debug.Print filterKomplett

Set eingang = olApp.GetNamespace("MAPI"). _
    GetDefaultFolder(olFolderInbox)
Set protokolle = eingang.Items.Restrict(filterKomplett)
For Each mailitm In protokolle
  mailitm.Attachments(1).SaveAsFile protPfad & "\" _
      & mailitm.Attachments(1).Filename
Next mailitm

Set mailitm = Nothing
Exit Sub

Fehler:
If Err.Number <> 0 Then
  Debug.Print Err.Number & ": " & Err.Description
End If
Resume Next
End Sub

' ---------------------------------------------------------------------------
' 8.5.3 Function DreiMonateZuruec
' ---------------------------------------------------------------------------
Function DreiMonateZurueck() As String
Dim d As Date
d = DateAdd("m", -3, Date)
DreiMonateZurueck = "'" & Format(d, "DD-MM-YYYY hh:nn") & "'"
End Function
