' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'    Outlook
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 8.3.1  Elemente in einem Outlook-Ordner auflisten
' ---------------------------------------------------------------------------
Sub ItemsAuflisten()
Dim outlApp As New Outlook.Application
Dim itm As Object
Dim inbox As Items
Dim meetreq As MeetingItem
Dim mailitm As mailItem
Dim max As Integer
max = 20  ' Stopper für volle Inboxes
Set inbox = _
  outlApp.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Items
For Each itm In inbox
  If itm.Class = olMail Then
    Set mailitm = itm
    Debug.Print "E-Mail von " & mailitm.Sender
  ElseIf itm.Class = olMeetingRequest Then
    Set meetreq = itm
    Debug.Print "Einladung von " & meetreq.SenderName
  End If
  max = max - 1
  If max <= 0 Then
    Exit For
  End If
Next itm
End Sub