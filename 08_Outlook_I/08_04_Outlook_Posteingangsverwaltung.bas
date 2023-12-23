' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'    Outlook
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 8.4 Konstante
' ---------------------------------------------------------------------------
CONST ABLEHN_BETREFF As STRING = "Einladung: P4"

' ---------------------------------------------------------------------------
' Sub PosteingangBearbeiten
' ---------------------------------------------------------------------------
Sub PosteingangBearbeiten()
Dim olApp As New Outlook.Application
Dim eingang As Outlook.MAPIFolder
Dim postheute As Items
Dim filterkriterien As String
Dim itm As Object
Dim meetreq As MeetingItem
Dim meeting As MeetingItem
Dim wsProt As Worksheet
Dim zeile As Integer: zeile = 1
Set wsProt = Worksheets.Add
On Error GoTo Fehler
filterkriterien = _
  "[ReceivedTime] > '" & Format(Date, "DD-MM-YYYY hh:nn") & "'"
  
Set eingang = olApp.GetNamespace("MAPI"). _ 
  GetDefaultFolder(olFolderInbox)
Set postheute = eingang.Items.Restrict(filterkriterien)
For Each itm In postheute
  wsProt.Cells(zeile, 1).Value = Format(itm.ReceivedTime, "hh:nn")
  wsProt.Cells(zeile, 2).Value = itm.Subject
  wsProt.Cells(zeile, 3).Value = itm.Class
  If itm.Class = olMeetingRequest Then
    Set meetreq = itm
    If meetreq.Subject = ABLEHN_BETREFF Then
      Set meeting = meetreq.GetAssociatedAppointment(True) _
        .Respond(olMeetingDeclined, True)
      meeting.Send
      meetreq.GetAssociatedAppointment(True).Delete
      wsProt.Cells(zeile, 4) = "Abgelehnt"
    End If
  End If
  zeile = zeile + 1
Next itm
Exit Sub
Fehler:
Debug.Print Err.Number & ": " & Err.Description
Resume Next
End Sub