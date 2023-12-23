' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'    Outlook
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 8.6 Sub FindAppts
' ---------------------------------------------------------------------------
Sub FindAppts()
Dim olApp As New Outlook.Application

Dim kalender As Outlook.Folder
Dim treffenTabelle As Outlook.Table
Dim besprechungen As Outlook.Items
Dim treffen As Outlook.Row

Dim filter1 As String, filter2 As String, filterKomp As String
Dim zeile As Integer: zeile = 1

Set kalender = olApp.Session.GetDefaultFolder(olFolderCalendar)

filter1 = _
   """urn:schemas:calendar:dtstart"" >= '2023/01/01'"
filter2 = _
  """urn:schemas:calendar:dtstart"" <= '2023/12/31'"
filterKomp = "@SQL= " & filter1 & " AND " & filter2
Set treffenTabelle = kalender.GetTable(filterKomp)

With treffenTabelle.Columns
  .RemoveAll
  .Add ("Subject")
  .Add ("Start")
  .Add ("Duration")
End With

Dim wks As Worksheet
Set wks = Worksheets.Add

wks.Cells(zeile, 1) = "Subject"
wks.Cells(zeile, 2) = "Start"
wks.Cells(zeile, 3) = "Duration"

Do Until (treffenTabelle.EndOfTable)
  zeile = zeile + 1
  Set treffen = treffenTabelle.GetNextRow()
  wks.Cells(zeile, 1) = treffen("Subject")
  wks.Cells(zeile, 2) = treffen("Start")
  wks.Cells(zeile, 3) = treffen("Duration")
Loop
End Sub