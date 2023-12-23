' ---------------------------------------------------------------------------
' läuft im Windows Script Host
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 11.3.2  initexcel.vbs	
' ---------------------------------------------------------------------------
On Error Resume Next
Set exApp = GetObject( ,"Excel.Application")
If Err.Number <> 0 Then
  Set exApp = CreateObject("Excel.Application")
End If
On Error GoTo 0
exApp.Visible = True
If exApp.Workbooks.Count < 1 Then
  exApp.Workbooks.Add
End If
exApp.Workbooks(1).Worksheets(1).Cells(1, 1) = "Hello World"