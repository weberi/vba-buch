' ---------------------------------------------------------------------------
' läuft im Windows Script Host
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 11.3.2  initscript.vbs	
' ---------------------------------------------------------------------------
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
Set Application = SapGuiAuto.GetScriptingEngine
Set connection = Application.Children(0)
Set session = connection.Children(0)
Set SessionInfo = session.info
If Err.Number <> 0 Then
  MsgBox "Keine SAP Session gefunden. Ende.", , "SAP GUI Skript"
  WScript.Quit
End If
On Error GoTo 0
WScript.ConnectObject session, "on"
WScript.ConnectObject Application, "on"
With SessionInfo
  msgResult = MsgBox("System: " & .SystemName & _
  " Mandant: " & .Client & _
  " User: " & .User & " Tcode: " & .Transaction & vbLf & _
  "Starten?", vbYesNo, "SAP GUI Skript")
End With
If msgResult <> vbYes Then
  MsgBox "Gestoppt.", , "SAP GUI Skript"
  WScript.Quit
End If
MsgBox "Skript beginnt...", , "SAP GUI Skript"


