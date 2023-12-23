' ---------------------------------------------------------------------------
' läuft nach Start im SAP GUI-Makrorecorder 
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 11.3.1  Makrorecorder-Befehlssequenz	
' ---------------------------------------------------------------------------
If Not IsObject(application) Then
  Set SapGuiAuto  = GetObject("SAPGUI")
  Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
  Set connection = application.Children(0)
End If
If Not IsObject(session) Then
  Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
  WScript.ConnectObject session,     "on"
  WScript.ConnectObject application, "on"
End If