' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   Microsoft Scripting Runtime
'   Microsoft Word (v2)
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 5.4 Log-Datei - Konstanten
' ---------------------------------------------------------------------------
Private LogDat As TextStream
Private Const LOGVERZEICHNIS As String = "Logs\"

' ---------------------------------------------------------------------------
' 5.4 Log-Datei - Sub Main
' ---------------------------------------------------------------------------
Sub Main()
  SchreibeLog ("Dies ist ein Text")
  SchreibeLog ("Dies ist noch ein Text")
  CloseLog
End Sub

' ---------------------------------------------------------------------------
' 5.4 Log-Datei - Sub SchreibeLog
' ---------------------------------------------------------------------------
Sub SchreibeLog(info As String)
If LogDat Is Nothing Then
  InitLog
End If
LogDat.WriteLine (Now() & " - " & Environ("username") & " - " & info)
End Sub

' ---------------------------------------------------------------------------
' 5.4 Log-Datei - Sub InitLog
' ---------------------------------------------------------------------------
Private Sub InitLog()
Dim FSO As FileSystemObject
Dim logVerz As String
Dim logDatPfad As String

Set FSO = New FileSystemObject
logVerz = ThisWorkbook.Path & "\" & LOGVERZEICHNIS & "\"
logDatPfad = logVerz & "\" & Format(Date, "yy-mm-dd") & "_log.txt"

If Not FSO.FolderExists(logVerz) Then
  FSO.CreateFolder (logVerz)
End If
Set LogDat = FSO.OpenTextFile(logDatPfad, ForAppending, True)
End Sub

' ---------------------------------------------------------------------------
' 5.4 Log-Datei - Sub CloseLog
' ---------------------------------------------------------------------------
Public Sub CloseLog()
If Not LogDat Is Nothing Then
  LogDat.Close
  Set LogDat = Nothing
End If
End Sub
