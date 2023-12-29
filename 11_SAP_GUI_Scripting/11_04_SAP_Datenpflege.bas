' ---------------------------------------------------------------------------
' läuft in SOLIDWORKS 
' 
' Benötigte Verweise:
' - keine
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 11.4.1  SAP GUI-Skript aus der Wirtsanwendung Excel starten

' ---------------------------------------------------------------------------

Sub Starten()
Dim SapGuiAuto
Dim sapApp As SAPFEWSELib.GuiApplication
Dim connection As SAPFEWSELib.GuiConnection
Dim session As SAPFEWSELib.GuiSession
On Error GoTo START_FEHLER
Set SapGuiAuto = GetObject("SAPGUI")
Set sapApp = SapGuiAuto.GetScriptingEngine
Set connection = sapApp.Children(0)
Set session = connection.Children(0)
On Error GoTo 0

If (sapApp.Children.Count > 1) Or connection.Children.Count > 1 Then
  MsgBox "SAP-Anmeldung nicht eindeutig. " _
    & "Bitte nur einmal anmelden!", , "SAP GUI Skript"
   Exit Sub
End If
If (session.info.User = vbNullString) Then
  MsgBox "Kein SAP-User angemeldet. " _
    & "Bitte am System anmelden!", , "SAP GUI Skript"
   Exit Sub
End If
Bearbeiten session ' hier passiert etwas Nützliches
Exit Sub
START_FEHLER:
MsgBox "Keine SAP-Session gefunden. " _
  & "Bitte SAP starten und anmelden!", , "SAP GUI Skript"
End Sub

' ---------------------------------------------------------------------------
' 11.4.4  Code bearbeiten und zusammenstellen
' ---------------------------------------------------------------------------
Sub Bearbeiten(ss As SAPFEWSELib.GuiSession)
Dim zeile As Long
zeile = 2
Do While Cells(zeile, 1).Value <> ""
  MaterialBearbeiten ss, zeile
  zeile = zeile + 1
Loop
End Sub
Sub MaterialBearbeiten (ss As SAPFEWSELib.GuiSession, zeile As Long)
Dim mainw As SAPFEWSELib.GuiMainWindow
Dim statuszeile As SAPFEWSELib.GuiStatusbar
Set mainw = ss.FindById("wnd[0]")
mainw.FindById("tbar[0]/okcd").Text = "/nMM02"
mainw.sendVKey 0
mainw.FindById("usr/ctxtRMMG1-MATNR").Text = Cells(zeile, 1).Value
mainw.sendVKey 0
If IstMaterialFalsch(ss, zeile) Then
  Exit Sub
End If
ss.FindById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").GetAbsoluteRow(0).Selected
=
True
ss.FindById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").GetAbsoluteRow(3).Selected
=
True
ss.FindById("wnd[1]/tbar[0]/btn[0]").press
ss.FindById("wnd[1]/usr/ctxtRMMG1-WERKS").Text = Cells(zeile,
2).Value
ss.FindById("wnd[1]/usr/ctxtRMMG1-VKORG").Text = Cells(zeile,
3).Value

ss.FindById("wnd/usr/ctxtRMMG1-VTWEG").Text = Cells(zeile, 4).Value
ss.FindById("wnd[1]/tbar[0]/btn[0]").press
If IstOrgFalsch(ss, zeile) Then
  Exit Sub
End If
mainw.FindByName("MARA-VOLUM", "GuiTextField").Text = Cells(zeile, 
5).Value
mainw.FindByName("MARA-VOLEH", "GuiCTextField").Text = Cells(zeile,
6).Value
mainw.FindByName("MARA-GROES", "GuiTextField").Text = Cells(zeile,
7).Value
mainw.FindById("tbar[0]/btn[11]").press
Abschliessen ss, zeile
End Sub

' ---------------------------------------------------------------------------
' 11.4.5.1. Fehler „Falsches Material“
' ---------------------------------------------------------------------------
Function IstMaterialFalsch(ss As GuiSession, zeile As Long) As Boolean
Dim statuszeile As GuiStatusbar
Set statuszeile = ss.FindById("wnd[0]/sbar")
Cells(zeile, 8).Value = statuszeile.Messagetype
Cells(zeile, 9).Value = statuszeile.Text

If statuszeile.Messagetype = "E" Then
  ss.FindById("wnd[0]/tbar[0]/btn[15]").press
  IstMaterialFalsch = True
End If
End Function

' ---------------------------------------------------------------------------
'  11.4.5.2. Fehler „Falsche Organisationseinheit“
' ---------------------------------------------------------------------------
Function IstOrgFalsch(ss As GuiSession, zeile As Long) As Boolean
Dim statuszeile As GuiStatusbar
If ss.ActiveWindow.Text = "Fehler" Then
  Cells(zeile, 8).Value = "E"
  Cells(zeile, 9).Value = ss.FindById("wnd[2]/usr/txtMESSTXT1").Text
  ss.FindById("wnd[2]/tbar[0]/btn[0]").press
  ss.FindById("wnd[1]/tbar[0]/btn[12]").press
  IstOrgFalsch = True
End If
End Function


' ---------------------------------------------------------------------------
' 11.4.5.3  Fehler Falscher Wert
' ---------------------------------------------------------------------------
Sub Abschliessen(ss As GuiSession, zeile As Long)
Dim statuszeile As GuiStatusbar
Set statuszeile = ss.FindById("wnd[0]/sbar")
Cells(zeile, 8).Value = statuszeile.Messagetype
Cells(zeile, 9).Value = statuszeile.Text
If statuszeile.Messagetype = "E" Then
  ss.FindById("wnd[0]/tbar[0]/btn[15]").press
  ss.FindById("wnd[1]/usr/btnSPOP-OPTION2").press
  End If
End Sub