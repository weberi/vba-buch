' ---------------------------------------------------------------------------
' läuft in SOLIDWORKS 
' 
' Benötigte Verweise:
' - keine
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 12.3.1  Code-Beispiel: Umgang mit Modellen - PartDoc
' ---------------------------------------------------------------------------

Sub DemoPartDoc()
Dim swApp As SldWorks.SldWorks
Set swApp = Application.SldWorks
Dim teil As PartDoc
On Error GoTo NOT_A_PART
Set teil = swApp.ActiveDoc
Debug.Print "Teil"
      ' Tu etwas mit dem Teil ...
Exit Sub
NOT_A_PART:
  MsgBox ("Nur Teil möglich")
End Sub

' ---------------------------------------------------------------------------
' 12.3.1  Code-Beispiel: Umgang mit Modellen -DemoInterfaceType
' ---------------------------------------------------------------------------
Sub DemoInterfaceType ()
Dim swApp As SldWorks.SldWorks
Set swApp = Application.SldWorks
Dim modell As ModelDoc2

Dim teil As PartDoc

Dim baugruppe As AssemblyDoc
Set modell = swApp.ActiveDoc
Select Case modell.GetType
  Case swDocPART:
    Set teil = modell
    Debug.Print "Teil"      
       ' Tue etwas mit dem Teil ...
  Case swDocASSEMBLY:
    Set baugruppe = modell
    Debug.Print "Baugruppe"
       ' Tue etwas mit der Baugruppe ...
  Case Else
    MsgBox ("Nur Teil oder Baugruppe möglich")
End Select
End Sub

