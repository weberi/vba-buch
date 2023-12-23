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
      ' Tue etwas mit dem Teil ...
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

' ---------------------------------------------------------------------------
' 12.3.2  Code-Beispiel: Elemente selektieren
' ---------------------------------------------------------------------------
Sub DemoSelect()
Dim swApp As Object
Dim SelMgr As SelectionMgr
Dim boolStatus As Boolean
Dim teil As PartDoc
Dim sk As SketchSegment

Set swApp = Application.SldWorks
Set teil = swApp.ActiveDoc

boolStatus = teil.Extension _
 .SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
Debug.Print boolStatus

boolStatus = teil.Extension _
 .SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
Debug.Print boolStatus

Set SelMgr = teil.SelectionManager
Set sk = SelMgr.GetSelectedObject(1)

Debug.Print sk.GetName
Debug.Print sk.GetLength
sk.Color = vbYellow
Set sk = SelMgr.GetSelectedObject(2)

Debug.Print sk.GetName
Debug.Print sk.GetLength
sk.Color = vbGreen
sk.DeSelect
End Sub