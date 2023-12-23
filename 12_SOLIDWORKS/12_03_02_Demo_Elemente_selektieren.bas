' ---------------------------------------------------------------------------
' läuft in SOLIDWORKS 
' 
' Benötigte Verweise:
' - keine
' ---------------------------------------------------------------------------

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