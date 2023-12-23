' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
' - SolidWorks 2022 Typbibliothek
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' 12.4.1  Type Koord zur Speicherung eines Punkts 
' ---------------------------------------------------------------------------

Type Koord
 x As Double
 y As Double
End Type
Dim swApp As SldWorks.SldWorks

' ---------------------------------------------------------------------------
' 12.4.2Sub Modellieren
' ---------------------------------------------------------------------------
Sub Modellieren()
Dim xy(0 To 12) As Koord     
xy(0).x = 0: xy(0).y = 0#
xy(1).x = 0.13: xy(1).y = 0#
xy(2).x = 0.13: xy(2).y = 0.01
xy(3).x = 0.12: xy(3).y = 0.01
xy(4).x = 0.12: xy(4).y = 0.04
xy(5).x = 0.08: xy(5).y = 0.04
xy(6).x = 0.08: xy(6).y = 0.06
xy(7).x = 0.06: xy(7).y = 0.06
xy(8).x = 0.06: xy(8).y = 0.02
xy(9).x = -0.04: xy(9).y = 0.04
xy(10).x = -0.04: xy(10).y = 0.07
xy(11).x = -0.06: xy(11).y = 0.07
xy(12).x = -0.06: xy(12).y = 0#

Dim modell As SldWorks.ModelDoc2
Dim defaultTemplate As String

Set swApp = Application.SldWorks
defaultTemplate = swApp.GetUserPreferenceStringValue( _
  swUserPreferenceStringValue_e.swDefaultTemplatePart)  
Set modell = swApp.NewDocument(defaultTemplate, 0, 0, 0)

Dim sketchMgr As SldWorks.SketchManager
Dim linie As SldWorks.SketchLine
Dim i As Integer
Set sketchMgr = modell.SketchManager
With sketchMgr
  .CreateCenterLine 0.15, 0#, 0#, -0.15, 0#, 0#
  For i = 1 To 11
    .CreateLine xy(i).x, xy(i).y, 0#, xy(i + 1).x, xy(i + 1).y, 0#
    DimErzeugen modell, xy(i), xy(i + 1)
  Next i
  Set linie = .CreateLine(xy(i).x, xy(i).y, 0#, xy(1).x, xy(1).y, 0#)
  DimErzeugen modell, xy(i), xy(1)
End With

Dim fixpunkt As SldWorks.SketchPoint
Set fixpunkt = linie.GetStartPoint2
fixpunkt.Select4 False, Nothing
modell.Extension.SelectByID2 "Point1@Ursprung", _
  "EXTSKETCHPOINT", 0, 0, 0, True, 6, Nothing, 0
DimErzeugen modell, xy(0), xy(i)

modell.ViewZoomtofit2
If MsgBox("Rotation anwenden?", vbYesNo) = vbYes Then
  linie.Select4 False, Nothing
  Dim rotationFeature As SldWorks.Feature
  Set rotationFeature = _
modell.FeatureManager.FeatureRevolve2(True, True, False, _
       False, False, False, 0, 0, 6.2831853071796, _
       0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True)
  modell.ShowNamedView2 "*Trimetrisch", 8
  modell.ViewZoomtofit2
End If
DimKoordBerechnen xy(0), xy(0)
End Sub

' ---------------------------------------------------------------------------
' 12.4.3  Sub DimErzeugen zur Bemaßung der Skizze
' ---------------------------------------------------------------------------

Sub DimErzeugen(modell As SldWorks.ModelDoc2, p1 As Koord, p2 As
Koord)
Dim dimKoord As Koord
swApp.SetUserPreferenceToggle _
  swUserPreferenceToggle_e.swInputDimValOnCreate, False
dimKoord = DimKoordBerechnen(p1, p2)
modell.AddDimension2
dimKoord.x, dimKoord.y, 0
swApp.SetUserPreferenceToggle _
  swUserPreferenceToggle_e.swInputDimValOnCreate, True
End Sub

' ---------------------------------------------------------------------------
' 12.4.4  Funktion DimKoordBerechnen zur Positionsberechnung 
'  mit Static Variablen
' ---------------------------------------------------------------------------

Function DimKoordBerechnen(p1 As Koord, p2 As Koord) As Koord
Static maxkoord As Koord
Dim abstandx As Double
Dim abstandy As Double
Dim erg As Koord
Dim OFFSET as Double
OFFSET = 0.015
' reset
If p1.x = p2.x And p1.y = p2.y Then   
 maxkoord.x = 0
 maxkoord.y = 0
End If
' init
If maxkoord.x = 0 Then   
  maxkoord.x = 0.14
End If
If maxkoord.y = 0 Then
  maxkoord.y = 0.08
End If
abstandx = p1.x - p2.x
abstandy = p1.y - p2.y
If (Abs(abstandx) > Abs(abstandy)) Then
  maxkoord.y = maxkoord.y + OFFSET
  erg.x = p1.x - abstandx/2
  erg.y = maxkoord.y
Else
  maxkoord.x = maxkoord.x + OFFSET
  erg.x = maxkoord.x
  erg.y = p1.y – abstandy/2
End If
DimKoordBerechnen = erg
End Function