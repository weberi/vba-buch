' ---------------------------------------------------------------------------
' läuft in Word

' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 6.6.4  Etiketten-Tool - Konstanten
' ---------------------------------------------------------------------------
Const ETIKETTENDRUCKER As String = "Microsoft Print to PDF"
Const NUM_PROPNAME As String = "Inventarnummer"
Const NUM_INITIAL As Long = 100


' ---------------------------------------------------------------------------
' 6.6.4  Etiketten-Tool - Sub Etikett
' ---------------------------------------------------------------------------
Sub Etikett()
Dim dok As Document
Dim tool As Document
Dim invNum As Long
Dim barcodeDef As String

Set tool = ThisDocument
invNum = DokEigenschaft(tool.CustomDocumentProperties, _
    NUM_PROPNAME, NUM_INITIAL)

Set dok = Documents.Add
With dok.PageSetup
 .PageWidth = CentimetersToPoints(7)
  .PageHeight = CentimetersToPoints(5)
  .RightMargin = 2
  .LeftMargin = 2
  .TopMargin = 2
  .BottomMargin = 2
End With

barcodeDef = "DISPLAYBARCODE """ & invNum & """ QR \q 3"
dok.Fields.Add dok.Range(0, 0), wdFieldEmpty, barcodeDef, False
dok.Content.InsertAfter (vbCrLf & "Inventar " & invNum)
' vbCrLf = Zeilenumbruch

With dok.Content.Paragraphs.Format
  .Alignment = wdAlignParagraphCenter
  .SpaceAfter = 0
  .SpaceBefore = 0
End With

With Dialogs(wdDialogFilePrintSetup)
  .Printer = ETIKETTENDRUCKER
  .DoNotSetAsSysDefault = True
  .Execute
End With

dok.PrintOut
dok.Close False
tool.CustomDocumentProperties(NUM_PROPNAME).Value = invNum + 1
tool.Save
End Sub

' ---------------------------------------------------------------------------
' 6.6.4  Etiketten-Tool - Function DokEigenschaft
' ---------------------------------------------------------------------------
Function DokEigenschaft(props As DocumentProperties, _
  pname As String, initial As Long) As Long
Err.Clear
On Error Resume Next
DokEigenschaft = props(pname).Value
If Err.Number <> 0 Then
  props.Add pname, False, msoPropertyTypeNumber, initial
  DokEigenschaft = initial
  Err.Clear
End If
End Function