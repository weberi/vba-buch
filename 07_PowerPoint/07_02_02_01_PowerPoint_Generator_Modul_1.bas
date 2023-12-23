' ---------------------------------------------------------------------------
' läuft in Excel

' Benötigte Verweise:
'   PowerPoint 
'   Microsoft Scripting Runtime
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 7.2.2.1 Type EingabeT
' ---------------------------------------------------------------------------
Type EingabeT
  ausgabeverz As String
  ausgabename As String
  musterpfad As String
  eingabename As String
  datenadresse As String
  titelzelle As String
  textzelle As String
End Type

' ---------------------------------------------------------------------------
' 7.2.2.1 Function Eingabe
' ---------------------------------------------------------------------------
Function Eingabe() As EingabeT
Dim eingeb As EingabeT

With Worksheets("Einstellungen")
  eingeb.ausgabeverz = .Cells(3, 2).Value
  eingeb.ausgabename = .Cells(4, 2).Value
  eingeb.musterpfad = .Cells(5, 2).Value
  eingeb.eingabename = .Cells(3, 5).Value
  eingeb.datenadresse = .Cells(4, 5).Value & ":" & .Cells(4, 6).Value
  eingeb.titelzelle = .Cells(5, 5).Value
  eingeb.textzelle = .Cells(6, 5).Value
End With
Eingabe = eingeb
End Function

' ---------------------------------------------------------------------------
' 7.2.2.1 Type EinstellungT
' ---------------------------------------------------------------------------
Type EinstellungT
  ausgabedat As String
  daten As Range
  titelbereich As Range
  textbereich As Range
  diagramm As Chart
  musterpfad As String
End Type

' ---------------------------------------------------------------------------
' 7.2.2.1 Function Einstellung
' ---------------------------------------------------------------------------
Function Einstellung(eing As EingabeT) As EinstellungT
Dim einst As EinstellungT
Dim ausgabepfad As String
Dim adresse As String

With eing
  ausgabepfad = ThisWorkbook.Path & "\" & .ausgabeverz
  If Not FSO.FolderExists(ausgabepfad) Then
    FSO.CreateFolder (ausgabepfad)
  End If
  einst.ausgabedat = ausgabepfad & "\" _
    & Format(Now(), "YY_MM_DD_hh_nn_ss") & "_" & .ausgabename
  
  einst.musterpfad = ThisWorkbook.Path & "\" & .musterpfad
  adresse = .eingabename & "!" & .titelzelle & ":" & .titelzelle
  Set einst.titelbereich = Range(adresse)
  
  adresse = .eingabename & "!" & .textzelle & ":" & .textzelle
  Set einst.textbereich = Range(adresse)
  adresse = .eingabename & "!" & eing.datenadresse
  Set einst.daten = Range(adresse)
End With

Set einst.diagramm = Worksheets("Diagrammvorlage") _
  .Shapes("Musterdiagramm").Chart
Einstellung = einst
End Function