' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 3.3.3 With-Schreibweise
' ---------------------------------------------------------------------------
Sub CollectionDemo()
Sub WithDemo ()
Dim chefs As Collection
Dim chef As Variant
Set chefs = New Collection
With chefs
  .Add "Erwin", "Co-Founder1"
  .Add "Darwin", "Co-Founder2"
  .Add "Franzi", "CFO"
  .Add "Uli", "CEO"
  .Add "Kim", "CIO"
  Debug.Print .Count   ' 5
End With
Set chefs = Nothing
End Sub
