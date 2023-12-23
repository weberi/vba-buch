' ---------------------------------------------------------------------------
' läuft in Excel
' 
' Benötigte Verweise:
'   keine
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' 3.2  Der Objekttyp Collection (Auflistung)
' ---------------------------------------------------------------------------
Sub CollectionDemo()
Dim chefs As Collection
Dim chef As Variant
Set chefs = New Collection
chefs.Add "Erwin", "Co-Founder1"
chefs.Add "Darwin", "Co-Founder2"
chefs.Add "Franzi", "CFO"
chefs.Add "Uli", "CEO"
chefs.Add "Kim", "CIO"
                                             ' Ausgabe:
Debug.Print chefs.Count                      ' 5
Debug.Print chefs(1)                         ' Erwin
Debug.Print chefs(5)                         ' Kim
Debug.Print chefs("CIO")                     ' Kim
chefs.Remove "Co-Founder1"
chefs.Add "Sam", "Big Boss", before:="CFO" 
Debug.Print "---"                            '  --

For Each chef In chefs
 Debug.Print chef                   ' Darwin Sam Franzi Uli Kim
Next chef
Set chefs = Nothing
End Sub
