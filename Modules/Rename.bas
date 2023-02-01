Attribute VB_Name = "Rename"
'macro qui permet de renomer des shapes en masse à l'aide d'un tableau préalablement construit et complété dans une feuille de calcule (Cf le fonctionnement de la macro groupage)
Sub Rename()
Call init
Dim shx As Shape, shy As Shape, shxShapes As Object
Set ws_corr = ThisWorkbook.Sheets("Corrections")
Dim shapeId As Collection
Set shapeId = New Collection
Dim shapesAll As String
shapesAll = ""
For i = 2 To 85
    shapesAll = shapesAll & ws_corr.Range("A" & i).Value
    shapeId.Add Item:=ws_corr.Range("E" & i).Value, Key:=ws_corr.Range("A" & i).Value
Next
For Each shx In ws_map.Shapes
  Set shxShapes = Nothing
  On Error Resume Next: Set shxShapes = shx.GroupItems: On Error GoTo 0
  If shxShapes Is Nothing Then
    Set shxShapes = New Collection: shxShapes.Add shx
  End If
  
  For Each shy In shxShapes
    If InStr(1, shapesAll, shy.name) > 0 Then
        If shapeId(shy.name) = "A SUPPRIMER" Then
            shy.Delete
        Else
            shy.name = shapeId(shy.name)
        End If
    End If
  Next shy
Next shx
  
End Sub
