Attribute VB_Name = "groupage"
'Les macros suivantes permettent de gérer le groupage quand on fait une mise à jour

Sub groupMap()              'Permet de grouper toutes les shapes qui doivent être dans "WORLDMAP"
Call initialisation
ws_map.Unprotect

Dim shx As Shape, shy As Shape, shxShapes As Object
Dim mapShapes() As String
Dim i As Integer
i = 0

For Each shx In ws_map.Shapes                                                   'On parcours toute les shapes dont certaines sont des groupes
  Set shxShapes = Nothing
  On Error Resume Next: Set shxShapes = shx.GroupItems: On Error GoTo 0         'On regarde si c'est un groupe de shapes et si c'est un groupe on prends les shapes qui sont à l'intérieur
  If shxShapes Is Nothing Then
    Set shxShapes = New Collection: shxShapes.Add shx                           'Si c'est pas un groupe, on rempli avec la shape elle même
  End If
  
  For Each shy In shxShapes                                                     'On parcours toutes les shapes qu'on vient de découvrir et si elles doivent être inclues dans le groupe alors on met leur nom dans la liste "mapshapes"
    If shy.name = "Sea-color 2" Or Left(shy.name, 2) = "T-" Or Left(shy.name, 2) = "C-" Or Left(shy.name, 2) = "S-" Or Left(shy.name, 2) = "A-" Or Left(shy.name, 3) = "CE-" Or Left(shy.name, 4) = "TXT-" Or Left(shy.name, 3) = "LB-" Or Left(shy.name, 2) = "N-" Then
        ReDim Preserve mapShapes(i)
        mapShapes(i) = shy.name
        i = i + 1
    End If
  Next shy
Next shx

  
With ws_map.Shapes.Range(mapShapes)                 'On prend toutes les shapes qui sont dans la liste construite et on en fait uin groupe
    .Group
    .name = "WORLDMAP"
End With


ws_map.Protect
End Sub
