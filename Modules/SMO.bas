Attribute VB_Name = "SMO"
'Macros associées aux boutons de selection/deselection de toutes les ABC d'un coup (la première est complètement commentée mais pas la 2ème)
Sub boutonABCALL()
    Call initialisation
    ws_map.Unprotect
    
    'On bloque l'actualisation
    Application.Calculation = xlManual
    Application.ScreenUpdating = False

    'On supprime le filtre pour que ça selectionne tout
    ThisWorkbook.Worksheets("TCD").PivotTables("TCD_ValeursAxes").PivotFields("ABC").ClearAllFilters
    
    For Each nomABC In ThisWorkbook.Worksheets("VraiParamètre").ListObjects("TD_ABC").ListColumns("ABC").DataBodyRange.Value       'On parcours toutes les ABC
        ws_map.Shapes.Range(Array("M_B-" & nomABC)).Fill.ForeColor.RGB = vert   'on change le bouton en ABC selectionnée
    Next
    
    Application.Calculation = xlAutomatic
    
    ColorHeatMap        'On recolorie la carte
    actualiserPonctuel  'On remet les ponctuels partout
    Application.ScreenUpdating = True
    ws_map.Protect
        
End Sub
Sub boutonABCDeselec()
    Call initialisation
    ws_map.Unprotect
    
    Application.Calculation = xlManual
    Application.ScreenUpdating = False

    
    
    For Each nomABC In ThisWorkbook.Worksheets("VraiParamètre").ListObjects("TD_ABC").ListColumns("ABC").DataBodyRange.Value        'On parcours toutes les ABC
        If ThisWorkbook.Worksheets("TCD").PivotTables("TCD_ValeursAxes").PivotFields("ABC").PivotItems(nomABC).Visible = True Then ThisWorkbook.Worksheets("TCD").PivotTables("TCD_ValeursAxes").PivotFields("ABC").PivotItems(nomABC).Visible = False    'on change le filtre
        ws_map.Shapes.Range(Array("M_B-" & nomABC)).Fill.ForeColor.RGB = gris   'on change le bouton
    Next
    
    Application.Calculation = xlAutomatic
    Calculate
    ColorHeatMap
    actualiserPonctuel
    
    Application.ScreenUpdating = True
    ws_map.Protect
End Sub

Sub actualiserPonctuel()        'appelée à la fin de l'execution de l'un des boutons "Défense et Sécurité" ou "ABC"
    Call initialisation
    Dim elem As Shape
    Set d = New data
    d.init
    
    For Each elem In ws_map.Shapes.Range(Array("WORLDMAP")).GroupItems
        sID = Right(elem.name, InStr(1, StrReverse(elem.name), "-") - 1)
        prefixe = Left(elem.name, InStr(1, elem.name, "-"))
        'cas des cercles
        If prefixe = "CE-" Then
            If d.nbAutre(CStr(sID)) = 0 Or showCircle = False Then
                elem.Visible = msoFalse
            Else
                Set centre = ws_map.Shapes.Range(Array("C-" & sID))
                elem.Visible = msoTrue
                taillefiguré = Sqr(d.nbAutre(CStr(sID))) * 1.5
                elem.Left = centre.Left - taillefiguré / 2
                elem.Top = centre.Top - taillefiguré / 2
                elem.Width = taillefiguré
                elem.Height = taillefiguré
        End If
        'cas du texte militaire et des bateaux
        ElseIf prefixe = "TXT-" Or Left(elem.name, 4) = "S-O_" Then
            If d.nbAutre(CStr(sID)) = 0 Or showCircle = False Then
                elem.Visible = msoFalse
            Else
                elem.Visible = msoTrue
            End If
        'cas des triangles
        ElseIf prefixe = "T-" Then
            If d.triangle(CStr(sID)) = 0 Or showTriangle = False Then
                elem.Visible = msoFalse
            Else
                elem.Visible = msoTrue
            End If
        'cas des alliances
        ElseIf prefixe = "A-" Then
            If d.triangle(CStr(sID)) = 0 Or showConnection = False Then  ' on se débarasse du problème de faire toutes les différentes alliances
                elem.Visible = msoFalse
            Else
                elem.Visible = msoTrue
            End If
        'Cas du
        ElseIf prefixe = "LB-" Then
            If showLB = False Then
                elem.Visible = msoFalse
            Else
                elem.Visible = msoTrue
            End If
        End If
    Next
            
End Sub

