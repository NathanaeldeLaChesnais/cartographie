Attribute VB_Name = "miseAJourRégulière"
'attention à bien sélectionner toutes les ABC avant d'executer
Sub redessinerPonctuels()
Attribute redessinerPonctuels.VB_ProcData.VB_Invoke_Func = " \n14"
    Call initialisation
    ThisWorkbook.RefreshAll             'On rafraichit les requètes et les TCD
    Dim sh As Shape
    
    ws_map.Unprotect
    
    ws_map.Shapes.Range(Array("WORLDMAP")).Ungroup          'On dégroupe
    
    For Each sh In ws_map.Shapes                            'On parcours toutes les shapes pour supprimer celles qu'on veut
        If Left(sh.name, 2) = "T-" Or Left(sh.name, 2) = "A-" Or Left(sh.name, 3) = "CE-" Or Left(sh.name, 4) = "TXT-" Or Left(sh.name, 3) = "LB-" Then
                sh.Delete
        End If
    Next
    
    groupMap                                                'On regroupe artificiellement ce quyi reste

    'On construit tout ce qu'on veut (voir les macros dans gestion ponctuels)
    getMapShapesTriangle
    getMapShapesConnection
    getMapShapesCircle
    drawtélérouge
    
    ws_map.Unprotect
    ws_map.Shapes.Range(Array("WORLDMAP")).Ungroup          'On re-détruit le groupe
    groupMap                                                'Pour pouvoir le reformer avec ce qu'on veut
    actualiserPonctuel                                      'On vérifie qu'on affiche bien ce qu'on veut
    ws_map.Protect
End Sub

