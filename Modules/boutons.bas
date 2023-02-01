Attribute VB_Name = "boutons"
'Les 4 macros suivantes sont associées au 4 boutons "défense etr sécurité et permettent d'afficher ou de masquer ces ponctuels (la première est bien commentée)
Sub boutonAfficherTriangle()            'Niveaux de danger MEAE
    Call initialisation
    ws_map.Unprotect
    If ws_map.Shapes("M_B_TRIANGLE").TextEffect.Text = "Afficher Dangers" Then  'On lit l'état du bouton quand on clique dessus
        ws_map.Range("A3") = 1                                                   ' On écrit en dur l'état du bouton
        ws_map.Shapes("M_B_TRIANGLE").TextEffect.Text = "Cacher Dangers"            'On passe le bouton dans l'autre position
        
    Else
        ws_map.Shapes("M_B_TRIANGLE").TextEffect.Text = "Afficher Dangers"
        ws_map.Range("A3") = 0
        
    End If
    'On applique les mises à jour
    actualiserPonctuel
    ws_map.Protect
End Sub
Sub boutonAfficherCircle()              'Militaires déployés
    Call initialisation
    ws_map.Unprotect
    If ws_map.Shapes("M_B_CIRCLE").TextEffect.Text = "Afficher" Then
        ws_map.Range("A5") = 1
        ws_map.Shapes("M_B_CIRCLE").TextEffect.Text = "Cacher"
        
    Else
        ws_map.Shapes("M_B_CIRCLE").TextEffect.Text = "Afficher"
        ws_map.Range("A5") = 0
        
    End If
    actualiserPonctuel
    ws_map.Protect
End Sub
Sub boutonAfficherConnection()             'alliances de la France
    Call initialisation
    ws_map.Unprotect
    If ws_map.Shapes("M_B_CONNECTION").TextEffect.Text = "Afficher" Then
        ws_map.Range("A6") = 1
        ws_map.Shapes("M_B_CONNECTION").TextEffect.Text = "Cacher"
        
    Else
        ws_map.Shapes("M_B_CONNECTION").TextEffect.Text = "Afficher"
        ws_map.Range("A6") = 0
        
    End If
    actualiserPonctuel
    ws_map.Protect
End Sub
Sub boutonAfficherLB()                     'Zones du
    Call initialisation
    ws_map.Unprotect
    If ws_map.Shapes("M_LB").TextEffect.Text = "Afficher" Then
        ws_map.Range("A4") = 1
        ws_map.Shapes("M_LB").TextEffect.Text = "Cacher"
        
        showMapShapes ("LB-")
    Else
        ws_map.Shapes("M_LB").TextEffect.Text = "Afficher"
        ws_map.Range("A4") = 0
        
        hideMapShapes ("LB-")
        
    End If
    ws_map.Protect
End Sub

Sub boutonABC() 'permet de gérer touts les boutons pour afficher/masquer les ABC
    Call initialisation
    Dim nomABC As String
    
    'On récupère la ABC dont on a cliqué sur le bouton
    nomABC = Right(Application.Caller, InStr(1, StrReverse(Application.Caller), "-") - 1)
    
    'On change le filtre correspondant et la couleur du bouton
    If ActiveSheet.Shapes.Range(Array("M_B-" & nomABC)).Fill.ForeColor.RGB = gris Then  'On regarde la couleur du bouton (indique si activé ou desactivé)
        ActiveWorkbook.SlicerCaches("Segment_ABC").SlicerItems(nomABC).Selected = True
        ActiveSheet.Shapes.Range(Array("M_B-" & nomABC)).Fill.ForeColor.RGB = vert
        
    Else
        ActiveWorkbook.SlicerCaches("Segment_ABC").SlicerItems(nomABC).Selected = False
        ActiveSheet.Shapes.Range(Array("M_B-" & nomABC)).Fill.ForeColor.RGB = gris

    End If
    
    'On applique les modifs à la carte
    ws_map.Unprotect
    ColorHeatMap
    actualiserPonctuel
    ws_map.Protect
End Sub

'Les deux macros qui suivent permettent de gérer l'affichage des ponctuels en fonction de l'activation des boutons
Sub hideMapShapes(prefixe)
    Set ws_map = ThisWorkbook.Sheets("Heat Map")
    ws_map.Unprotect
    tailleprefixe = Len(prefixe)
    
    For Each Shape In ws_map.Shapes.Range(Array("WORLDMAP")).GroupItems         'On parcours toutes les shapes du groupe Worldmap
        If Left(Shape.name, tailleprefixe) = prefixe Then                       'On regarde si la shape a le préfixe qui nous interesse (=est du type de forme qui nous interresse)
                Shape.Visible = msoFalse                                        'Si c'est le cas on la cache
        End If
    Next

    ws_map.Protect
End Sub
Sub showMapShapes(prefixe)
    Set ws_map = ThisWorkbook.Sheets("Heat Map")
    ws_map.Unprotect
    tailleprefixe = Len(prefixe)
    
    For Each Shape In ws_map.Shapes.Range(Array("WORLDMAP")).GroupItems
        If Left(Shape.name, tailleprefixe) = prefixe Then
                Shape.Visible = msoTrue
        End If
    Next
    
    ws_map.Protect
End Sub

