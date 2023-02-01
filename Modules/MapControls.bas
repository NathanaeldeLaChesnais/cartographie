Attribute VB_Name = "MapControls"
'Les 4 macros suivantes correspondent à l'utilisation des boutons pour se deplacer sur la carte et les deux suivantes aux boutons de zoom
Sub testMoveTopMap()
    Application.ScreenUpdating = False
    Call initialisation
    s_map.Top = s_map.Top + 160                 'On déplace la carte
    Call center_map                             'On vérifie qu'elle ne sorte pas trop du cadre
    removeMapShapeTBAll                         'On supprime les shapes de texte qu'obn aurait pu faire apparaitre en cliquant sur un pays (pour éviter d'avoir à les déplacer)
    Application.ScreenUpdating = True
End Sub
Sub testMoveDownMap()
    Application.ScreenUpdating = False
    Call initialisation
    s_map.Top = s_map.Top - 160
    Call center_map
    removeMapShapeTBAll
    Application.ScreenUpdating = True
End Sub
Sub testMoveLeftMap()
    Application.ScreenUpdating = False
    Call initialisation
    s_map.Left = s_map.Left + 160
    Call center_map
    removeMapShapeTBAll
    Application.ScreenUpdating = True
End Sub
Sub testMoveRigthMap()
    Application.ScreenUpdating = False
    Call initialisation
    s_map.Left = s_map.Left - 160
    Call center_map
    removeMapShapeTBAll
    Application.ScreenUpdating = True
End Sub
Sub testMoveZoomInMap()
    Application.ScreenUpdating = False
    Call initialisation
    If s_map.Width < 20000 Then                 'On vérifie qu'on ne zoom pas déjà trop
        s_map.Width = s_map.Width * 2
        s_map.Left = s_border.Left + (s_border.Width / 2) - (((s_border.Left + (s_border.Width / 2)) - s_map.Left) / s_map.Width) * (s_map.Width * 2)
        s_map.Height = s_map.Height * 2
        s_map.Top = s_border.Top + (s_border.Height / 2) - (((s_border.Top + (s_border.Height / 2)) - s_map.Top) / s_map.Height) * (s_map.Height * 2)
        Call center_map
    End If
    actualiserZoomToutesFormes
    Application.ScreenUpdating = True
End Sub
Sub testMoveZoomOutMap()
    Application.ScreenUpdating = False
    Call initialisation
    s_map.Width = s_map.Width * 0.5
    s_map.Left = s_border.Left + (s_border.Width / 2) - (((s_border.Left + (s_border.Width / 2)) - s_map.Left) / s_map.Width) * (s_map.Width * 0.5)
    s_map.Height = s_map.Height * 0.5
    s_map.Top = s_border.Top + (s_border.Height / 2) - (((s_border.Top + (s_border.Height / 2)) - s_map.Top) / s_map.Height) * (s_map.Height * 0.5)
    Call center_map
    actualiserZoomToutesFormes
    Application.ScreenUpdating = True
End Sub

Sub center_map()
    H = s_map.Height
    W = s_map.Width
    Dim coté As Double
    
    If s_map.Width < s_border.Width Or s_map.Height < s_border.Height Then
        s_map.Width = s_border.Width
        s_map.Height = s_border.Height
    End If
    
    If s_map.Left > s_border.Left Then
        s_map.Left = s_border.Left
    End If
    coté = s_border.Left + s_border.Width - s_map.Width
    If s_map.Left < coté Then
         s_map.Left = coté
    End If
    
    If s_map.Top > s_border.Top Then
        s_map.Top = s_border.Top
    End If
    coté = s_border.Top + s_border.Height - s_map.Height
    If s_map.Top < coté Then
        s_map.Top = coté
    End If
    

End Sub

Sub actualiserZoomToutesFormes()
    Call initialisation
    
    removeMapShapeTBAll
    
    ws_map.Unprotect
    
    Set tableau = New data
    tableau.init
    
    Dim sID As String
    Dim texte As Shape
    
    For i = 1 To ws_map.Shapes("WORLDMAP").GroupItems.Count
        Set dd = ws_map.Shapes("WORLDMAP").GroupItems(i)
        
        'Gestion du cas des triangles
        If Left(dd.name, 2) = "T-" Then
        taillesufix = InStr(1, StrReverse(dd.name), "-")
            sID = Right(dd.name, taillesufix - 1)
                Set forme = ws_map.Shapes.Range(Array(dd.name))
                Set centre = ws_map.Shapes.Range(Array("C-" & sID))
                taillefiguré = 20
                forme.Left = centre.Left - taillefiguré / 2
                forme.Top = centre.Top - taillefiguré / 2
                forme.Width = taillefiguré
                forme.Height = taillefiguré
        End If
        
        'Gestion du cas des militaires
        If Left(dd.name, 3) = "CE-" Then
        taillesufix = InStr(1, StrReverse(dd.name), "-")
            sID = Right(dd.name, taillesufix - 1)
                Set forme = ws_map.Shapes.Range(Array(dd.name))
                Set centre = ws_map.Shapes.Range(Array("C-" & sID))
                Set texte = ws_map.Shapes("TXT-" & sID)
                taillefiguré = Sqr(tableau.nbAutre(CStr(sID))) * 1.5
                forme.Left = centre.Left - taillefiguré / 2
                forme.Top = centre.Top - taillefiguré / 2
                forme.Width = taillefiguré
                forme.Height = taillefiguré
                texte.Left = centre.Left
                texte.Top = centre.Top
                texte.TextFrame.AutoSize = True
        End If
        If Left(dd.name, 2) = "C-" Then
        taillesufix = InStr(1, StrReverse(dd.name), "-")
            sID = Right(dd.name, taillesufix - 1)
                Set forme = ws_map.Shapes.Range(Array(dd.name))
                Set centre = ws_map.Shapes.Range(Array("C-" & sID))
                taillefiguré = 25
                forme.AutoShapeType = msoShapeRectangle
                forme.Visible = msoFalse
                forme.Width = taillefiguré
                forme.Height = taillefiguré
        End If
        
        'gestion du cas alliance
        If Left(dd.name, 2) = "A-" Then
            taillefiguré = 10
            taillesufix = InStr(1, StrReverse(dd.name), "-")
            sID = Right(dd.name, taillesufix - 1)
            Set centre = ws_map.Shapes.Range(Array("C-" & sID))
            If Left(dd.name, 5) = "A-UE-" Then
                dd.Left = centre.Left
                dd.Top = centre.Top - taillefiguré
                dd.Width = taillefiguré
                dd.Height = taillefiguré
            Else
                dd.Left = centre.Left - taillefiguré
                dd.Top = centre.Top - taillefiguré
                dd.Width = taillefiguré
                dd.Height = taillefiguré
            End If
        End If
    Next
    ws_map.Protect
End Sub



