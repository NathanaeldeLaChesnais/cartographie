Attribute VB_Name = "MapControls"
'Les 4 macros suivantes correspondent � l'utilisation des boutons pour se deplacer sur la carte et les deux suivantes aux boutons de zoom
Sub testMoveTopMap()
    Application.ScreenUpdating = False
    Call initialisation
    s_map.Top = s_map.Top + 160                 'On d�place la carte
    Call center_map                             'On v�rifie qu'elle ne sorte pas trop du cadre
    removeMapShapeTBAll                         'On supprime les shapes de texte qu'obn aurait pu faire apparaitre en cliquant sur un pays (pour �viter d'avoir � les d�placer)
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
    If s_map.Width < 20000 Then                 'On v�rifie qu'on ne zoom pas d�j� trop
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
    Dim cot� As Double
    
    If s_map.Width < s_border.Width Or s_map.Height < s_border.Height Then
        s_map.Width = s_border.Width
        s_map.Height = s_border.Height
    End If
    
    If s_map.Left > s_border.Left Then
        s_map.Left = s_border.Left
    End If
    cot� = s_border.Left + s_border.Width - s_map.Width
    If s_map.Left < cot� Then
         s_map.Left = cot�
    End If
    
    If s_map.Top > s_border.Top Then
        s_map.Top = s_border.Top
    End If
    cot� = s_border.Top + s_border.Height - s_map.Height
    If s_map.Top < cot� Then
        s_map.Top = cot�
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
                taillefigur� = 20
                forme.Left = centre.Left - taillefigur� / 2
                forme.Top = centre.Top - taillefigur� / 2
                forme.Width = taillefigur�
                forme.Height = taillefigur�
        End If
        
        'Gestion du cas des militaires
        If Left(dd.name, 3) = "CE-" Then
        taillesufix = InStr(1, StrReverse(dd.name), "-")
            sID = Right(dd.name, taillesufix - 1)
                Set forme = ws_map.Shapes.Range(Array(dd.name))
                Set centre = ws_map.Shapes.Range(Array("C-" & sID))
                Set texte = ws_map.Shapes("TXT-" & sID)
                taillefigur� = Sqr(tableau.nbAutre(CStr(sID))) * 1.5
                forme.Left = centre.Left - taillefigur� / 2
                forme.Top = centre.Top - taillefigur� / 2
                forme.Width = taillefigur�
                forme.Height = taillefigur�
                texte.Left = centre.Left
                texte.Top = centre.Top
                texte.TextFrame.AutoSize = True
        End If
        If Left(dd.name, 2) = "C-" Then
        taillesufix = InStr(1, StrReverse(dd.name), "-")
            sID = Right(dd.name, taillesufix - 1)
                Set forme = ws_map.Shapes.Range(Array(dd.name))
                Set centre = ws_map.Shapes.Range(Array("C-" & sID))
                taillefigur� = 25
                forme.AutoShapeType = msoShapeRectangle
                forme.Visible = msoFalse
                forme.Width = taillefigur�
                forme.Height = taillefigur�
        End If
        
        'gestion du cas alliance
        If Left(dd.name, 2) = "A-" Then
            taillefigur� = 10
            taillesufix = InStr(1, StrReverse(dd.name), "-")
            sID = Right(dd.name, taillesufix - 1)
            Set centre = ws_map.Shapes.Range(Array("C-" & sID))
            If Left(dd.name, 5) = "A-UE-" Then
                dd.Left = centre.Left
                dd.Top = centre.Top - taillefigur�
                dd.Width = taillefigur�
                dd.Height = taillefigur�
            Else
                dd.Left = centre.Left - taillefigur�
                dd.Top = centre.Top - taillefigur�
                dd.Width = taillefigur�
                dd.Height = taillefigur�
            End If
        End If
    Next
    ws_map.Protect
End Sub



