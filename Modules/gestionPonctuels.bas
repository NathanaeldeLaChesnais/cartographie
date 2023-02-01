Attribute VB_Name = "gestionPonctuels"
'Les fonctions suivantes permettent de dessiner ou de supprimer les diff�rents figur�s ponctuels lors des mises � jours r�guli�res sur les donn�es. Elles sont principalement utilis�es dans les fonctions du module miseAJourR�guli�re
'Les deux premi�res macros sont bien comment�es et les autres sont construites sur le m�me mod�le
Sub getMapShapesTriangle()
    Call initialisation
    Dim dd, qq As Shape
    Dim i As Integer
    Dim sID As String
    Dim X, Y As Double
    Dim colorC As Double
    
    'On commence par supprimer les anciens triangles
    removeMapShapesTriangle
    
    'On lit le tableau de synth�se
    Set d = New data
    d.init

    ws_map.Unprotect
    For i = 1 To ws_map.Shapes("WORLDMAP").GroupItems.Count                 'On parcouys toutes les shapes
        Set dd = ws_map.Shapes("WORLDMAP").GroupItems(i)
        If Left(dd.name, 2) = "C-" Then                                     'On cherche les centro�des
            sID = Mid(dd.name, 3, 10000)
             If d.triangle(CStr(sID)) > 0 Then                                      'On d�termine la couleur du triangle � afficher
                If d.triangle(CStr(sID)) = 10 Then colorC = RGB(0, 255, 0)
                If d.triangle(CStr(sID)) = 1 Then colorC = RGB(255, 255, 0)
                If d.triangle(CStr(sID)) = 2 Then colorC = RGB(255, 125, 0)
                If d.triangle(CStr(sID)) = 3 Then colorC = RGB(255, 0, 0)
                'On r�cup�re les coordonn�es du centroide
                X = dd.Left
                Y = dd.Top
                'On construit le triangle correspondant avec les bonnes propri�t�s
                Set qq = ws_map.Shapes.AddShape(msoShapeIsoscelesTriangle, X - 10, Y - 10, 20, 20)
                qq.Fill.ForeColor.RGB = colorC
                qq.name = "T-" & sID
                qq.OnAction = "DetailsPays"
            End If
        End If
    Next

    s_border.ZOrder msoBringToFront                     'On r�organise les diff�rentes couches pour avoir ce qu'on veut en premier plan
    s_menu.ZOrder msoBringToFront
    m_global.ZOrder msoBringToFront
    m_fr.ZOrder msoBringToFront
    ws_map.Protect

End Sub
Sub removeMapShapesTriangle()
    Set ws_map = ThisWorkbook.Sheets("Heat Map")
    ws_map.Unprotect
    For Each Shape In ws_map.Shapes                             'On parcours toutes les shapes et on supprime celles qui sont des triangles
        If Left(Shape.name, 2) = "T-" Then
                Shape.Delete
        End If
    Next
    ws_map.Protect
End Sub
Sub getMapShapesCircle()

    Call initialisation
    Dim dd, qq As Shape
    Dim i As Integer
    Dim taille As Double
    Dim X, Y As Double
    Dim sID As String
    
    Set d = New data
    d.init

    ws_map.Unprotect
    For i = 1 To ws_map.Shapes("WORLDMAP").GroupItems.Count
        Set dd = ws_map.Shapes("WORLDMAP").GroupItems(i)
        If Left(dd.name, 2) = "C-" Then
            sID = Mid(dd.name, 3, 10000)
             If d.nbAutre(CStr(sID)) > 0 Then
                taille = d.nbAutre(CStr(sID))
                'coordonn�es centroide
                X = dd.Left
                Y = dd.Top
                'construction des cercles
                Set qq = ws_map.Shapes.AddShape(msoShapeOval, X - taille / 2, Y - taille / 2, taille, taille)
                    With qq.Line
                        .Visible = msoTrue
                        .ForeColor.RGB = RGB(79, 129, 189)
                        .transparency = 0
                        .Weight = 3
                    End With
                    With qq.Fill
                        .Visible = msoTrue
                        .ForeColor.RGB = RGB(37, 64, 97)
                        .ForeColor.Brightness = 0.400000006
                        .transparency = 0.5
                        .Solid
                    End With
                qq.OnAction = "DetailsCercle"                       'On associe la macro pour pouvoir avoir une message box
                qq.name = "CE-" & sID
                'construction des �tiquettes
                rayon = Sqr(taille) * 1.5
                Set qq = ws_map.Shapes.AddShape(msoTextOrientationHorizontal, X, Y, 500, 44.0217322835)
                qq.TextFrame2.TextRange.Characters.Text = d.OpAutre(CStr(sID))
                qq.Fill.Visible = msoFalse
                qq.Line.Visible = msoFalse
                qq.name = "TXT-" & sID
                With qq.TextFrame2.TextRange.Font                               'On param�tre les propri�t�s d'affichage du texte
                    .Size = 25
                    .Fill.ForeColor.RGB = RGB(37, 64, 97)
                    .Bold = msoTrue
                    .Caps = msoSmallCaps
                End With
                qq.TextFrame.AutoSize = True                        'On dimensionne la zone de texte pour qu'elle soit adapt�e au texte
                qq.OnAction = "DetailsCercle"                       'On associe la macro pour pouvoir avoir une message box
            End If
        End If
    Next
    s_border.ZOrder msoBringToFront
    s_menu.ZOrder msoBringToFront
    m_global.ZOrder msoBringToFront
    m_fr.ZOrder msoBringToFront
    ws_map.Protect

End Sub

Sub removeMapShapesCircle()
    Set ws_map = ThisWorkbook.Sheets("Heat Map")
    ws_map.Unprotect
    For Each Shape In ws_map.Shapes
        If Left(Shape.name, 3) = "CE-" Or Left(Shape.name, 4) = "TXT-" Then
                Shape.Delete
        End If
    Next
    ws_map.Protect
End Sub
Sub getMapShapesConnection()
    Call initialisation
    Dim dd, qq, cc As Shape
    Dim i As Integer
    Dim sID As String
    Dim X, Y As Double
    Set d = New data
    d.init
    
    ws_map.Unprotect
    
    For i = 1 To ws_map.Shapes("WORLDMAP").GroupItems.Count
        Set dd = ws_map.Shapes("WORLDMAP").GroupItems(i)
        If Left(dd.name, 2) = "C-" Then
            sID = Mid(dd.name, 3, 10000)
            cot�carr� = 10                                  'On peut param�trer ici la taille des carr�s
            'coordonn�es centroide
            X = dd.Left - cot�carr� / 2
            Y = dd.Top - cot�carr� - cot�carr� / 2          'On place toujours eles alliances au dessus des triangles
            If d.UE(CStr(sID)) > 0 Then                     'On regarde si le pays est dans l'UE
                If d.TRUC(CStr(sID)) > 0 Then
                        'construction des triangles
                        Set qq = ws_map.Shapes.AddShape(msoShapeRectangle, X - 5, Y, cot�carr�, cot�carr�)
                        qq.Line.Visible = msoFalse
                        qq.Fill.ForeColor.RGB = RGB(0, 0, 120)
                        qq.name = "A-TRUC-" & sID
                        Set qq = ws_map.Shapes.AddShape(msoShapeRectangle, X + 5, Y, cot�carr�, cot�carr�)
                        qq.Line.Visible = msoTrue
                        qq.Line.ForeColor.RGB = RGB(255, 255, 0)
                        qq.Line.Weight = 2.5
                        qq.Fill.ForeColor.RGB = RGB(113, 113, 255)
                        qq.name = "A-UE-" & sID
                    Else
                        If d.cviolet(CStr(sID)) > 0 Then
                            'construction des triangles
                            Set qq = ws_map.Shapes.AddShape(msoShapeRectangle, X - 5, Y, cot�carr�, cot�carr�)
                            qq.Line.Visible = msoFalse
                            qq.Fill.ForeColor.RGB = RGB(228, 166, 240)
                            qq.name = "A-ECH-" & sID
                            Set qq = ws_map.Shapes.AddShape(msoShapeRectangle, X + 5, Y, cot�carr�, cot�carr�)
                            qq.Line.Visible = msoTrue
                            qq.Line.ForeColor.RGB = RGB(255, 255, 0)
                            qq.Line.Weight = 2.5
                            qq.Fill.ForeColor.RGB = RGB(113, 113, 255)
                            qq.name = "A-UE-" & sID
                        
                        
                        Else
                            'construction des triangles
                            Set qq = ws_map.Shapes.AddShape(msoShapeRectangle, X, Y, cot�carr�, cot�carr�)
                            qq.Line.Visible = msoTrue
                            qq.Line.ForeColor.RGB = RGB(255, 255, 0)
                            qq.Line.Weight = 2.5
                            qq.Fill.ForeColor.RGB = RGB(113, 113, 255)
                            qq.name = "A-UE-" & sID
                        End If
                    End If
                Else   'si pas UE
                    If d.TRUC(CStr(sID)) > 0 Then
                        'construction des triangles
                        Set qq = ws_map.Shapes.AddShape(msoShapeRectangle, X, Y, cot�carr�, cot�carr�)
                        qq.Line.Visible = msoFalse
                        qq.Fill.ForeColor.RGB = RGB(0, 0, 120)
                        qq.name = "A-TRUC-" & sID
                    End If
                    If d.cvert(CStr(sID)) > 0 Then
                        'construction des triangles
                        Set qq = ws_map.Shapes.AddShape(msoShapeRectangle, X, Y, cot�carr�, cot�carr�)
                        qq.Line.Visible = msoFalse
                        qq.Fill.ForeColor.RGB = RGB(180, 9, 233)
                        qq.name = "A-SECU-" & sID
                    End If
                    If d.cbleu(CStr(sID)) > 0 Then
                        'construction des triangles
                        Set qq = ws_map.Shapes.AddShape(msoShapeRectangle, X, Y, cot�carr�, cot�carr�)
                        qq.Line.Visible = msoFalse
                        qq.Fill.ForeColor.RGB = RGB(208, 111, 227)
                        qq.name = "A-COOP-" & sID
                    End If
                    If d.cviolet(CStr(sID)) > 0 Then
                        'construction des triangles
                        Set qq = ws_map.Shapes.AddShape(msoShapeRectangle, X, Y, cot�carr�, cot�carr�)
                        qq.Line.Visible = msoFalse
                        qq.Fill.ForeColor.RGB = RGB(228, 166, 240)
                        qq.name = "A-ECH-" & sID
                    End If
            End If
        End If
    Next

    s_border.ZOrder msoBringToFront
    s_menu.ZOrder msoBringToFront
    m_global.ZOrder msoBringToFront
    m_fr.ZOrder msoBringToFront
    ws_map.Protect

End Sub


Sub removeMapShapesConnection()
    Set ws_map = ThisWorkbook.Sheets("Heat Map")
    ws_map.Unprotect
    For Each Shape In ws_map.Shapes
        If Left(Shape.name, 2) = "A-" Then
                Shape.Delete
        End If
    Next
    ws_map.Protect
End Sub









