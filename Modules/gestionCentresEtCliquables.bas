Attribute VB_Name = "gestionCentresEtCliquables"
'MessageBox des opérations Mili (quand on clique sur un cercle ou sur un nom d'opération)
Sub DetailsCercle()
    Call initialisation
    Set d = New data
    d.init
         
    Dim sID As String
    nomcomplet = Application.Caller
    If Left(nomcomplet, 3) = "CE-" Then                 'On récupère le sID en fonction de si c'est un cercle ou le nom d'une opération
        sID = Mid(nomcomplet, 4, 100)
    ElseIf Left(nomcomplet, 4) = "TXT-" Then
        sID = Mid(nomcomplet, 5, 100)
    End If
    
    MsgBox d.TextAutre(CStr(sID)), , d.nom(CStr(sID))            'On renvoie un texte prédéfini dans le tableau de synthèse ainsi que le nom du pays en titre
        
End Sub
'MessageBox par pays (Quand on clique sur un pays)
Sub DetailsPays()
    Call initialisation
    Set d = New data
    d.init
    
    sID = Mid(Application.Caller, 3, 100)           'On récupère le "code excel" du pays
    ws_map.Unprotect
    
    Dim cc, qq As Shape
    Set cc = ws_map.Shapes.Range(Array("C-" & sID))             'On récupère le centroïde du pays
    Set qq = ws_map.Shapes.AddShape(msoShapeRoundedRectangle, cc.Left, cc.Top, 500, 44.0217322835)      'On créé la zone de texte
        qq.TextFrame2.TextRange.Characters.Text = d.TextPays(CStr(sID))                                 'On écrit le bon texte (pré-écrit dans une case du tableau de synthèse)
        qq.TextFrame2.TextRange.Font.Size = 22                                                          'On met la bonne mise en forme pour la textebox
        qq.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        qq.Fill.ForeColor.RGB = RGB(255, 255, 255)
        qq.Line.Visible = msoFalse
        qq.Fill.transparency = 0.1
        qq.TextFrame.AutoSize = True
        qq.name = "TB-" & sID                                                                           'On nomme correctement la zone de texte
        qq.OnAction = "removeMapShapeTB"                                                                'On associe une macro à la txtbox pour pouvoir la supprimer quand on clique dessus
    Call centerTB(qq)                                                                                   'On replace si besoin la zone de texte pour qu'elle s'affiche à l'intérieur du cadre où on voit la carte
    ws_map.Protect
End Sub
'Macro permettant de suprimer les txtbox de pays quand on clique dessus
Sub removeMapShapeTB()
    Set ws_map = ThisWorkbook.Sheets("Heat Map")
    ws_map.Unprotect
    
    ws_map.Shapes.Range(Array(CStr(Application.Caller))).Delete

    ws_map.Protect
End Sub
'Macro qui supprime toutes les txtbox de pays quand on zoom ou quand on bouge sur la carte
Sub removeMapShapeTBAll()
    Set ws_map = ThisWorkbook.Sheets("Heat Map")
    Dim elem As Shape
    
    ws_map.Unprotect
    For Each elem In ws_map.Shapes                                  'On parcours toutes les shapes
        If Left(elem.name, 3) = "TB-" Then elem.Delete              'Et on regarde si ce sont des textebox
    Next
    
    ws_map.Protect
End Sub
'Macro appelée à chaque fois qu'on appelle "détails pays" pour vérifier que la txtbox soit dans le cadre de la carte (même fonctionnement que la macro "centerMap")
Sub centerTB(ss As Shape)

    Dim coté As Double
    
    If ss.Left < s_border.Left Then
        ss.Left = s_border.Left
    End If
    coté = s_border.Left + s_border.Width - ss.Width
    If ss.Left > coté Then
         ss.Left = coté
    End If
    
    If ss.Top < s_border.Top Then
        ss.Top = s_border.Top
    End If
    coté = s_border.Top + s_border.Height - ss.Height
    If ss.Top > coté Then
        ss.Top = coté
    End If
    
End Sub
