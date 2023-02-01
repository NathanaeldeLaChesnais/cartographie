Attribute VB_Name = "gestionLivreBlanc"
Type télérouge
    name As Variant
    X As Collection
    Y As Collection
    L As Collection
    H As Collection
    color As Collection
End Type
'Permet de dessiner les zones du en cas de mise à jour
Sub drawtélérouge()
    Call initialisation
    ws_map.Unprotect
    Dim lb As télérouge
    Set lb.X = New Collection
    Set lb.Y = New Collection
    Set lb.L = New Collection
    Set lb.H = New Collection
    Set lb.color = New Collection
    lb.name = ws_param.Range("F2:F" & ws_param.Range("F1000").End(xlUp).Row).Value
    
    For i = 2 To ws_param.Range("F1000").End(xlUp).Row
        lb.X.Add Item:=ws_param.Range("G" & i).Value, Key:=CStr(ws_param.Range("F" & i).Value)
        lb.Y.Add Item:=ws_param.Range("H" & i).Value, Key:=CStr(ws_param.Range("F" & i).Value)
        lb.H.Add Item:=ws_param.Range("I" & i).Value, Key:=CStr(ws_param.Range("F" & i).Value)
        lb.L.Add Item:=ws_param.Range("J" & i).Value, Key:=CStr(ws_param.Range("F" & i).Value)
        lb.color.Add Item:=ws_param.Range("K" & i).Value, Key:=CStr(ws_param.Range("F" & i).Value)
    Next
    
    Dim X, Y, L, H As Double
    Dim s_lb As Shape
    For Each elem In lb.name
        X = s_fr.Left + s_fr.Width * lb.X(CStr(elem))
        Y = s_fr.Top + s_fr.Height * lb.Y(CStr(elem))
        L = s_fr.Width * lb.L(CStr(elem))
        H = s_fr.Height * lb.H(CStr(elem))
        Set s_lb = ws_map.Shapes.AddShape(msoShapeRoundedRectangle, X, Y, H, L)
        s_lb.Fill.Visible = msoFalse
        s_lb.Line.Visible = msoTrue
        s_lb.Line.ForeColor.RGB = lb.color(CStr(elem))
        s_lb.Line.transparency = 0
        s_lb.Line.Weight = 5
        s_lb.name = CStr(elem)
    Next
    s_border.ZOrder msoBringToFront
    s_menu.ZOrder msoBringToFront
    m_global.ZOrder msoBringToFront
    m_fr.ZOrder msoBringToFront
    ws_map.Protect
End Sub
'Permet de supprimer les zones de n cas de mise à jour
Sub removeMapShapesLB()
    Call initialisation
        ws_map.Unprotect
    For Each Shape In ws_map.Shapes
        If Left(Shape.name, 3) = "LB-" Then
                Shape.Delete
        End If
    Next
        ws_map.Protect
End Sub
'Permet de récupérer les dessins des zones du si on veut modifier leurs tailles et placement (commencer par les modifier, executer la macro et après la macro "draxMapShapeLB" fera les bons ddessins)( quelques détails à vérifier avant de l'executer comme du nomage)
Sub gettélérouge()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Heat Map")
    Dim shapeName As String

    Dim X, Y, L, H, lb_x, lb_y, lb_L, lb_H, lb_color As Double
    Dim i As Integer
    i = 2
    
    'On commence par récupérer les paramètres de la shape France car tout est calculé par rapport à elle
    X = ws.Shapes("S_FR").Left
    Y = ws.Shapes("S_FR").Top
    L = ws.Shapes("S_FR").Width
    H = ws.Shapes("S_FR").Height
    For Each Shape In ws.Shapes
        If Left(Shape.name, 7) = "LB-zone" Then
            On Error Resume Next
            shapeName = Shape.name
            lb_x = (Shape.Left - X) / L
            lb_y = (Shape.Top - Y) / H
            lb_L = Shape.Width / L
            lb_H = Shape.Height / H
            lb_color = Shape.Line.ForeColor.RGB
            ws_param.Range("F" & i) = Shape.name
            ws_param.Range("G" & i) = lb_x
            ws_param.Range("H" & i) = lb_y
            ws_param.Range("I" & i) = lb_L
            ws_param.Range("J" & i) = lb_H
            ws_param.Range("K" & i) = lb_color
            i = i + 1
        End If
            

        
    Next
End Sub
