Attribute VB_Name = "miseAJourRare"
'Rendre les shape des pays cliquables (attention à regarder si on veut rendre les shapes de bateau cliquables ou pas)
Sub shapesCliquables()
    Call initialisation
    ws_map.Unprotect
    
    Dim sID As Variant
    Dim i As Integer

    Set d = New data
    'lecture du tableau de synthèse
    d.init
    d.ws.Calculate
    
    Dim dd As Shape
    For Each sID In d.id
        ws_map.Shapes.Range(Array("S_" & CStr(sID))).Select
        Selection.OnAction = "DetailsPays"
    Next
    ws_map.Protect
End Sub

'Les 3 macros qui suivent concernent les centres cliquables
Sub calculerCentresCliquables()
'
' GestionGroupe Macro
'

'
    Call initialisation
    
    Dim offsetLeft As Collection
    Dim danger As Collection
    Dim offsetTop As Collection
    Dim offsetTopGlobal As Double
    offsetTopGlobal = ws_param.Range("Y2")
    Dim offsetLeftGlobal As Double
    offsetLeftGlobal = ws_param.Range("Z2")
    
    removeMapShapesTriangle
    
    'Gestion des groupes
    Dim A() As String
    Dim ii As Integer
    ii = 0
    

    Set d = New data
    d.init
    Set offsetLeft = New Collection
    Set offsetTop = New Collection
    Set danger = New Collection
    
    Dim ws_off As Worksheet
    Set ws_off = ThisWorkbook.Sheets("Parametres")
    

    Dim j As Integer
    For j = 2 To ws_off.Range("A1000").End(xlUp).Row
        offsetLeft.Add Item:=ws_off.Range("B" & j).Value, Key:=CStr(ws_off.Range("A" & j))
        offsetTop.Add Item:=ws_off.Range("C" & j).Value, Key:=CStr(ws_off.Range("A" & j))
    Next
    
    
    Dim ws_data As Worksheet
    Set ws_data = ThisWorkbook.Sheets("SynthèseAffichage")
    Dim k As Integer
    For k = 4 To ws_data.Range("K1000").End(xlUp).Row
    'chercher la valeur de danger
        danger.Add Item:=ws_data.Range("J" & k).Value, Key:=CStr(ws_data.Range("B" & k))
    Next
        
    ws_map.Unprotect
    Dim dd, qq As Shape
    Dim i As Integer
    'calcul centroide
    For i = 1 To ws_map.Shapes("WORLDMAP").GroupItems.Count
        Set dd = ws_map.Shapes("WORLDMAP").GroupItems(i)
        If Left(dd.name, 6) = "S-O_MI" Then
            Dim sID As String
            sID = Mid(dd.name, 3, 10000)
             
                 Dim X, Y As Double
                 Dim xp, yp As Long
                 'coordonnées centroide
                 X = dd.Left - offsetLeftGlobal + dd.Width / 2 ' + dd.Width * offsetLeft(CStr(sID))
                 Y = dd.Top - offsetTopGlobal + dd.Height / 2 '+ dd.Height * offsetTop(CStr(sID))
                'construction des centroïdes
                Set qq = ws_map.Shapes.AddShape(msoShapeOval, X, Y, 5, 5)
                qq.Visible = msoFalse
                qq.name = "C-" & sID
        End If
    Next

    
    s_border.ZOrder msoBringToFront
    s_menu.ZOrder msoBringToFront
    m_global.ZOrder msoBringToFront
    m_fr.ZOrder msoBringToFront
    ws_map.Protect

End Sub
Sub removeMapShapesCentres()
    Call initialisation
        ws_map.Unprotect
    For Each Shape In ws_map.Shapes '.Range(Array("WORLDMAP")).GroupItems
        If Left(Shape.name, 2) = "C_" Then
                Shape.Delete
        End If
    Next
'    ActiveSheet.Shapes.Range(Array("Triangles")).Delete
        ws_map.Protect
End Sub
Sub hideMapShapesCentre()
    hideMapShapes ("C-")
End Sub
Sub replaceMenu()
    Call initialisation
    s_menu.Top = ws_map.Range("S28").Top
    s_menu.Left = ws_map.Range("S28").Left
End Sub
