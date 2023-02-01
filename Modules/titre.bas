Attribute VB_Name = "titre"
Sub boutonActeurFrance()
    Call initialisation
    'On modifie le filtre
    With ActiveWorkbook.SlicerCaches("Segment_Acteur")
        .SlicerItems("FRANCE").Selected = True
        .SlicerItems("GLOBAL").Selected = False
    End With
    ws_map.Unprotect
    'On modifie les apparences des boutons
    ActiveSheet.Shapes.Range(Array("M_ACTEUR_FR")).Fill.ForeColor.RGB = vert
    ActiveSheet.Shapes.Range(Array("M_ACTEUR_GLOBAL")).Fill.ForeColor.RGB = gris

    title                   'On fait les modifs sur le texte du titre et de la légende (voir la macro)
    ColorHeatMap            'On refait la coloration
    actualiserPonctuel      'On refait les ponctuels
    ws_map.Protect
End Sub
Sub boutonActeurGlobal()
    Call initialisation
    'On modifie le filtre
    With ActiveWorkbook.SlicerCaches("Segment_Acteur")
        .SlicerItems("GLOBAL").Selected = True
        .SlicerItems("FRANCE").Selected = False
    End With
    ws_map.Unprotect
    'On modifie les apparences des boutons
    ActiveSheet.Shapes.Range(Array("M_ACTEUR_FR")).Fill.ForeColor.RGB = gris
    ActiveSheet.Shapes.Range(Array("M_ACTEUR_GLOBAL")).Fill.ForeColor.RGB = vert

    title                   'On fait les modifs sur le texte du titre et de la légende (voir la macro)
    ColorHeatMap            'On refait la coloration
    actualiserPonctuel      'On refait les ponctuels
    ws_map.Protect
    
End Sub
Sub title()
    Call initialisation
    ws_map.Unprotect
    
    Dim titre As Shape
    Set titre = ws_map.Shapes("M_TITRE")
    
    'On change le titre en fonction du bouton sur lequel on a appuyé
    titre.TextFrame2.TextRange.Characters.Text = ws_map.Shapes(CStr(Application.Caller)).TextFrame2.TextRange.Characters.Text

    'On lit le tableau
    Dim td As Range
    Set td = ThisWorkbook.Sheets("Légende").ListObjects("TD_Légende").DataBodyRange
    Dim i As Integer
    
    For i = 1 To td.Rows.Count
        ws_map.Shapes("M_" & td(i, 5).Value & "LABEL").TextFrame2.TextRange.Characters.Text = td(i, 3).Value
    Next
    
    ws_map.Protect
End Sub
