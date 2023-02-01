Attribute VB_Name = "init"
Public ws_map As Worksheet
Public ws_param As Worksheet
Public ws_data As Worksheet
Public s_menu As Shape
Public s_border As Shape
Public m_global As Shape
Public m_fr As Shape
Public s_map As Shape
Public s_autreval As Shape
Public s_diploval As Shape
Public s_fr As Shape
Public showCircle As Boolean
Public showTriangle As Boolean
Public showLB As Boolean
Public showConnection As Boolean
Public d As data
Public vert As Long
Public gris As Long




Sub initialisation()
    Set ws_map = ThisWorkbook.Sheets("Heat Map")
    Set ws_param = ThisWorkbook.Sheets("Parametres")
    Set ws_data = ThisWorkbook.Sheets("SynthèseAffichage")
    Set s_menu = ws_map.Shapes("MENU")
    Set s_border = ws_map.Shapes("MAP_BORDER")
    On Error Resume Next: Set s_map = ws_map.Shapes("WORLDMAP"): On Error GoTo 0
    Set s_fr = ws_map.Shapes("S-FR")
    Set m_global = ws_map.Shapes("M_ACTEUR_GLOBAL")
    Set m_fr = ws_map.Shapes("M_ACTEUR_FR")
    'On initialise les couleurs des boutons qui changent de couleur
    gris = RGB(89, 89, 89)
    vert = RGB(247, 150, 70)        'en fait c'est du orange
    'A partir de là on initialise les variables pour l'affichage ou pas des ponctuels
    If ws_map.Range("A3") = 1 Then
        showTriangle = True
    Else
        showTriangle = False
    End If
    
    If ws_map.Range("A4") = 1 Then
        showLB = True
    Else
        showLB = False
    End If
    If ws_map.Range("A5") = 1 Then
        showCircle = True
    Else
        showCircle = False
    End If
    If ws_map.Range("A6") = 1 Then
        showConnection = True
    Else
        showConnection = False
    End If
End Sub




