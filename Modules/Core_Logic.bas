Attribute VB_Name = "Core_Logic"
Sub ColorHeatMap()
    Call initialisation
    ws_map.Unprotect

    Dim scales As Variant
    Dim scalesColor As Collection
    Dim sID As Variant
    Dim rgbColor As Long
    Dim i As Integer
    
    'niveaux de couleur
    scales = ws_param.Range("E2:E17").Value
    Set scalesColor = New Collection
    For i = 2 To 17
        scalesColor.Add Item:=ws_param.Range("E" & i).Interior.color, Key:=CStr(ws_param.Range("E" & i).Value)
    Next
    
    
    'lecture du tableau de données
    Set d = New data
    d.init
    d.ws.Calculate
    
    For Each sID In d.id                        'On parcours tous les pays
        If Not Left(CStr(sID), 2) = "O_" Then           'on vérifie que ce soit pas un océan
            'On cherche la couleur pour le pays
            For i = UBound(scales) To 1 Step -1                                 'On parcours les notes associé aux niveaux de couleur
                If d.indiceAll(CStr(sID)) > scales(i, 1) Or i = 1 Then          'On compare avec la note du pays
                    rgbColor = scalesColor(CStr(scales(i, 1)))                  'Quand on a trouvé la bonne couleur on finit la boucle
                    i = -1
                End If
            Next
            ws_map.Shapes("S-" & sID).Fill.ForeColor.RGB = rgbColor                                 'On colorie la pays de la bonne couleur
        End If
    Next
    ws_map.Protect
End Sub

