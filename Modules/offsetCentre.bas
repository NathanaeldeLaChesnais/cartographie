Attribute VB_Name = "offsetCentre"
'Macr qui permet de récupérer les offsets des centroïdes si on les déplaces pour le jour où on a besoin de les reconstruire
Sub getOffsetCentre()
    Dim ws_map As Worksheet
    Set ws_map = ThisWorkbook.Sheets("Heat Map")
    Dim ws_offset As Worksheet
    Set ws_offset = ThisWorkbook.Sheets("Parametres")
    Dim shapeName As String

    Dim Xc, Yc, Xs, Ys, L As Double
    Dim i As Integer
    i = 2
    
    For Each Shape In ws_map.Shapes("WORLDMAP").GroupItems

        If Left(Shape.name, 2) = "C-" Then
            On Error Resume Next
            shapeName = "S-" & Mid(Shape.name, 3, 10000)

            If (ws_map.Shapes(shapeName).Width > 0) Then
                Xc = Shape.Left + Shape.Width / 2
                Yc = Shape.Top + Shape.Height / 2
 
                
                Xs = ws_map.Shapes(shapeName).Left + ws_map.Shapes(shapeName).Width / 2
                Ys = ws_map.Shapes(shapeName).Top + ws_map.Shapes(shapeName).Height / 2
                L = ws_map.Shapes(shapeName).Width
                H = ws_map.Shapes(shapeName).Height
                
                ws_offset.Range("A" & i) = shapeName
                ws_offset.Range("B" & i) = CInt(((Xc - Xs) / L) * 100) / 100
                ws_offset.Range("C" & i) = CInt(((Yc - Ys) / H) * 100) / 100
            Else
                ws_offset.Range("A" & i) = shapeName
                ws_offset.Range("B" & i) = 0
                ws_offset.Range("C" & i) = 0
            End If
            i = i + 1
        End If
        
    Next
End Sub
