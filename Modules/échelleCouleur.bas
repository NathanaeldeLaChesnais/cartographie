Attribute VB_Name = "échelleCouleur"
'Les deux macr suivantes permettent d'éviter un croisement du min et du max pour l'échelle de couleur (par exemple pour la première, quand on clique sur le curseur pour changer le maximum, le minimum de ce curseur est fixé à la valeur du minimum des notes-1 pour éviter qu'on descende en dessous)

Sub minDuMax()
Attribute minDuMax.VB_ProcData.VB_Invoke_Func = " \n14"

    Call initialisation
    ws_map.Unprotect
    ws_map.Shapes.Range(Array("M_MAXSCROLL")).Select
    Selection.Min = ws_param.Range("O3").Value2 + 1
    ws_map.Range("AD54").Select                                             'permet d'éviter que le curseur reste sélectionné en apparence
    ws_map.Protect

End Sub
Sub maxDuMin()
    Call initialisation
    ws_map.Unprotect
    ws_map.Shapes.Range(Array("M_MINSCROLL")).Select
    Selection.Max = ws_param.Range("O1").Value2 - 1
    ws_map.Range("AD54").Select
    ws_map.Protect
    
End Sub
