Attribute VB_Name = "�chelleCouleur"
'Les deux macr suivantes permettent d'�viter un croisement du min et du max pour l'�chelle de couleur (par exemple pour la premi�re, quand on clique sur le curseur pour changer le maximum, le minimum de ce curseur est fix� � la valeur du minimum des notes-1 pour �viter qu'on descende en dessous)

Sub minDuMax()
Attribute minDuMax.VB_ProcData.VB_Invoke_Func = " \n14"

    Call initialisation
    ws_map.Unprotect
    ws_map.Shapes.Range(Array("M_MAXSCROLL")).Select
    Selection.Min = ws_param.Range("O3").Value2 + 1
    ws_map.Range("AD54").Select                                             'permet d'�viter que le curseur reste s�lectionn� en apparence
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
