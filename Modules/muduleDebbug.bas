Attribute VB_Name = "muduleDebbug"
'acc�s par le racourcis CTRL+MAJ+H
'Permet de sortir du mode plein �cran bloqu�
Sub Debbug()
Attribute Debbug.VB_ProcData.VB_Invoke_Func = "H\n14"
    Call initialisation
    ActiveWindow.DisplayFormulas = False
    ActiveWindow.Zoom = False                                   'On enl�ve le zoom sur la carte
    ActiveWindow.DisplayVerticalScrollBar = True                'On remet les cureurs horizontaux et verticaux
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayWorkbookTabs = True                     'On r�affiche les onglets (les feuilles)
    ws_map.ScrollArea = ""
    Application.DisplayFullScreen = False                       'On sort du plein �cran
    ws_map.Unprotect                                            'On d�prot�ge la feuille
End Sub

