Attribute VB_Name = "muduleDebbug"
'accès par le racourcis CTRL+MAJ+H
'Permet de sortir du mode plein écran bloqué
Sub Debbug()
Attribute Debbug.VB_ProcData.VB_Invoke_Func = "H\n14"
    Call initialisation
    ActiveWindow.DisplayFormulas = False
    ActiveWindow.Zoom = False                                   'On enlève le zoom sur la carte
    ActiveWindow.DisplayVerticalScrollBar = True                'On remet les cureurs horizontaux et verticaux
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayWorkbookTabs = True                     'On réaffiche les onglets (les feuilles)
    ws_map.ScrollArea = ""
    Application.DisplayFullScreen = False                       'On sort du plein écran
    ws_map.Unprotect                                            'On déprotège la feuille
End Sub

