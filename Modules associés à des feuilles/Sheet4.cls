VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CheckBox1_Click()
MsgBox CheckBox1.Value
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub M_B_AfficherPoint_Click()

End Sub





Private Sub Worksheet_Activate()
Call initialisation
ws_map.Unprotect
Application.DisplayFullScreen = True

ws_map.ScrollArea = "T44:AR53"

ws_map.Range("T44:AR53").Select
        With ActiveWindow
            .Zoom = True
            .DisplayHeadings = False
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
            .DisplayWorkbookTabs = False
        End With

ws_map.Protect


End Sub

Private Sub Worksheet_Calculate()
'    Call initialisation
'    ws_map.Unprotect
'    Call ColorHeatMap
'    s_diploval.TextFrame2.TextRange.Characters.Text = ws_map.Range("A2").Value
'    s_autreval.TextFrame2.TextRange.Characters.Text = ws_map.Range("A1").Value
'    ws_map.Protect
End Sub
Private Sub Worksheet_Deactivate()
  Application.DisplayFullScreen = False
    Application.DisplayScrollBars = True
    Application.DisplayFullScreen = False
    With ActiveWindow
            .DisplayHeadings = True
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
            .DisplayWorkbookTabs = True
            .Zoom = False
    End With
End Sub


