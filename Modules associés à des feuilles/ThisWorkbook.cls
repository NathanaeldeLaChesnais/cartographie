VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
Call initialisation
ws_map.Unprotect
Application.DisplayFullScreen = True

ws_map.Activate

ws_map.ScrollArea = "T44:AR53"

ws_map.Range("T44:AR53").Select
ActiveWindow.Zoom = False
        With ActiveWindow
            .Zoom = True
            .DisplayHeadings = False
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
            .DisplayWorkbookTabs = False
        End With

ws_map.Protect
End Sub
