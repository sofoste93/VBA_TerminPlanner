VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Keep startup small: all application behavior stays in the standard module,
' where learners can run and test each public procedure independently.
Private Sub Workbook_Open()
    SetupSession
    RefreshPlanner
    Worksheets("Dashboard").Activate
End Sub
