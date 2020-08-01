VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'If starting on a new workbook, copy the code starting on the next line to the "ThisWorkbook" object in the VBA IDE
Private Sub Workbook_Open()
    Call Hide_Tabs
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call Hide_Tabs
    'Me.Save
End Sub

