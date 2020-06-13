VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddProjectForm 
   Caption         =   "Add Senior Project"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "AddProjectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddProjectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddProjectForm_Initialize()
    projectNameTextBox.Value = ""
End Sub

Private Sub addToWorkbookButton_Click()
    'Still needs the case when projectNameTextBox.Value is "" and
    'when project name is already taken
    If isProjectTaken(projectNameTextBox.Value) = False Then
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = projectNameTextBox.Value
        Call createWIGTable
        Call createLeadMTable
        Worksheets(projectNameTextBox.Value).Range("A1").Value = "Total Points:"
        Worksheets(projectNameTextBox.Value).Range("B1").Value = 0#
        Worksheets(projectNameTextBox.Value).Columns("A:M").AutoFit
        Call addWIGButton
        Call addLeadMButton
    End If
End Sub

Public Function isProjectTaken(worksheetName As String) As Boolean
    Dim Work_sheet As Worksheet
    exists = False
    For Each Work_sheet In ThisWorkbook.Worksheets
        If Work_sheet.Name = worksheetName Then
            exists = True
        End If
    Next
End Function
