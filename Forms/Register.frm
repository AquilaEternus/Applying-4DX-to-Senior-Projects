VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Register 
   Caption         =   "Register"
   ClientHeight    =   9375.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14190
   OleObjectBlob   =   "Register.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CLEAR_Click()

    myName.Value = ""
    projectName.Value = ""
    username.Value = ""
    pw.Value = ""

End Sub

'Generates a project's 4DX if it wasn't added already, and store the student registration data.
Private Sub OK_Click()

    If myName.Value = "" Or projectName.Value = "" Or username.Value = "" Or pw.Value = "" Then
    MsgBox "Please Fill Out All Text Boxes", vbCritical
    End If
    Dim i As Integer
    Dim store As Worksheet
    Set store = ThisWorkbook.Sheets("InformationInput")
    ThisWorkbook.Unprotect
    
    If ProjectExists(projectName.Value) = False Then
        'Call on the functions from the Modules Team4DXGenerator and Team4DXButtons to generate
        'a teams 4DX sheets
        Sheets.add.Name = projectName.Value
        range("A3").Value = myName
        range("C3").Value = 0
        Call createTeam(projectName)
        Call createWIGtable
        Call add_WIG
        Call createLeadMeasureTable
        Call add_Lead
        Call edit_WIG
        Call edit_Lead
        Call add_Logout
        Call createContributionChart(projectName.Value)
    Else
        'Project already exists so just place registree name in the first empty spot of the Scoreboard
        Worksheets(projectName.Value).Activate
        Worksheets(projectName.Value).Unprotect
        If IsEmpty(range("A4").Value) = True Then
            range("A4").Value = myName
            range("C4").Value = 0
        ElseIf IsEmpty(range("A5").Value) = True Then
            range("A5").Value = myName
            range("C5").Value = 0
        ElseIf IsEmpty(range("A6").Value) = True Then
            range("A6").Value = myName
            range("C6").Value = 0
        Else
            'If no spot is empty, team reach max capacity and cannot be registered too.
            Worksheets("Start").Activate
            MsgBox "Team is already full."
            Call Hide_Tabs
            Unload Register
            Exit Sub
        End If
    End If
    
    'Store the registration data in the "InformationInput" sheet
    For i = 2 To store.range("A" & Application.Rows.Count).End(xlUp).Row
        If store.range("A" & i).Value = myName.Value Then
            store.range("C" & i).Value = projectName.Value
            store.range("D" & i).Value = username.Value
            store.range("E" & i).Value = pw.Value
        End If
    Next i
    
    Worksheets(projectName.Value).Protect
    Call Hide_Tabs
    Unload Register
    Worksheets("Start").Activate

End Sub

Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    Dim i As Integer
    Set ws = ThisWorkbook.Sheets("InformationInput")


    For i = 2 To ws.range("A" & Application.Rows.Count).End(xlUp).Row
        If ws.range("A" & i).Value <> "" Then
            myName.AddItem ws.range("A" & i).Value
        End If
    Next i
    
    For i = 2 To ws.range("B" & Application.Rows.Count).End(xlUp).Row
        If ws.range("B" & i).Value <> "" Then
            projectName.AddItem ws.range("B" & i).Value
        End If
    Next i


End Sub

'Is needed to check if project needs to be generated for the first time upon user registering
Function ProjectExists(ByVal wsName As String) As Boolean

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If Application.Proper(ws.Name) = Application.Proper(wsName) Then
            ProjectExists = True
            Exit Function
        End If
    Next ws
    ProjectExists = False

End Function


