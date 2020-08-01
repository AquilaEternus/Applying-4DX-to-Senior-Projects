VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddLeadMeasure 
   Caption         =   "Add Your Team's Lead Measures"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7410
   OleObjectBlob   =   "AddLeadMeasure.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddLeadMeasure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets(ActiveSheet.Name)
    Me.wigRefComboBox.CLEAR
    
    'Sets wigRefComboBox to the list of WIG ID's and their corresponding descriptions
    For i = 15 To ws.range("A" & Application.Rows.Count).End(xlUp).Row
        Me.wigRefComboBox.AddItem ws.range("A" & i).Value & " - " & ws.range("B" & i).Value
    Next i
    
    Me.descriptionTextBox.Value = ""
    
    'Set assignedTo combobox to the list of students from the scoreboard
    With assignedTo
    .AddItem ActiveSheet.range("A3").Value
    .AddItem ActiveSheet.range("A4").Value
    .AddItem ActiveSheet.range("A5").Value
    .AddItem ActiveSheet.range("A6").Value
    .AddItem "Everyone"
    End With
    
    'Me.pointsTextBox = ""
    
    Me.pointsComboBox.CLEAR
    
    Me.pointsComboBox.AddItem 3
    Me.pointsComboBox.AddItem 4
    Me.pointsComboBox.AddItem 5
    Me.pointsCaptionLabel.Caption = ""

End Sub

'Ensures a team can only assign a limited range of points to a Lead
'to avoid putting too many points in a single Lead.
Private Sub pointsComboBox_Change()
    If Me.pointsComboBox.Value = 3 Then
         Me.pointsCaptionLabel.Caption = "Small Lead"
    ElseIf Me.pointsComboBox.Value = 4 Then
        Me.pointsCaptionLabel.Caption = "Medium Lead"
    ElseIf Me.pointsComboBox.Value = 5 Then
        Me.pointsCaptionLabel.Caption = "Large Lead"
    ElseIf Me.pointsCaptionLabel <> "" Then
        Me.pointsComboBox.Value = ""
        MsgBox "Please enter valid WIG ID."
    End If
End Sub

'Adds a Lead to the "LeadM_Table" with an auto-incremented ID, and adds
'the points of the Lead to the referenced WIG.
Private Sub addLeadMButton_Click()
    ActiveSheet.Unprotect
    
    Dim trimmedWIGStr() As String
    Dim addRow As Integer
    Dim data(5)
    
    'Seperate WIG ID from description in wigRefComboBox
    trimmedWIGStr = Split(Me.wigRefComboBox, " ", 2)
    
    'data(0) = CInt(Me.wigRefTextBox)
    data(0) = CInt(trimmedWIGStr(0))
    data(1) = range("P13").Value
    data(2) = Me.descriptionTextBox.Value
    'data(3) = Me.pointsTextBox.Value
    data(3) = Me.pointsComboBox.Value
    data(4) = Me.assignedTo.Value
    data(5) = "Incomplete"

    'If addToWIGTotal(CInt(Me.wigRefTextBox), Me.pointsTextBox.Value) = 0 Then
    If addToWIGTotal(CInt(trimmedWIGStr(0)), Me.pointsComboBox.Value) = 0 Then
        addDataRow "LeadM_Table", ActiveSheet.Name, data
        range("P13").Value = range("P13").Value + 1
    Else
        MsgBox "Could not find WIG!"
        'Me.wigRefTextBox = ""
        Me.wigRefComboBox = ""
        Me.wigRefComboBox.CLEAR
    End If
    
    addRow = ActiveSheet.ListObjects("LeadM_Table").range.Rows.Count + 13
    range("P" & addRow).Interior.ColorIndex = 44
    
    ActiveSheet.Protect

End Sub

'Adds to the "Total Points" column of the referenced WIG if WIG ID match is
'found in "WIG_Table" and returns a 0 value. Else, returns 1 if no WIG ID is
'matched.
Private Function addToWIGTotal(wigID As Integer, points As String) As Integer
    Dim range As ListObject
    Set range = ActiveSheet.ListObjects("WIG_Table")
    Dim error As Integer
    
    Dim matchResult As Long
    On Error GoTo NoSuchWIG
    matchResult = WorksheetFunction.Match(wigID, range.ListColumns("ID").DataBodyRange, 0)
    
    Dim rowNum As Long
    rowNum = matchResult + 14
    Cells(rowNum, 7).Value = Cells(rowNum, 7).Value + CInt(points)
    Unload AddLeadMeasure
    addToWIGTotal = 0
    Exit Function
    
NoSuchWIG:
    addToWIGTotal = 1
    Exit Function
End Function
