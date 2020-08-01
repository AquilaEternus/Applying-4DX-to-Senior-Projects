VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModifyLead 
   Caption         =   "Edit or Delete Team's Lead Measures"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6570
   OleObjectBlob   =   "ModifyLead.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModifyLead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ActiveSheet.Unprotect
    Dim ws As Worksheet
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets(ActiveSheet.Name)
    Me.wigIDComboBox.CLEAR
    
    For i = 15 To ws.range("A" & Application.Rows.Count).End(xlUp).Row
        Me.wigIDComboBox.AddItem ws.range("A" & i).Value
    Next i
    
    Me.leadIDComboBox.CLEAR
    For i = 15 To ws.range("L" & Application.Rows.Count).End(xlUp).Row
        Me.leadIDComboBox.AddItem ws.range("L" & i).Value
    Next i
    
    Call clearValues
    ActiveSheet.Protect
End Sub

'Is called whenever an error occurs in another function of this form on initialize.
Private Sub clearValues()
    Me.wigIDComboBox = ""
    Me.leadIDComboBox = ""
    Me.descriptionTextBox = ""
End Sub

'Refreshes the list of WIG ID's that may be selected by the user
Private Sub wigIDComboBox_Change()
    Dim index As Long
    Me.leadIDComboBox.CLEAR
    
    If Me.wigIDComboBox = "" Then
        For i = 15 To ActiveSheet.range("L" & Application.Rows.Count).End(xlUp).Row
            Me.leadIDComboBox.AddItem ActiveSheet.range("L" & i).Value
        Next i
        Exit Sub
    End If
     
    On Error GoTo NotAnInt
    index = findWIG(CInt(Me.wigIDComboBox), ActiveSheet.Name)
    
    'Narrows down the list of lead's that appear in leadIdComboBox when a WIG is selected
    If index <> -1 Then
        Dim tbl As ListObject
        Dim lastRow As Long
        
        Set tbl = ActiveSheet.ListObjects("LeadM_Table")
        If tbl.ListRows.Count > 0 Then
            Dim leadRow As ListRow
            lastRow = tbl.ListRows.Count
            For index = 1 To lastRow
                Set leadRow = tbl.ListRows(index)
                If Intersect(leadRow.range, tbl.ListColumns("WIG ID").range).Value = CStr(Me.wigIDComboBox) Then
                    'Do whatever clean up is necessary, like removing points given to somewhere
                    Me.leadIDComboBox.AddItem leadRow.range.Cells(1, 2).Value
                End If
            Next index
        End If
    Else
        MsgBox "Sorry, but WIG does not exist. Please, select an existing WIG."
        Call clearValues
    End If
    Exit Sub
    
NotAnInt:
    MsgBox "Please enter a valid integer."
    Call clearValues
    Exit Sub
End Sub

'If "Remove" is clicked, deleteLeadMeasure() gets called and cleans up any points it may
'have given to the WIG it belonged to and the Scoreboard.
Private Sub removeLeadButton_Click()
    ActiveSheet.Unprotect
    If Me.leadIDComboBox.Value = "" Then
        MsgBox "Please select or add Lead Measure goal to modify.", vbCritical, "Error"
        Exit Sub
    End If
    Call deleteLeadMeasure(Me.leadIDComboBox, ActiveSheet.Name)
    Call clearValues
    ActiveSheet.Protect
End Sub

'If "Update" button is clicked, the changes made to the Description will be reflected on
'the lead measure table.
Private Sub updateLeadButton_Click()
    ActiveSheet.Unprotect
    If Me.leadIDComboBox.Value = "" Then
        Me.completeLeadCheckBox.Value = False
        MsgBox "Please select Lead Measure goal to modify or close form.", vbCritical, "Error"
        Exit Sub
    End If
    Dim rowNum As Integer
    index = findLeadMeasure(CInt(Me.leadIDComboBox), ActiveSheet.Name)
    rowNum = index + 14
    ActiveSheet.range("M" & rowNum).Value = Me.descriptionTextBox.Value
    Unload ModifyLead
    ActiveSheet.Protect
End Sub

'If checked complete, points will be added to the "Aquired Points" column of a WIG and added
'to which ever members it is to Scoreboard. It will also change the look and text of
'the lead measure's "Status" column
Private Sub completeLeadCheckBox_Click()
    ActiveSheet.Unprotect
    
    If Me.leadIDComboBox.Value = "" Then
        Me.completeLeadCheckBox.Value = False
        MsgBox "Please select or add Lead Measure goal to modify.", vbCritical, "Error"
        Exit Sub
    End If
    Dim rowNum As Integer
    Dim index As Integer
    Dim wigID As String
    Dim points As Integer
    Dim tbl As ListObject
    Dim i As Integer
    index = findLeadMeasure(CInt(Me.leadIDComboBox), ActiveSheet.Name)
    rowNum = index + 14
     
    'Change to complete and add the appropriate points
    If Me.completeLeadCheckBox.Value = True And ActiveSheet.range("P" & rowNum).Value = "Incomplete" Then
        ActiveSheet.range("P" & rowNum).Value = "Complete"
        ActiveSheet.range("P" & rowNum).Interior.ColorIndex = 35
        wigID = ActiveSheet.Cells(rowNum, 11).Value
        points = ActiveSheet.Cells(rowNum, 14).Value
        Set tbl = ActiveSheet.ListObjects("WIG_Table")

        matchResult = WorksheetFunction.Match(CInt(wigID), tbl.ListColumns("ID").DataBodyRange, 0)
        rowNum = matchResult + 14
    
        'Add completed lead measure points to WIG's "Aquired Points"
        If ActiveSheet.Cells(rowNum, 6).Value < ActiveSheet.Cells(rowNum, 7).Value Then
            ActiveSheet.Cells(rowNum, 6).Value = ActiveSheet.Cells(rowNum, 6).Value + CInt(points)
        End If
        
        rowNum = index + 14
        'Add the Values to the scoreboard
        If ActiveSheet.range("O" & rowNum).Value = "Everyone" Then
            For i = 3 To 6
                If range("A" & i).Value <> "" Then
                    range("C" & i).Value = range("C" & i).Value + points
                End If
            Next i
            range("C7").Value = range("C7").Value + points
        Else
            For i = 3 To 6
                If ActiveSheet.range("O" & rowNum).Value = range("A" & i).Value Then
                    range("C" & i).Value = range("C" & i).Value + points
                End If
            Next i
            range("C7").Value = range("C7").Value + points
        End If
        ActiveSheet.Protect
        Unload ModifyLead
    End If
    
    'Revert changes and subtract what was added from previous If statement
    If Me.completeLeadCheckBox.Value = False Then
        ActiveSheet.range("P" & rowNum).Value = "Incomplete"
        ActiveSheet.range("P" & rowNum).Interior.ColorIndex = 44
        wigID = ActiveSheet.Cells(rowNum, 11).Value
        points = ActiveSheet.Cells(rowNum, 14).Value
        Set tbl = ActiveSheet.ListObjects("WIG_Table")
        'On Error Resume Next
        matchResult = WorksheetFunction.Match(CInt(wigID), tbl.ListColumns("ID").DataBodyRange, 0)
        rowNum = matchResult + 14
    
        'Remove added lead measure points from WIG's "Aquired Points"
        ActiveSheet.Cells(rowNum, 6).Value = ActiveSheet.Cells(rowNum, 6).Value - CInt(points)
        
        rowNum = index + 14
        'Subtract the values from the scoreboard
        If ActiveSheet.range("O" & rowNum).Value = "Everyone" Then
            For i = 3 To 6
                If range("A" & i).Value <> "" Then
                    range("C" & i).Value = range("C" & i).Value - points
                End If
            Next i
            range("C7").Value = range("C7").Value - points
        Else
            For i = 3 To 6
                If ActiveSheet.range("O" & rowNum).Value = range("A" & i).Value Then
                    range("C" & i).Value = range("C" & i).Value - points
                End If
            Next i
            range("C7").Value = range("C7").Value - points
        End If
        ActiveSheet.Protect
        Unload ModifyLead
    End If
    'ActiveSheet.Protect
End Sub


'Refreshes the list of Lead ID's
Private Sub leadIDComboBox_Change()
    Dim index As Long
    
    If Me.leadIDComboBox = "" Then
        For i = 15 To ActiveSheet.range("L" & Application.Rows.Count).End(xlUp).Row
            Me.leadIDComboBox.AddItem ActiveSheet.range("L" & i).Value
        Next i
        Me.descriptionTextBox = ""
        Exit Sub
    End If
     
    On Error GoTo NotAnInt
    index = findLeadMeasure(CInt(Me.leadIDComboBox), ActiveSheet.Name)
    
    'Populates the description textbox with the specified lead if found.
    If index <> -1 Then
        Dim rowNum As Integer
        rowNum = index + 14
        Me.descriptionTextBox.Value = ActiveSheet.range("M" & rowNum).Value
        
    Else
        MsgBox "Sorry, but Lead Measure does not exist. Please, select an existing Lead Measure."
        Call clearValues
    End If
    Exit Sub
    
NotAnInt:
    MsgBox "Please enter a valid integer."
    Call clearValues
    Exit Sub
End Sub
