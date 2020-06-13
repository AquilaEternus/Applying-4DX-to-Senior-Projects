Attribute VB_Name = "Team4DXGenerator"
Sub createWIGTable()
Attribute createWIGTable.VB_ProcData.VB_Invoke_Func = " \n14"

    'Create title for table
    Range("A3").Select
    Range("A3").Value = "WIG"
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
    End With
    
    'Create WIG_Table using ListObject
    Range("A4").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$4:$F$4"), , xlYes).Name = _
        "WIG_Table"
        
    'Change the names of the column headers
    Cells(4, 1).Value = "ID"
    Cells(4, 2).Value = "Description"
    Cells(4, 3).Value = "Start Line"
    Cells(4, 4).Value = "End Line"
    Cells(4, 5).Value = "Dead Line"
    Cells(4, 6).Value = "Points"
    
    Range("A1").Select
End Sub

Sub createLeadMTable()

    'Create title for table
    Range("I3").Select
    Range("I3").Value = "Lead Measures"
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
    End With
    
    'Create WIG_Table using ListObject
    Range("A4").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$I$4:$M$4"), , xlYes).Name = _
        "LeadM_Table"
        
    'Change the names of the column headers
    Cells(4, 9).Value = "WIG_ID"
    Cells(4, 10).Value = "ID"
    Cells(4, 11).Value = "Description"
    Cells(4, 12).Value = "Points"
    Cells(4, 13).Value = "Status"
    
    Range("A1").Select
End Sub

Sub addWIGButton()
Attribute addWIGButton.VB_ProcData.VB_Invoke_Func = " \n14"

    'Generates the wig button
    ActiveSheet.Buttons.Add(373.5, 49.5, 54.75, 18.75).Select
    Selection.Characters.Text = "Add Wig"
    
    
    With Selection.Characters(Start:=1, Length:=8).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .ColorIndex = 1
    End With
    
    Selection.OnAction = "addWIGButton_OnClick"
End Sub

Sub addLeadMButton()
Attribute addLeadMButton.VB_ProcData.VB_Invoke_Func = " \n14"
    'Generates the Lead button
    ActiveSheet.Buttons.Add(737.25, 50.25, 110.25, 21).Select
    Selection.Characters.Text = "Add Lead Measure"
    With Selection.Characters(Start:=1, Length:=16).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .ColorIndex = 1
    End With
    Selection.OnAction = "addLeadMButton_OnClick"
End Sub

Sub addWIGButton_OnClick()
    AddWIGForm.Show
End Sub

Sub addLeadMButton_OnClick()
    AddLeadMForm.Show
End Sub
