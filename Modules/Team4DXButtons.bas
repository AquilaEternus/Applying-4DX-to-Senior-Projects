Attribute VB_Name = "Team4DXButtons"
'This module is used to generate the buttons of a team's 4DX sheet

Sub add_WIG()
Attribute add_WIG.VB_ProcData.VB_Invoke_Func = "A\n14"

    ActiveSheet.Unprotect
    ActiveSheet.Buttons.add(515.25, 208, 87.75, 32.25).Select
    Selection.Characters.Text = "ADD WIG"
    Selection.OnAction = "add_wig_click"
        
End Sub


Sub add_Lead()

    ActiveSheet.Unprotect
    ActiveSheet.Buttons.add(1135.25, 208, 87.75, 32.25).Select
    Selection.Characters.Text = "ADD LEAD"
    Selection.OnAction = "add_lead_click"
    
End Sub

Sub edit_WIG()

    ActiveSheet.Unprotect
    ActiveSheet.Buttons.add(515.25, 250, 87.75, 32.25).Select
    Selection.Characters.Text = "EDIT"
    Selection.OnAction = "edit_wig_click"

End Sub

Sub edit_Lead()

    ActiveSheet.Unprotect
    ActiveSheet.Buttons.add(1135.25, 250, 87.75, 32.25).Select
    Selection.Characters.Text = "EDIT"
    Selection.OnAction = "edit_lead_click"

End Sub

Sub add_Logout()

    ActiveSheet.Unprotect
    ActiveSheet.Buttons.add(1135.25, 10, 87.75, 32.25).Select
    Selection.Characters.Text = "LOGOUT"
    Selection.OnAction = "add_logout_click"

End Sub

Sub add_wig_click()
    AddWIG.Show
End Sub

Sub add_lead_click()
    AddLeadMeasure.Show
End Sub

Sub edit_wig_click()
    ModifyWIG.Show
End Sub

Sub edit_lead_click()
    ModifyLead.Show
End Sub

Sub add_logout_click()
    Call Hide_Tabs
End Sub



