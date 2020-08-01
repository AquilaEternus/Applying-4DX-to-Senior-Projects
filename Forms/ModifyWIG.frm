VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModifyWIG 
   Caption         =   "Modify Team's Wildly Important Goals"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6015
   OleObjectBlob   =   "ModifyWIG.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModifyWIG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Removes a WIG from the table and all of its associated lead measures.
Private Sub removeWIGButton_Click()
    ActiveSheet.Unprotect
    If Me.wigIDComboBox.Value = "" Then
        MsgBox "Please select or add a WIG to modify.", vbCritical
        Exit Sub
    End If
    Call deleteWIG(Me.wigIDComboBox, ActiveSheet.Name)
    ActiveSheet.Protect
    Unload ModifyWIG
    
End Sub

'Updates the appropriate row where the WIG is found once the "Update" button is clicked
Private Sub updateWIGButton_Click()
    ActiveSheet.Unprotect
    If Me.wigIDComboBox.Value = "" Then
        MsgBox "Please select or add a WIG to modify.", vbCritical
        Exit Sub
    End If
    
    If IsDate(Me.startLineTextBox.Value) And IsDate(Me.endLineTextBox.Value) And IsDate(Me.deadLineTextBox.Value) Then
        Dim index As Integer
        Dim rowNum As Integer
    
        index = findWIG(CInt(Me.wigIDComboBox), ActiveSheet.Name)
        rowNum = index + 14
    
        ActiveSheet.range("B" & rowNum).Value = Me.descriptionTextBox.Value
        ActiveSheet.range("C" & rowNum).Value = Me.startLineTextBox.Value
        ActiveSheet.range("D" & rowNum).Value = Me.endLineTextBox.Value
        ActiveSheet.range("E" & rowNum).Value = Me.deadLineTextBox.Value
        Unload ModifyWIG
    Else
        MsgBox "Please make sure dates are written in mm/dd/yyyy format."
    End If
    ActiveSheet.Protect

End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets(ActiveSheet.Name)
    Me.wigIDComboBox.CLEAR
    For i = 15 To ws.range("A" & Application.Rows.Count).End(xlUp).Row
        Me.wigIDComboBox.AddItem ws.range("A" & i).Value
    Next i
    
    Call clearValues
End Sub

'Populates the form's textboxes with the information pertaining to that WIG when WIG gets selected
Private Sub wigIDComboBox_Change()
    Dim index As Long
    If Me.wigIDComboBox = "" Then
        Exit Sub
    End If
     
    On Error GoTo NotAnInt
    index = findWIG(CInt(Me.wigIDComboBox), ActiveSheet.Name)
    
    If index <> -1 Then
        Dim rowNum As Integer
        rowNum = index + 14
        Me.descriptionTextBox.Value = ActiveSheet.range("B" & rowNum).Value
        Me.startLineTextBox.Value = ActiveSheet.range("C" & rowNum).Value
        Me.endLineTextBox.Value = ActiveSheet.range("D" & rowNum).Value
        Me.deadLineTextBox.Value = ActiveSheet.range("E" & rowNum).Value
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

Private Sub clearValues()
    Me.wigIDComboBox = ""
    Me.descriptionTextBox.Value = ""
    Me.startLineTextBox.Value = ""
    Me.endLineTextBox.Value = ""
    Me.deadLineTextBox.Value = ""
End Sub
