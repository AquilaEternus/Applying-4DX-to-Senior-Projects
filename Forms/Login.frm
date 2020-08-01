VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Login 
   Caption         =   "UserForm1"
   ClientHeight    =   8430.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15060
   OleObjectBlob   =   "Login.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Clears form data if user decides to click on the Clear button-
Private Sub CLEAR_Click()
    
    projectName.Value = ""
    username.Value = ""
    pw.Value = ""
    
End Sub

'On clicking "Ok". Checks to set if student is registered with given project. If so, then it
'checks to see if the given password matches the password on the same row where the student
'username and assigned projects is located.
Private Sub OK_Click()
    Dim i As Integer
    Dim ws As Worksheet
    Dim found As Boolean
    found = False
    Set ws = ThisWorkbook.Sheets("InformationInput")
    For i = 2 To ws.range("C" & Application.Rows.Count).End(xlUp).Row
        If ws.range("D" & i).Value = username.Value And ws.range("C" & i).Value = projectName.Value Then
            found = True
            If CStr(ws.range("E" & i).Value) = CStr(pw.Value) Then
                Unload Login
                ThisWorkbook.Unprotect
                Worksheets(projectName.Value).Visible = True
                ThisWorkbook.Protect
                Worksheets(projectName.Value).Activate
                MsgBox "Welcome to 4DX!"
            Else
                MsgBox "Wrong Password, please try again."
                Exit Sub
            End If
        End If
    Next i
    
    If Not found Then
        MsgBox "Username not found, please register before logging in"
        Exit Sub
    End If
    

End Sub

Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    Dim i As Integer
    Set ws = ThisWorkbook.Sheets("InformationInput")
    
    'Populates projectName combobox with the list of projects from "InformationInput"
    For i = 2 To ws.range("B" & Application.Rows.Count).End(xlUp).Row
        If ws.range("B" & i).Value <> "" Then
            projectName.AddItem ws.range("B" & i).Value
        End If
    Next i

End Sub
