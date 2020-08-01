VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdminLogin 
   Caption         =   "Admin User Login"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3030
   OleObjectBlob   =   "AdminLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'On clicking Login as admin, check if user is really an admin. If admin account is found,
'make "InformationInput" the active sheet and call Unhide_Tabs to unprotect workbook and
'make everything visible to admin.
Private Sub LoginButton_Click()
    Dim infoSheet As Worksheet
    Dim range As ListObject
    
    Set infoSheet = ThisWorkbook.Worksheets("InformationInput")
    Set range = infoSheet.ListObjects("Admin_Table")
    
    Dim matchResult As Long
    On Error GoTo NoSuchEntryExists
    matchResult = WorksheetFunction.Match(Me.UsernameTextBox, range.ListColumns("Admin").DataBodyRange, 0)
    
    Dim adminRow As ListRow
    Set adminRow = range.ListRows(matchResult)
    If adminRow.range(1, 1) = Me.UsernameTextBox Then
        
        If CStr(adminRow.range(1, 2)) = CStr(Me.PasswordTextBox) Then
            Worksheets("InformationInput").Activate
            Call Unhide_Tabs
            Unload AdminLogin
            MsgBox "Successfully signed in!"
        Else
            MsgBox "Your username or password wrong."
            Unload AdminLogin
        End If
    Else
        MsgBox "Your username or password wrong."
        Unload AdminLogin
    End If
    
    Exit Sub
    
NoSuchEntryExists:
    MsgBox "Username or password is incorrect."
    Me.UsernameTextBox = ""
    Me.PasswordTextBox = ""
    Exit Sub
End Sub




