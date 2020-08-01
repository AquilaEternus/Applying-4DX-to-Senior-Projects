VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddWIG 
   Caption         =   "Add Wildly Important Goals"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7515
   OleObjectBlob   =   "AddWIG.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddWIG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addWigButton_Click()

    ActiveSheet.Unprotect
    Dim data(6)
    data(0) = range("G13").Value
    data(1) = Me.descriptionTextBox.Value
    data(2) = Me.startLineTextBox.Value
    data(3) = Me.endLineTextBox.Value
    data(4) = Me.deadLineTextBox.Value
    data(5) = 0
    data(6) = 0
    'Pass the form info stored in the variant data to addDataRow
    'and add it to the WIG table
    addDataRow "WIG_Table", ActiveSheet.Name, data
    
    'Increment WIG count at top right to simulate an auto-incremented ID
    range("G13").Value = range("G13").Value + 1
    ActiveSheet.Protect
    Unload AddWIG

End Sub



