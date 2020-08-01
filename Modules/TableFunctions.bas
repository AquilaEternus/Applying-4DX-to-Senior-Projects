Attribute VB_Name = "TableFunctions"
'This module contains funcitons for adding and deleting rows from both
'WIG and Lead Measure tables, as well as functions for retrieving the
'index of a WIG or Lead on these tables.

'Adds an extra row to any ListObject table and fills up the new row
'with a varying amount of data.
Sub addDataRow(tableName As String, sheetName As String, values() As Variant)
    Dim table As ListObject
    Dim col As Integer
    Dim lastRow As range

    Set table = ActiveWorkbook.Worksheets(sheetName).ListObjects.Item(tableName)

    If table.ListRows.Count > 0 Then
        Set lastRow = table.ListRows(table.ListRows.Count).range
        For col = 1 To lastRow.Columns.Count
            If Trim(CStr(lastRow.Cells(1, col).Value)) <> "" Then
                table.ListRows.add
                Exit For
            End If
        Next col
    Else
        table.ListRows.add
    End If

    Set lastRow = table.ListRows(table.ListRows.Count).range
    
    For col = 1 To lastRow.Columns.Count
        If col <= UBound(values) + 1 Then lastRow.Cells(1, col) = values(col - 1)
    Next col
End Sub


'Searches for a given WIG in the ListObject "WIG_Table" based on its ID. Returns an
'index of it's position in the table, or a -1 if it was not found.
Function findWIG(wigID As Integer, sheetName As String) As Integer
    Dim range As ListObject
    Set range = ActiveWorkbook.Worksheets(sheetName).ListObjects("WIG_Table")
    
    Dim matchResult As Long
    On Error GoTo NoSuchWIG
    matchResult = WorksheetFunction.Match(wigID, range.ListColumns("ID").DataBodyRange, 0)
    'Dim rowNum As Long
    'rowNum = matchResult + 14
    findWIG = matchResult
    Exit Function
    
NoSuchWIG:
        findWIG = -1
    Exit Function
End Function

'Searches for a given Lead in the ListObject "LeadM_Table" based on its ID. Returns an
'index of it's position in the table, or a -1 if it was not found.
Function findLeadMeasure(wigID As Integer, sheetName As String) As Integer
    Dim range As ListObject
    Set range = ActiveWorkbook.Worksheets(sheetName).ListObjects("LeadM_Table")
    
    Dim matchResult As Long
    On Error GoTo NoSuchLeadMeasure
    matchResult = WorksheetFunction.Match(wigID, range.ListColumns("ID").DataBodyRange, 0)
    'Dim rowNum As Long
    'rowNum = matchResult + 14
    findLeadMeasure = matchResult
    Exit Function
    
NoSuchLeadMeasure:
        findLeadMeasure = -1
    Exit Function
End Function

'Deletes a WIG and all corresponding Lead Measures from both tables.
Function deleteWIG(id As Integer, sheetName As String) As Boolean
    Dim tbl As ListObject
    Dim matchResult As Long
    Dim lastRow As Long
    
    'Delete row of selected WIG
    Set tbl = ActiveWorkbook.Worksheets(sheetName).ListObjects("WIG_Table")
    'On Error GoTo NoSuchWIG
    matchResult = WorksheetFunction.Match(id, tbl.ListColumns("ID").DataBodyRange, 0)
    'MsgBox matchResult
    tbl.ListRows(matchResult).Delete
    
    'Delete corresponding Lead Measures
    Set tbl = ActiveWorkbook.Worksheets(sheetName).ListObjects("LeadM_Table")
    'MsgBox tbl.ListRows.Count
    'If tbl.ListRows.Count > 0 Then
        Dim leadRow As ListRow
        lastRow = tbl.ListRows.Count
        For index = 1 To lastRow
            'MsgBox leadRow
            'MsgBox index
            'MsgBox lastRow
            If index > lastRow Then
                deleteWIG = True
                Exit Function
            End If
             Set leadRow = tbl.ListRows(index)
            If leadRow.range(1, 1) = CStr(id) Then
                'Clean up points from scoreboard
                Call deleteScoreBoardPts(leadRow, sheetName)
                'Delete row and move back index by 1 to avoid missing a row
                leadRow.Delete
                index = index - 1
                lastRow = lastRow - 1
                                
            End If
        Next index
    'End If
    
    deleteWIG = True
    Exit Function
    
NoSuchWIG:
        deleteWIG = False
    Exit Function
End Function

'Helper function to deleteWIG function that deletes the points assigned to a certain
'individual in the Scoreboard when a lead is deleted.
Sub deleteScoreBoardPts(leadRow As ListRow, sheetName As String)
    Dim ws As Worksheet
    Set ws = Worksheets(sheetName)
    
    If leadRow.range(1, 5) = "Everyone" And leadRow.range(1, 6) = "Complete" Then
        For i = 3 To 6
            If ws.range("A" & i).Value <> "" Then
                ws.range("C" & i).Value = ws.range("C" & i).Value - CInt(leadRow.range(1, 4))
            End If
        Next i
        ws.range("C7").Value = ws.range("C7").Value - CInt(leadRow.range(1, 4))
    ElseIf leadRow.range(1, 6) = "Complete" Then
        For i = 3 To 6
            If leadRow.range(1, 5) = ws.range("A" & i).Value Then
                ws.range("C" & i).Value = ws.range("C" & i).Value - CInt(leadRow.range(1, 4))
            End If
        Next i
        ws.range("C7").Value = ws.range("C7").Value - CInt(leadRow.range(1, 4))
    End If
End Sub

'Deletes a Lead measure and subtracts the points belonging to it from the
'corresponding WIG's "Aquired Points" and "Total Points" column.
Function deleteLeadMeasure(id As Integer, sheetName As String) As Boolean
    Dim tbl2 As ListObject
    Dim matchResult As Long
    Dim rowNum As Long
    Dim wigID As String
    Dim status As String
    Dim points As Integer
    Dim tbl As ListObject
    'Delete row of selected Lead Measure
    Set tbl = ActiveWorkbook.Worksheets(sheetName).ListObjects("LeadM_Table")
    On Error GoTo NoSuchEntry
    matchResult = WorksheetFunction.Match(id, tbl.ListColumns("ID").DataBodyRange, 0)
    
    rowNum = matchResult + 14
    
    wigID = Worksheets(sheetName).Cells(rowNum, 11).Value
    points = Worksheets(sheetName).Cells(rowNum, 14).Value
    status = Worksheets(sheetName).Cells(rowNum, 16).Value
    
    Call deleteScoreBoardPts(tbl.ListRows(matchResult), sheetName)
    tbl.ListRows(matchResult).Delete
    
    'Subtract from corresponding WIG's Aquired Points and Total Points
    Set tbl = ActiveWorkbook.Worksheets(sheetName).ListObjects("WIG_Table")
    'On Error Resume Next
    matchResult = WorksheetFunction.Match(CInt(wigID), tbl.ListColumns("ID").DataBodyRange, 0)
    rowNum = matchResult + 14
    
    'Subtract lead measure points from WIG's "Total Points"
    'MsgBox Worksheets(sheetName).Cells(rowNum, 7).Value
    Worksheets(sheetName).Cells(rowNum, 7).Value = Worksheets(sheetName).Cells(rowNum, 7).Value - CInt(points)
    
    'Subtract lead measure points from WIG's "Aquired Points" column
    If CInt(Worksheets(sheetName).Cells(rowNum, 6).Value) <> 0 And status = "Complete" Then
        Worksheets(sheetName).Cells(rowNum, 6).Value = Worksheets(sheetName).Cells(rowNum, 6).Value - CInt(points)
    End If
    
    deleteLeadMeasure = True
    Exit Function
    
NoSuchEntry:
        deleteLeadMeasure = False
    Exit Function
End Function

'Sorts the LeadM_Table on an active sheet by WIG ID
Sub sortByWIGRef(sheetName As String)
    range("LeadM_Table[[#Headers],[WIG ID]]").Select
    ActiveSheet.ListObjects( _
        "LeadM_Table").Sort.SortFields.CLEAR
    ActiveSheet.ListObjects( _
        "LeadM_Table").Sort.SortFields.Add2 Key:=range( _
        "LeadM_Table[[#Headers],[WIG ID]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.ListObjects( _
        "LeadM_Table").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
