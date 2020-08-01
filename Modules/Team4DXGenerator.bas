Attribute VB_Name = "Team4DXGenerator"
'This module is used to generate the tables and pie chart of a team's 4DX sheet

Sub createTeam(pName As String)

    range("A1").Value = "Scoreboard"
    range("A1").Font.Bold = True
    range("A1").Font.Size = 20
    range("A1").HorizontalAlignment = xlCenter
    
    range("A1:C1").MergeCells = True
    range("A2:B2").MergeCells = True
    range("A2").Value = "Name"
    range("C2").Value = "Pts"
    range("A2").Font.Size = 14
    range("C2").Font.Size = 14
    range("A2").Font.Bold = True
    range("C2").Font.Bold = True
    range("A3:B3").MergeCells = True
    range("A4:B4").MergeCells = True
    range("A5:B5").MergeCells = True
    range("A6:B6").MergeCells = True
    range("A7:B7").MergeCells = True
    range("C2:C7").HorizontalAlignment = xlRight
    range("A7").Value = "Team"
    range("C7").Value = 0
    range("A1:C7").BorderAround LineStyle:=xlContinuous, Weight:=xlThick
    range("A2:C7").Borders.LineStyle = xlContinuous
    range("A2:A7").RowHeight = 20
    range("A1:C7").Interior.ColorIndex = 35
    
    range("A8:V10").MergeCells = True
    range("A8").HorizontalAlignment = xlCenter
    range("A8").VerticalAlignment = xlCenter
    range("A8").Value = pName
    range("A8").Font.Bold = True
    range("A8").Font.Underline = True
    range("A8").Font.Size = 30
    range("A8").Interior.Color = vbGreen
    
    range("A11:Z100").Locked = False

End Sub

Sub createWIGtable()
    
    range("A13:E13").Select
    range("A13:E13").Merge
    range("A13:E13").Value = "WIG"
    range("A13:E13").Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
    End With
    
    range("F13:G13").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
    End With
    
    range("F13").Value = "Count: "
    range("G13").Value = 0
    
    range("A14").Select
    ActiveSheet.ListObjects.add(xlSrcRange, range("$A$14:$G$14"), , xlYes).Name = _
        "WIG_Table"
    Cells(14, 1).Value = "ID"
    Cells(14, 2).Value = "Description"
    Cells(14, 3).Value = "Start Line"
    Cells(14, 4).Value = "End Line"
    Cells(14, 5).Value = "Dead Line"
    Cells(14, 6).Value = "Acquired Points"
    Cells(14, 7).Value = "Total Points"
    
    range("A14:G14").Columns.AutoFit
    range("B15:B30").WrapText = True
    range("B15").ColumnWidth = 25
    
End Sub

Sub createLeadMeasureTable()

    range("K13:N13").Select
    range("K13:N13").Merge
    range("K13:N13").Value = "Lead Measures"
    range("K13:N13").Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
    End With
    
    range("O13:P13").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
    End With
    
    range("O13").Value = "Count: "
    range("P13").Value = 0
    
    range("L14").Select
    ActiveSheet.ListObjects.add(xlSrcRange, range("$K$14:$P$14"), , xlYes).Name = _
        "LeadM_Table"
    Cells(14, 11).Value = "WIG ID"
    Cells(14, 12).Value = "ID"
    Cells(14, 13).Value = "Description"
    Cells(14, 14).Value = "Points"
    Cells(14, 15).Value = "Assigned To"
    Cells(14, 16).Value = "Status"
    
    range("K14:P14").Columns.AutoFit
    range("M15:M39").WrapText = True
    range("M15").ColumnWidth = 25
    range("O15").ColumnWidth = 25
    range("P15").ColumnWidth = 15

End Sub

Sub createContributionChart(sheetName As String)
    Dim ws As Worksheet
    Dim chartName As String
    Set ws = Worksheets(sheetName)
    chartName = "scoreBreakdown"
    
    Application.CutCopyMode = False
    ActiveSheet.Shapes.AddChart2(256, xlPie).Select
    ActiveChart.SetSourceData Source:=range("'" & sheetName & "'" & "!$C$3:$C$6")
    ActiveChart.ApplyDataLabels Type:=xlDataLabelsShowLabelAndPercent
    ActiveChart.FullSeriesCollection(1).XValues = "'" & sheetName & "'" & "!$A$3:$B$6"
    
    'ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Scoreboard Breakdown"
    ActiveChart.Parent.Name = "scoreBreakdown"
    'ActiveChart.Parent.Select
    With ws.Shapes(chartName)
        .Left = range("F1").Left
        .Top = range("F1").Top
    End With
    ws.Shapes(chartName).ScaleWidth 1.1770833333, msoFalse, msoScaleFromTopLeft
    ws.Shapes(chartName).ScaleHeight 0.6631944444, msoFalse, msoScaleFromTopLeft
End Sub

Sub Hide_Tabs()
Attribute Hide_Tabs.VB_ProcData.VB_Invoke_Func = "H\n14"
    ThisWorkbook.Unprotect
    ActiveWindow.DisplayWorkbookTabs = False
    
    'Set every sheet except "Start" to VeryHidden
    Dim ws As Worksheet
    Worksheets("Start").Activate
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Start" Then
            If ws.Visible = xlSheetVeryHidden Then
                'Do nothing if sheet is already hidden.
            Else
                ws.Visible = xlSheetVeryHidden
            End If
        End If
    Next ws
    ThisWorkbook.Protect
    CommandBars.ExecuteMso "HideRibbon"
        
End Sub

Sub Unhide_Tabs()
Attribute Unhide_Tabs.VB_ProcData.VB_Invoke_Func = "U\n14"
    ThisWorkbook.Unprotect
    ActiveWindow.DisplayWorkbookTabs = True
    
    'Set every sheet except "Start" to Hidden
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Start" Then
            ws.Visible = True
        End If
    Next ws
    
    Sheets("InformationInput").Visible = True
    CommandBars.ExecuteMso "HideRibbon"
End Sub
