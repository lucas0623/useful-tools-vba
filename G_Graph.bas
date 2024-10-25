Attribute VB_Name = "G_Graph"
Sub FormatGraph1()
    chartIndex = ActiveChart.Parent.Index
    chartName = ActiveChart.Parent.Name

    ActiveChart.ChartArea.Select
    
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesMinorMajor)
    ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMinorMajor)
    
    ActiveSheet.Shapes(chartName).height = 340.157480315
    ActiveSheet.Shapes(chartName).width = 548.2204724409
    
    With ActiveSheet.ChartObjects(chartName).Chart.ChartArea.Format.TextFrame2.TextRange.Font
        .NameComplexScript = "Arial"
        .NameFarEast = "Arial"
        .Name = "Arial"
    End With
    
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.size = 20
    
    On Error GoTo ErrorHandler1
        ActiveChart.Axes(xlValue).AxisTitle.Select
    On Error GoTo 0

    
    ActiveChart.Axes(xlValue).AxisTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.size = 16
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.size = 16
    
    ActiveChart.Axes(xlCategory).TickLabels.Font.size = 12
    ActiveChart.Axes(xlValue).TickLabels.Font.size = 12
    
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlCategory).Select
    Selection.MajorTickMark = xlCross
    Selection.MinorTickMark = xlInside
    ActiveChart.Axes(xlValue).Select
    Selection.MajorTickMark = xlCross
    Selection.MinorTickMark = xlInside
    
    ActiveChart.ChartTitle.Select
    With Selection.Format.TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Solid
    End With
    ActiveChart.Axes(xlValue).AxisTitle.Select
    With Selection.Format.TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Solid
    End With
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    With Selection.Format.TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Solid
    End With
    
    ActiveChart.Axes(xlValue).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
    End With
    ActiveChart.Axes(xlCategory).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
    End With

    ActiveChart.Axes(xlCategory).Select
    With Selection.TickLabels.Font
        .Bold = msoTrue
    End With
    
    ActiveChart.Axes(xlValue).Select
    With Selection.TickLabels.Font
        .Bold = msoTrue
    End With
    
    Exit Sub
    
ErrorHandler1:
    MsgBox ("Please enable graph axis title first")
    
End Sub
