Attribute VB_Name = "SwapPlotData"
'@Folder("VBAProject")
Sub SwapPlotDataXY()
    Dim cht As Chart
    Dim ser As Series
    Dim formula As String
    Dim dataSheet As String
    Dim xValues As String
    Dim yValues As String
    Dim seriesName As String
    Dim serFormulaPrefix As String
    Dim serFormulaSuffix As String
    
    ' Activate the chart manually by clicking on it in Excel
    Set cht = ActiveChart
    
    ' Check if the chart exists
    If Not cht Is Nothing Then
        ' Iterate through each series in the chart
        For Each ser In cht.SeriesCollection
            ' Get the formula of the series
            formula = ser.formula
            
            ' Extract the sheet name, XValues and YValues from the formula
            'dataSheet = Split(Split(formula, ",")(2), "!")(0)
            'xValues = Split(Split(formula, ",")(1), "!")(1)
            'yValues = Split(Split(formula, ",")(2), "!")(1)
            'seriesName = Split(Split(formula, ",")(0), "(")(1)
            serFormulaPrefix = Split(formula, ",")(0)
            xValues = Split(formula, ",")(1)
            yValues = Split(formula, ",")(2)
            serFormulaSuffix = Split(formula, ",")(3)
            ' Swap the XValues and YValues in the formula
            'ser.formula = "=SERIES(" & seriesName & ",'" & dataSheet & "'!" & yValues & ",'" & dataSheet & "'!" & xValues & ",1)"
            ser.formula = serFormulaPrefix & "," & yValues & "," & xValues & "," & serFormulaSuffix
        Next ser
        
        MsgBox "Plot data switched between X and Y axes for all series in the chart."
    Else
        MsgBox "Chart not found. Please ensure you have an active chart."
    End If
End Sub
