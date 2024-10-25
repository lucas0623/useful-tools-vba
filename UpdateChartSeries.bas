Attribute VB_Name = "UpdateChartSeries"
'@Folder("VBAProject")
Sub UpdateChartsToCurrentWorksheet()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chrt As Chart
    Dim ser As Series
    Dim seriesInfo As Collection
    Dim serName As String
    Dim serXValues As String
    Dim serYValues As String
    Dim seriesData As Variant
    Dim serFormula As String
    Dim serFormulaPrefix As String, serFormulaSuffix As String
    Dim isChartFound As Boolean
    Set ws = ActiveSheet
    Set seriesInfo = New Collection
    
    ' Loop through all charts in the current worksheet
    For Each chartObj In ws.ChartObjects
        Set chrt = chartObj.Chart
        If chrt Is Nothing Then GoTo ExitSub
        isChartFound = True
        ' Save chart details
        Debug.Print "Chart Name: " & chartObj.Name
        
        ' Loop through all series in the chart
        For Each ser In chrt.SeriesCollection
            serName = ser.Name
            serFormulaPrefix = Split(ser.formula, ",")(0)
            serXValues = Split(ser.formula, ",")(1)
            serYValues = Split(ser.formula, ",")(2)
            serFormulaSuffix = Split(ser.formula, ",")(3)
            
            
            ' Print series details to Immediate Window
            Debug.Print "Series Name: " & serName
            Debug.Print "Series X Values: " & serXValues
            Debug.Print "Series Y Values: " & serYValues
            Debug.Print "Is referrencing current sheet: " & IsReferringToCurrentSheet(serXValues)
            Debug.Print "Converted to: " & ConvertToCurrentSheetReference(serXValues)
            
            ' Save series details into a collection as a 1-based array
            seriesData = Array(serName, serXValues, serYValues)
            seriesInfo.Add seriesData
            
            'Reassign the formula
            serFormula = serFormulaPrefix & "," _
                        & ConvertToCurrentSheetReference(serXValues) & "," _
                        & ConvertToCurrentSheetReference(serYValues) & "," _
                        & serFormulaSuffix
            ser.formula = serFormula
            
        Next ser
        
        ' Delete all series in the chart
'        For Each ser In chrt.SeriesCollection
'            ser.Formula = ""
'            ser.Formula = serFormula
'        Next ser
        
'        ' Re-add series to the chart
'        For Each seriesData In seriesInfo
'            With chrt.SeriesCollection.NewSeries
'                .Name = seriesData(0)
'                .Formula = seriesData(2)
'            End With
'        Next seriesData
        
        ' Clear seriesInfo after processing each chart
        Set seriesInfo = New Collection
    Next chartObj
    If Not isChartFound Then GoTo ExitSub
    MsgBox "Operation Completed."
    Exit Sub
ExitSub:
    MsgBox "No Chart is found in the current worksheet."
End Sub
Sub UpdateChartPlottingExtent()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chrt As Chart
    Dim ser As Series
    Dim minX As Double, maxX As Double
    Dim minY As Double, maxY As Double
    Dim xRange As Double, yRange As Double
    Dim xMajorUnit As Double, yMajorUnit As Double
    Dim i As Integer
    Dim isChartFound As Boolean
    Set ws = ActiveSheet
    
    ' Loop through all charts in the current worksheet
    For Each chartObj In ws.ChartObjects
        Set chrt = chartObj.Chart
        If chrt Is Nothing Then GoTo ExitSub
        isChartFound = True
        ' Initialize min and max values
        minX = chrt.Axes(xlCategory).MinimumScale
        maxX = chrt.Axes(xlCategory).MaximumScale
        minY = chrt.Axes(xlValue).MinimumScale
        maxY = chrt.Axes(xlValue).MaximumScale
        
        ' Loop through all series in the chart
        For Each ser In chrt.SeriesCollection
            ' Loop through each point in the series
            For i = 1 To ser.Points.count
                ' Get x and y values
                Dim xValue As Double
                Dim yValue As Double
                xValue = ser.xValues(i)
                yValue = ser.Values(i)
                
                ' Update min and max bounds for x
                If xValue < minX Then minX = xValue
                If xValue > maxX Then maxX = xValue
                
                ' Update min and max bounds for y
                If yValue < minY Then minY = yValue
                If yValue > maxY Then maxY = yValue
            Next i
        Next ser
        
        ' Calculate ranges
        xRange = maxX - minX
        yRange = maxY - minY
        
        ' Calculate major unit spacings
        xMajorUnit = Round((xRange / 8) / 25) * 25
        yMajorUnit = Round((yRange / 8) / 25) * 25
        
        ' Update chart axes
        With chrt.Axes(xlCategory)
            '.MinimumScale = minX
            '.MaximumScale = maxX
            .MajorUnit = xMajorUnit
        End With
        
        With chrt.Axes(xlValue)
            '.MinimumScale = minY
            '.MaximumScale = maxY
            .MajorUnit = yMajorUnit
        End With
        
    Next chartObj
    If Not isChartFound Then GoTo ExitSub
    MsgBox "Chart plotting extents updated."
    Exit Sub
ExitSub:
    MsgBox "No Chart is found in the current worksheet."
End Sub
Function IsReferringToCurrentSheet(formulaString As String) As Boolean
    Dim currentSheetName As String
    Dim extractedSheetName As String
    Dim exclamPos As Integer
    Dim singleQuotePos As Integer
    Dim endQuotePos As Integer
    
    ' Get the current sheet name
    currentSheetName = ActiveSheet.Name
    
    ' Find the position of the exclamation mark that separates the sheet name and cell address
    exclamPos = InStr(1, formulaString, "!")
    
    If exclamPos = 0 Then
        ' If there is no exclamation mark, assume it's a reference within the current sheet
        IsReferringToCurrentSheet = True
        Exit Function
    End If
    
    ' Check if the sheet name is wrapped in single quotes
    singleQuotePos = InStr(1, formulaString, "'")
    
    If singleQuotePos = 1 Then
        ' If the sheet name is wrapped in single quotes, find the closing quote
        endQuotePos = InStr(2, formulaString, "'")
        ' Extract the sheet name
        extractedSheetName = Mid(formulaString, 2, endQuotePos - 2)
    Else
        ' Extract the sheet name directly
        extractedSheetName = left(formulaString, exclamPos - 1)
    End If
    
    ' Compare the extracted sheet name with the current sheet name
    If extractedSheetName = currentSheetName Then
        IsReferringToCurrentSheet = True
    Else
        IsReferringToCurrentSheet = False
    End If
End Function

Function ConvertToCurrentSheetReference(formulaString As String) As String
    Dim currentSheetName As String
    Dim exclamPos As Integer
    Dim singleQuotePos As Integer
    Dim endQuotePos As Integer
    Dim cellAddress As String
    Dim newFormulaString As String
    
    ' Get the current sheet name
    currentSheetName = ActiveSheet.Name
    
    ' Find the position of the exclamation mark that separates the sheet name and cell address
    exclamPos = InStr(1, formulaString, "!")
    
    If exclamPos = 0 Then
        ' If there is no exclamation mark, the formula is already referring to the current sheet
        cellAddress = formulaString
    Else
        ' Check if the sheet name is wrapped in single quotes
        singleQuotePos = InStr(1, formulaString, "'")
        
        If singleQuotePos = 1 Then
            ' If the sheet name is wrapped in single quotes, find the closing quote
            endQuotePos = InStr(2, formulaString, "'")
            ' Extract the cell address part after the exclamation mark
            cellAddress = Mid(formulaString, exclamPos + 1)
        Else
            ' Extract the cell address part directly after the exclamation mark
            cellAddress = Mid(formulaString, exclamPos + 1)
        End If
    End If
    
    ' Construct the new formula string with the current sheet name
    If InStr(currentSheetName, " ") > 0 Then
        ' If the current sheet name contains spaces, wrap it in single quotes
        newFormulaString = "'" & currentSheetName & "'!" & cellAddress
    Else
        newFormulaString = currentSheetName & "!" & cellAddress
    End If
    
    ' Return the new formula string
    ConvertToCurrentSheetReference = newFormulaString
End Function
