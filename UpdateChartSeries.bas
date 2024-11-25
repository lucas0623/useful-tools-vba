Attribute VB_Name = "UpdateChartSeries"
Sub UpdateChartsToCurrentWorksheet()
    Dim activeWs As Worksheet
    Dim targetWs As Worksheet
    Dim chartObj As ChartObject
    Dim chrt As Chart
    Dim ser As Series
    Dim serFormulaParts() As String
    Dim sheetName As String
    Dim isChartSelected As Boolean
    Dim userChoice As VbMsgBoxResult

    ' Set the currently active worksheet
    Set activeWs = ActiveSheet

    ' Ask user to choose whether to update all charts or only the selected one
    userChoice = MsgBox("Do you want to update all charts?" & vbCrLf & "Click 'Yes' to update ALL chart in this sheet" & vbCrLf & "Click 'No' to update only the selected chart (only support selecting a single chart).", vbYesNo + vbQuestion, "Update Charts")

    ' Prompt user for the sheet name to reference
    sheetName = inputBox("Enter the name of the worksheet to reference in the chart, enter nothing will default to the active sheet:")

    ' Check if the user entered a sheet name
    If sheetName = "" Then
        ' Default to the active sheet name
        sheetName = activeWs.Name
        MsgBox "No sheet name entered. Defaulting to the active sheet: " & sheetName
    End If

    ' Check if the target sheet exists
    On Error Resume Next
    Set targetWs = Worksheets(sheetName)
    On Error GoTo 0

    If targetWs Is Nothing Then
        MsgBox "The specified worksheet does not exist. Please try again."
        Exit Sub
    End If

    ' Check if a chart is selected
    On Error Resume Next
    Set chrt = ActiveChart
    On Error GoTo 0

    isChartSelected = Not chrt Is Nothing

    ' Update charts based on user selection
    If userChoice = vbYes Then
        ' Update all charts in the active worksheet
        For Each chartObj In activeWs.ChartObjects
            Set chrt = chartObj.Chart
            If Not chrt Is Nothing Then
                Call UpdateChartSeries(chrt, sheetName)
            End If
        Next chartObj
    Else
        ' Update only the selected chart
        If isChartSelected Then
            Call UpdateChartSeries(chrt, sheetName)
        Else
            MsgBox "No chart is selected. Please select a chart or choose to update all charts."
        End If
    End If

    MsgBox "Operation Completed."
End Sub

Sub UpdateChartSeries(chrt As Chart, sheetName As String)
    Dim ser As Series
    Dim serFormulaParts() As String

    ' Loop through all series in the chart
    For Each ser In chrt.SeriesCollection
        serFormulaParts = Split(ser.formula, ",")

        ' Ensure the array has the expected number of elements
        If UBound(serFormulaParts) >= 2 Then
            ' Update the series formula to reference the specified sheet
            serFormulaParts(1) = ReplaceSheetNameInReference(serFormulaParts(1), sheetName)
            serFormulaParts(2) = ReplaceSheetNameInReference(serFormulaParts(2), sheetName)

            ' Attempt to reassemble the formula with error handling
            On Error Resume Next
            ser.formula = Join(serFormulaParts, ",")
            If Err.Number <> 0 Then
                Debug.Print "Error updating series formula for chart: " & chrt.Name & " with series: " & ser.Name
                Err.Clear
            End If
            On Error GoTo 0
        Else
            Debug.Print "Unexpected formula structure: " & ser.formula
        End If
    Next ser
End Sub

Function ReplaceSheetNameInReference(reference As String, sheetName As String) As String
    Dim refParts() As String
    refParts = Split(reference, "!")
    If UBound(refParts) > 0 Then
        ' Replace the sheet name part of the reference
        ReplaceSheetNameInReference = "'" & sheetName & "'!" & refParts(1)
    Else
        ' If reference does not have a sheet name, just prefix it
        ReplaceSheetNameInReference = "'" & sheetName & "'!" & reference
    End If
End Function


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
