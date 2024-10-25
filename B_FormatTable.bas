Attribute VB_Name = "B_FormatTable"
Sub AddBorder_Hori()

    Dim myRow As Range
    Dim num_lastRow As Long
    Dim str_lastRow As String
    Dim num_cuurentRow As Long
    Dim str_currentRow As String
    Dim colNum As Long, colRef As Range
    Dim count As Long
    Dim inputBox As Variant
    Dim lineType As Long
    
    lineType = Application.inputBox( _
                                        prompt:="PLEASE INPUT THE LINE TYPE NUMBER. " & vbNewLine & vbNewLine _
                                        & "NOTE" & vbNewLine _
                                        & "Input 1 for thin line  " & vbNewLine _
                                        & "Input 2 for thick line" & vbNewLine _
                                        & "Input 3 for double line", _
                                        default:=1, _
                                        Type:=1)
    If lineType = 0 Then End
    Set colRef = AddInputBox_rng( _
                                        prompt:="PLEASE SELECT THE COLUMN. " & vbNewLine & vbNewLine _
                                        & "NOTE" & vbNewLine _
                                        & "1. The column must be selected. " & vbNewLine _
                                        & "2. The bottom border will be added whenever the data in the current row of the SELECTED COLUMN is different from that in the last row")
    
    count = 0

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    colNum = colRef.Column
    For Each myRow In Selection.Rows
        
        If count > 0 Then
            
            num_lastRow = myRow.row - 1
            num_currentRow = myRow.row
            str_lastRow = Cells(num_lastRow, colNum)
            str_currentRow = Cells(num_currentRow, colNum)
            
            
            If str_currentRow <> str_lastRow Then
                
                If lineType = 1 Then
                    With Rows(num_lastRow).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = 1
                    End With
                ElseIf lineType = 2 Then
                    With Rows(num_lastRow).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                        .ColorIndex = 1
                    End With
                ElseIf lineType = 3 Then
                    With Rows(num_lastRow).Borders(xlEdgeBottom)
                        .LineStyle = xlDouble
                        .Weight = xlThick
                        .ColorIndex = 1
                    End With
                Else:
                    MsgBox ("Please Input Valid Line Type Number")
                    End
                End If
            End If
        
       End If
        
        count = count + 1
    Next
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub

Sub BoldMaxOfSameGroup()

    Dim myRow As Range
    Dim lRow As Long
    'Dim str_lastRow As String
    Dim cRow As Long
    Dim lRowOfSameGrp As Long
    'Dim str_currentRow As String
    Dim colNum As Long
    Dim colNum2 As Long 'the column number to determine the maximum value
    Dim count As Long
    Dim inputBox As Variant
    Dim i As Long, j As Long
    
    Dim maxValue As Double 'a temp variable to compare value
    Dim cValue As Double
    Dim row_maxValue As Long
    Dim colRef As Range, colRef2 As Range
    'Dim lineType As Integer
    
    
    Set colRef = AddInputBox_rng(prompt:="PLEASE SELECT THE COLUMN FOR DETERMINING GROUP. ")
    Set colRef2 = AddInputBox_rng(prompt:="PLEASE SELECT THE COLUMN FOR IDENTIFYING MAXIMUM VALUE. ")
    colNum = colRef.Column
    colNum2 = colRef2.Column
    count = 0
    
    cRow = Selection.row
    lRow = Selection.row + Selection.count - 1
    
    For i = cRow To lRow
    
        lRowOfSameGrp = FindLastRowOfSameGrp(i, colNum)
        
        cValue = Cells(i, colNum2)
        maxValue = cValue
        row_maxValue = i
        
        For j = i To lRowOfSameGrp
            cValue = Cells(j, colNum2)
            If cValue > maxValue Then
                maxValue = cValue
                row_maxValue = j
            End If
        Next j
        Rows(row_maxValue).Font.Bold = True

        i = j
    Next i

End Sub

Sub ClearBorder_Hori()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Dim myRow As Range
    
    For Each myRow In Selection.Rows
    
        With myRow.Borders(xlEdgeBottom)
            .LineStyle = xlLineStyleNone
        End With
    Next
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub

Sub ClearBorder_Right()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Dim myCol As Range
    
    For Each myCol In Selection.Columns
    
        With myCol.Borders(xlEdgeRight)
            .LineStyle = xlLineStyleNone
        End With
    Next
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub
Sub GetMaxOfSameGroup()

    Dim myRow As Range
    Dim lRow As Long
    'Dim str_lastRow As String
    Dim cRow As Long
    Dim lRowOfSameGrp As Long
    'Dim str_currentRow As String
    Dim colNum As Long
    Dim colNum2 As Long 'the column number to determine the maximum value
    Dim count As Long
    Dim inputBox As Variant
    Dim i As Long, j As Long
    
    Dim maxValue As Double 'a temp variable to compare value
    Dim cValue As Double
    Dim row_maxValue As Long
    Dim colRef As Range, colRef2 As Range, outputRng As Range
    Dim oWS As Worksheet
    'Dim lineType As Integer
    Set oWS = ActiveSheet
    
    Set colRef = AddInputBox_rng(prompt:="PLEASE SELECT THE COLUMN FOR DETERMINING GROUP. ")
    Set colRef2 = AddInputBox_rng(prompt:="PLEASE SELECT THE COLUMN FOR IDENTIFYING MAXIMUM VALUE. ")
    Set outputRng = AddInputBox_rng(prompt:="PLEASE SELECT THE TARGET OUTPUT LOCATION. ")
    colNum = colRef.Column
    colNum2 = colRef2.Column
    count = 0
    
    cRow = Selection.row
    lRow = Selection.row + Selection.count - 1
    
    For i = cRow To lRow
    
        lRowOfSameGrp = FindLastRowOfSameGrp(i, colNum)
        
        cValue = Cells(i, colNum2)
        maxValue = cValue
        row_maxValue = i
        
        For j = i To lRowOfSameGrp
            cValue = Cells(j, colNum2)
            If cValue > maxValue Then
                maxValue = cValue
                row_maxValue = j
            End If
        Next j
        outputRng.Worksheet.Cells(outputRng.row + count, outputRng.Column) = oWS.Cells(row_maxValue, colNum)
        outputRng.Worksheet.Cells(outputRng.row + count, outputRng.Column + 1) = oWS.Cells(row_maxValue, colNum2)
        count = count + 1
        i = j
    Next i

End Sub

Sub MergeCellsInSameGroup()
    Dim fCol As Long, lCol As Long
    Dim fRow As Long, lRow As Long
    Dim i As Long, j As Long
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    fCol = Selection.Column
    lCol = Selection.Column + Selection.Columns.count - 1
    'colNum2 = colRef2.Column
    count = 0
    
    fRow = Selection.row
    lRow = Selection.row + Selection.Rows.count - 1
    
    For i = fRow To lRow
        lRowOfSameGrp = LastRowOfSameGrp(i, fCol, lRow)
        For j = fCol To lCol
            Range(Cells(i, j), Cells(lRowOfSameGrp, j)).Merge
        Next j
        i = lRowOfSameGrp
    Next i
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True

End Sub
Private Function LastRowOfSameGrp(ByVal fRow As Long, ByVal rCol As Long, Optional lRow As Long, Optional ws As Worksheet) As Long
    Dim cRow As Long
    
    If lRow = 0 Then lRow = 1048576
    If ws Is Nothing Then Set ws = ActiveSheet
    If lRow < fRow Then
        LastRowOfSameGrp = 0
        Exit Function
    End If
    cRow = fRow
'find lRow_SchSameGrp
    With ws
        Do Until (Not .Cells(fRow, rCol).text = .Cells(cRow, rCol).text And Not .Cells(cRow, rCol) = vbNullString) Or cRow > lRow
            cRow = cRow + 1
        Loop
        LastRowOfSameGrp = cRow - 1
    End With
End Function

Private Function AddInputBox_rng(prompt As String, Optional title As String, Optional default As Range) As Range
    If default Is Nothing Then
        Set default = Range("D1")
    End If
    On Error GoTo ErrHandler:
        Set AddInputBox_rng = Application.inputBox(prompt:=prompt, title:=title, default:=default.Address, Type:=8)
    On Error GoTo 0
    Exit Function
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    End
End Function

Private Function FindLastRowOfSameGrp(ByVal rRow As Long, ByVal rCol As Long, Optional ws As Worksheet) As Long
    Dim cRow As Long
    Dim lRow As Long
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    cRow = rRow
    lRow = rRow + 1
'find lRow_SchSameGrp
    With ws
        Do Until .Cells(cRow, rCol).text <> .Cells(lRow, rCol).text
            cRow = lRow
            lRow = lRow + 1
        Loop
        FindLastRowOfSameGrp = cRow
    End With
End Function

