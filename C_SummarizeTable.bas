Attribute VB_Name = "C_SummarizeTable"
Sub SummarizeByMaxValInSelection()

    Dim myRow As Range
    Dim lRow As Long
    'Dim str_lastRow As String
    Dim cRow As Long
    Dim lRowOfSameGrp As Long
    'Dim str_currentRow As String
    Dim colNum As Long, colNum2 As Long 'the column number to determine the maximum value
    Dim count As Long

    Dim i As Long, j As Long
    
    Dim maxValue As Double 'a temp variable to compare value
    Dim cValue As Double
    Dim row_maxValue As Long
    Dim colRef As Range, colRef2 As Range, oWS As Worksheet
    Dim tarRng As Range, tarWS As Worksheet
    Dim arrL_Summarized As Object
    'Dim lineType As Integer
    
    Set arrL_Summarized = CreateObject("System.Collections.ArrayList")
    Set oWS = ActiveSheet
    Set colRef = AddInputBox_rng(prompt:="PLEASE SELECT THE COLUMN FOR DETERMINING GROUP. ")
    Set colRef2 = AddInputBox_rng(prompt:="PLEASE SELECT THE COLUMN FOR IDENTIFYING MAXIMUM VALUE. ")
    
    Set tarRng = AddInputBox_rng(prompt:="PLEASE SELECT THE TARGET OUTPUT LOCATION. ")
    Set tarWS = tarRng.Worksheet
    
    
    colNum = colRef.Column
    colNum2 = colRef2.Column
    count = 0
    
    cRow = Selection.row
    lRow = Selection.row + Selection.count - 1
    
    For i = cRow To lRow
    
        lRowOfSameGrp = FindLastRowOfSameGrp(i, colNum)
        
        cValue = oWS.Cells(i, colNum2)
        maxValue = cValue
        
        For j = i To lRowOfSameGrp
            cValue = oWS.Cells(j, colNum2)
            If cValue > maxValue Then
                maxValue = cValue
                row_maxValue = j
            End If
        Next j
        arrL_Summarized.Add Array(oWS.Cells(i, colNum), maxValue)
        'Rows(row_maxValue).Font.Bold = True

        i = j
    Next i

End Sub
