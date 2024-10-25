VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWorksheetsInteraction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************
'This class contains some useful functions in macro development.
'Arthor: Lucas LEUNG
'Update Log
'25 Oct 2022 - Add Functions (FindFirstNonEmptyCell, CountNonEmptyCell, GetNonEmptyCell)
'31 Jul 2022 - Add Functions to change varaint array to double/string array

'*******************************************************

Function getLocVar(ws As Worksheet, label As String, isRow As Boolean, Optional errText As String = "correct data", Optional isMustMatch = True) As Long
    
    If isMustMatch Then
        On Error GoTo Err_noMatch
    Else
        On Error Resume Next
    End If
    
    With ws
        If isRow Then
            getLocVar = Application.Match(label, .Rows(1), 0)
            
        Else
            getLocVar = Application.Match(label, .Columns(1), 0)
        End If
    End With
    
    On Error GoTo 0

    Exit Function
    
Err_noMatch:
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    
    If isRow Then
        MsgBox "Err: Cannot locate the column containing the " & errText & "!!! :(" & vbNewLine & _
                "Please input '" & label & "' at the column containing the" & errText & " in Row 1 of the '" & ws.Name & "' worksheet!!!!"
    Else:
        MsgBox "Err: Cannot locate the row containing the " & errText & "!!! :(" & vbNewLine & _
                "Please input '" & label & "' at the row containing the" & errText & " in Column 1 of the '" & ws.Name & "' worksheet!!!!"
    End If
    
    End
End Function

Function getLocVar_containText(ws As Worksheet, label As String, isSearchRow As Boolean) As Long()
    Dim lCol As Long, searchRng As Range, lRow As Long
    Dim rng As Range, i As Long
    Dim arrL_matched As Object, arr_matched() As Long
    '1. set the range by finding the first and last row
    Set arrL_matched = CreateObject("System.Collections.ArrayList")
    If isSearchRow Then
        lCol = FindLastColOfWS(1, ws)
        Set searchRng = Range(ws.Cells(1, 1), ws.Cells(1, lCol))
    Else
        lRow = FindLastRowOfWS(1, ws)
        Set searchRng = Range(ws.Cells(1, 1), ws.Cells(lRow, 1))
    End If
    
    For Each rng In searchRng
        If InStr(1, rng.text, label) > 0 Then
            If isSearchRow Then
                arrL_matched.Add rng.Column
            Else
                arrL_matched.Add rng.row
            End If
        End If
    Next

    ReDim arr_matched(arrL_matched.count - 1)
    For i = 0 To arrL_matched.count - 1
        arr_matched(i) = arrL_matched(i)
    Next i
    getLocVar_containText = arr_matched
End Function

Function setWorksheet(wsName As String, Optional errText As String = "") As Worksheet
    
    On Error GoTo Err_cantFindWS
    
    Set setWorksheet = Worksheets(wsName)
    
    On Error GoTo 0
    Exit Function
    
Err_cantFindWS:
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    MsgBox "Err: Cannot find the Worksheet '" & wsName & "' " & errText & " :(" '& vbNewLine & _
           '"The worksheet shall be name as '" & wsName & "' and follow the required format. "
    End
End Function

Function FindFirstNonEmptyCell(rng As Range) As Range
    Set FindFirstNonEmptyCell = rng.Find(what:="*", LookIn:=xlValues)
End Function

Function FindLastNonEmptyCell(rng As Range) As Range
    Set FindLastNonEmptyCell = rng.Find(what:="*", LookIn:=xlValues, SearchDirection:=xlPrevious)
End Function

Function CountNonEmptyCell(rng As Range) As Long
    Dim num_const As Long, num_formulas As Long
    On Error Resume Next
    num_formulas = rng.Cells.SpecialCells(xlCellTypeFormulas).count
    num_const = rng.Cells.SpecialCells(xlCellTypeConstants).count
    On Error GoTo 0
    CountNonEmptyCell = num_const + num_formulas
End Function

Function FindNonEmptyCell(rng As Range) As Range()
    Dim rng_nonEmpty() As Range, cell As Range
    Dim num_nonEmpty As Long, count As Long, i As Long
    Dim fRng As Range, lRng As Range, searchRng As Range
    
    Set fRng = FindFirstNonEmptyCell(rng)
    Set lRng = FindLastNonEmptyCell(rng)
    Set searchRng = Range(fRng, lRng)
    num_nonEmpty = CountNonEmptyCell(rng)
    ReDim rng_nonEmpty(num_nonEmpty - 1)
    
    For Each cell In searchRng
        If Not cell.text = "" Then
            Set rng_nonEmpty(count) = cell
            count = count + 1
        End If
        If count = num_nonEmpty Then Exit For
    Next
    FindNonEmptyCell = rng_nonEmpty
End Function

Function FindLastRowOfWS(Optional rCol As Long = 1, Optional ws As Worksheet) As Long
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    FindLastRowOfWS = ws.Cells(ws.Rows.count, rCol).End(xlUp).row
    'FindLastRowOfWS = .Range("A" & .Rows.count).End(xlUp).row
    '
End Function
Function FindLastRow(ByVal ref_Row As Long, ByVal refCol As Long, Optional ws As Worksheet) As Long 'return the last row of a tables
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    With ws
        If .Cells(ref_Row, refCol).text = "" Then
            FindLastRow = ref_Row
        'MsgBox "NO DATA IN THE WS"
        ElseIf .Cells(ref_Row + 1, refCol).text <> "" Then
            FindLastRow = .Cells(ref_Row, refCol).End(xlDown).row
        Else
            FindLastRow = ref_Row
        End If
    End With

End Function

Function FindLastRowOfSameGrp(ByVal rRow As Long, ByVal rCol As Long, Optional ws As Worksheet) As Long
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

Function FindLastCol(ByVal ref_Row As Long, ByVal refCol As Long, Optional ws As Worksheet) As Long
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    With ws
        If .Cells(ref_Row, refCol).text = "" Then
            FindLastCol = refCol
'            MsgBox "NO DATA IN THE WS"
        ElseIf .Cells(ref_Row, refCol + 1) <> "" Then
            FindLastCol = .Cells(ref_Row, refCol).End(xlToRight).Column

        Else
            FindLastCol = refCol
        End If
    End With
End Function

Function FindLastColOfWS(Optional rRow As Long = 1, Optional ws As Worksheet) As Long
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    FindLastColOfWS = ws.Cells(rRow, ws.Columns.count).End(xlToLeft).Column
    'FindLastRowOfWS = .Range("A" & .Rows.count).End(xlUp).row
    '
End Function

Function SetArrayOneDim(fRow As Long, fCol As Long, Optional isByRow As Boolean = True, Optional arrSize As Long = 0, Optional ws As Worksheet) As Variant
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    Dim i As Long
    Dim lRow As Long, lCol As Long
    
    
    With ws
        If arrSize = 0 Then
            If isByRow Then
                lRow = FindLastRow(fRow, fCol, ws)
                arrSize = lRow - fRow + 1
            Else
                lCol = FindLastCol(fRow, fCol, ws)
                arrSize = lCol - fCol + 1
            End If
        End If
                
        If isByRow Then
            SetArrayOneDim = .Range(.Cells(fRow, fCol), .Cells(fRow + arrSize - 1, fCol))
        Else
            SetArrayOneDim = .Range(.Cells(fRow, fCol), .Cells(fRow, fCol + arrSize - 1))
        End If
        
    End With
    
End Function




Function SetArrayTwoDim(fRow As Long, fCol As Long, Optional isGrpByRow As Boolean = True, Optional arrSizeY As Long = 0, Optional arrSizeX As Long = 0, Optional ws As Worksheet) As Variant
    
    Dim rowNum As Long, colNum As Long
    Dim lRow As Long, lCol As Long
    Dim size1 As Long, size2 As Long
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    lRow = FindLastRow(fRow, fCol, ws)
    lCol = FindLastCol(fRow, fCol, ws)
    
    If arrSizeY = 0 Then
        rowNum = lRow - fRow + 1
    Else
        rowNum = arrSizeY
    End If
    
    If arrSizeX = 0 Then
        colNum = lCol - fCol + 1
    Else
        colNum = arrSizeX
    End If
    
    SetArrayTwoDim = ws.Range(ws.Cells(fRow, fCol), ws.Cells(fRow + rowNum - 1, fCol + colNum - 1)).Value
    
    If Not isGrpByRow Then
        SetArrayTwoDim = TransposeArray(SetArrayTwoDim)
    End If
End Function

Function TransposeArray(MyArray As Variant) As Variant
Dim x As Long
Dim Y As Long
Dim Xupper As Long
Dim Yupper As Long
Dim Xlower As Long
Dim Ylower As Long

Dim tempArray As Variant
    Xupper = UBound(MyArray, 2)
    Yupper = UBound(MyArray, 1)
    Xlower = LBound(MyArray, 2)
    Ylower = LBound(MyArray, 1)
    
    ReDim tempArray(Xlower To Xupper, Ylower To Yupper)
    For x = Xlower To Xupper
        For Y = Ylower To Yupper
            tempArray(x, Y) = MyArray(Y, x)
        Next Y
    Next x
    TransposeArray = tempArray
End Function

Sub WriteArrToRngTwoDim(myArr As Variant, fRow As Long, fCol As Long, Optional ws As Worksheet)
    Dim arrSizeX As Long, arrSizeY As Long
    Dim myRng As Range
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    arrSizeY = UBound(myArr, 1) - LBound(myArr, 1) + 1
    arrSizeX = UBound(myArr, 2) - LBound(myArr, 2) + 1
    Set myRng = ws.Range(ws.Cells(fRow, fCol), ws.Cells(fRow + arrSizeY - 1, fCol + arrSizeX - 1))
    myRng.Value = myArr
End Sub

Sub WriteArrToColumn(myArr As Variant, fRow As Long, fCol As Long, Optional ws As Worksheet, Optional size As Long)
    Dim arrSizeY As Long
    Dim myRng As Range
    Dim arr_2D As Variant
    Dim i As Long
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    If size > 0 Then
        arrSizeY = size
    Else
        arrSizeY = UBound(myArr) - LBound(myArr) + 1
    End If
    
    ReDim arr_2D(LBound(myArr) To UBound(myArr), 0)
    For i = LBound(myArr) To UBound(myArr)
        arr_2D(i, 0) = myArr(i)
    Next i
    
    Set myRng = ws.Range(ws.Cells(fRow, fCol), ws.Cells(fRow + arrSizeY - 1, fCol))
    myRng.Value = arr_2D
End Sub

Sub WriteArrToRow(myArr As Variant, fRow As Long, fCol As Long, Optional ws As Worksheet, Optional size As Long)
    Dim arrSizeX As Long
    Dim myRng As Range
    Dim arr_2D As Variant
    Dim i As Long
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    If size > 0 Then
        arrSizeX = size
    Else
        arrSizeX = UBound(myArr) - LBound(myArr) + 1
    End If
    
    ReDim arr_2D(0, LBound(myArr) To UBound(myArr))
    For i = LBound(myArr) To UBound(myArr)
        arr_2D(0, i) = myArr(i)
    Next i
    
    Set myRng = ws.Range(ws.Cells(fRow, fCol), ws.Cells(fRow, fCol + arrSizeX - 1))
    myRng.Value = arr_2D
End Sub

Function SetWorkbook(filePath As String) As Workbook
    Dim fso As New FileSystemObject
    wbName = fso.GetFileName(filePath)
    
    On Error GoTo WBNotOpenYet
    Set SetWorkbook = Workbooks(wbName)
    Exit Function
    
WBNotOpenYet:
    Set SetWorkbook = Workbooks.Open(filePath, ReadOnly:=True)
    
End Function

Function LastNonEmptyColumn(ws As Worksheet, rowNum As Long) As Long
    Dim lRow As Long, n As Long, i As Long
    
    
'    'Here we look in Column A
'    Dim i As Long
'    For i = LastRow To 1 Step -1
'        If Not (IsEmpty(Cells(i, 1))) Then Exit For
'    Next i
'    LastRow = i
    n = Application.CountA(ws.Rows(rowNum))
    'n = ws.Rows(rowNum).SpecialCells(xlCellTypeConstants).count
    i = 0
    lRow = 1
    Do While i < n
        If ws.Cells(rowNum, lRow).text <> "" Then
            i = i + 1
        End If
        lRow = lRow + 1

    Loop
    LastNonEmptyColumn = lRow - 1

End Function
