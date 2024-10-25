VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This module contains the data frame class
'Reference was made to python pandas

'Last Update on 30 Nov 2022
'30 Nov 2022: Minor Bug Fix
'29 Nov 2022: Function "DuplicateDataFrame", some "Let Properties","CondenseColumns"
'28 Nov 2022: Fixed bug for the "Columns" Property
'25 Nov 2022: Read "3D" tables and condense the data into a single table, condense df to 1 column, and combining dfs.
'26 Oct 2022: Revised "iColumns" return array lBound and UBound number, add "heads_slice" property
'25 Oct 2022: Add "Init_ReadWorksheet" method to initialize the object
'5 Sept 2022: add "idata_multiple" subroutine
'*************************************************************************
Private genFuncs As clsGeneralFunctions
Private wsInteraction As clsWorksheetsInteraction
Private pHeads() As String 'first row of the table (optional)
Private pRowTags() As String 'first column of the table (optional)
Private pData() As Variant '(i,j) where i is the row number, j is the column number

Sub Class_Initialize()
    Set genFuncs = New clsGeneralFunctions
    Set wsInteraction = New clsWorksheetsInteraction
End Sub

Sub Init_ReadWS_3D(ws As Worksheet, col_tag As Long, row_tag() As Long, Optional isLastRowOfWs As Boolean = False)
    Dim i As Long, j As Long
    Dim rng_rRow() As Range 'rRow for each table
    Dim tempDF_1D As clsDataFrame
    Dim tempDF() As clsDataFrame
    With ws
        rng_rRow = wsInteraction.FindNonEmptyCell(.Columns(col_tag))
    End With
    
    ReDim tempDF(0 To UBound(rng_rRow))
    For i = 0 To UBound(rng_rRow)
        Set tempDF(i) = New clsDataFrame
        For j = 0 To UBound(row_tag)
            Set tempDF_1D = New clsDataFrame
            tempDF_1D.Init_ReadWorksheet ws, row_tag:=row_tag(j), rRow:=rng_rRow(i).row, isLastRowOfWs:=isLastRowOfWs
            tempDF_1D.CondenseToSingleCol
            tempDF(i).AddDataFrame tempDF_1D, False
            
        Next j
        Me.AddDataFrame tempDF(i), True
    Next i
    
    
End Sub

'This read a worksheet data according to the heading specified in the 1st row and 1st column of the worksheet
'The data will be save to the dataframe.
Sub Init_ReadWorksheet(Optional ws As Worksheet, Optional col_tag As Long = 1, Optional row_tag As Long = 1, Optional rRow As Long, Optional isLastRowOfWs As Boolean = False)
    'if no heads are specified, all the labels at the first row of the worksheet will be read
    'first row of data will be the reference row (first label in the first column) +1.

    Dim count As Long, i As Long, j As Long
    Dim fRow As Long, lRow As Long
    Dim rng_heads() As Range, rng_rRow As Range
    
    If ws Is Nothing Then Set ws = ActiveSheet
    
    With ws
        'get the range with head tag

        'If UBound(heads) = -1 Then
            rng_heads = wsInteraction.FindNonEmptyCell(.Rows(row_tag))
        'End If
        If Not genFuncs.isInitialised(rng_heads) Then
            Debug.Print "No head tag is found in the first row of the " & ws.Name & " Worksheet!!!!"""
            Exit Sub
        End If
        
        'get the fRow
        If rRow < 1 Then
            Set rng_rRow = wsInteraction.FindFirstNonEmptyCell(.Columns(col_tag))
            If rng_rRow.text = "" Then
                Debug.Print "No rowtag is found in the first column of the " & ws.Name & " Worksheet!!!!"""
                Exit Sub
            End If
            fRow = rng_rRow.row + 1
        Else
            fRow = rRow + 1
        End If
        
        'Get the lRow of the table
        If Not isLastRowOfWs Then
            lRow = wsInteraction.FindLastRow(fRow, rng_heads(0).Column, ws)
        Else
            lRow = wsInteraction.FindLastRowOfWS(rng_heads(0).Column, ws)
        End If
        
        'set the dataframe
        ReDim pHeads(1 To UBound(rng_heads) + 1)
        ReDim pRowTags(1 To lRow - fRow + 1)
        ReDim pData(1 To UBound(pRowTags), 1 To UBound(pHeads))
        
        Dim cell As Variant
        count = 1
        For Each cell In rng_heads
            pHeads(count) = cell.text
            count = count + 1
        Next
        
        For i = LBound(pRowTags) To UBound(pRowTags)
            pRowTags(i) = CStr(i)
        Next i
        
        For i = 1 To UBound(pRowTags)
            For j = LBound(pHeads) To UBound(pHeads)
                pData(i, j) = .Cells(fRow + i - 1, rng_heads(j - 1).Column)
            Next j
        Next i
    End With
    
End Sub

Sub Init_byArr(data As Variant, isHvHead As Boolean, isHvRowTag As Boolean)
    Dim i As Long, j As Long, count As Long, count2 As Long
    Dim headStart As Long
    Dim rowTagAdj As Long
    'check if the input data is valid
    
    If isHvHead Then
        rowTagAdj = 0
    Else
        rowTagAdj = 1
    End If
    
    If isHvRowTag Then
        headStart = 0
    Else
        headStart = 1
    End If
    
    ReDim pHeads(headStart To UBound(data, 2) - LBound(data, 2) + headStart)
    ReDim pRowTags(1 To UBound(data, 1) - LBound(data, 1) + rowTagAdj)
    ReDim pData(1 To UBound(pRowTags), 1 To UBound(pHeads))
    
    'assign the HEAD data into the dataframe
    If isHvHead Then
        count = headStart
        For i = LBound(data, 2) To UBound(data, 2)
            pHeads(count) = data(LBound(data, 1), i)
            count = count + 1
        Next i
    Else
        For i = LBound(pHeads) To UBound(pHeads)
            pHeads(i) = CStr(i)
        Next i
    End If
    
    
    'assign the rowtag data into the dataframe
    If isHvRowTag Then
        count = 1
        For i = LBound(data, 1) + 1 - rowTagAdj To UBound(data, 1)
            pRowTags(count) = data(i, LBound(data, 2))
            count = count + 1
        Next i
    Else
        For i = LBound(pRowTags) To UBound(pRowTags)
            pRowTags(i) = CStr(i)
        Next i
    End If
    
    
    'assign the table data into the dataframe
    count = 1
    For i = LBound(data, 1) + 1 - rowTagAdj To UBound(data, 1)
        count2 = 1
        For j = LBound(data, 2) + 1 - headStart To UBound(data, 2)
            pData(count, count2) = data(i, j)
            count2 = count2 + 1
        Next j
        count = count + 1
    Next i
'        'assign the table data into the dataframe
'    count = 1
'    For i = LBound(data, 1) + 1 - headStart To UBound(data, 1)
'        count2 = 1
'        For j = LBound(data, 2) + 1 - rowTagAdj To UBound(data, 2)
'            pData(count, count2) = data(i, j)
'            count2 = count2 + 1
'        Next j
'        count = count + 1
'    Next i
End Sub
Sub Init(data As Variant, Heads() As String, Optional rowTags As Variant)
    pData = data
    pHeads = Heads
    If genFuncs.isInitialised(rowTags) Then
        pRowTags = rowTags
    Else
        ReDim pRowTags(1 To UBound(data, 1) - LBound(data, 1) + 1)
        For i = LBound(pRowTags) To UBound(pRowTags)
            pRowTags(i) = CStr(i)
        Next i
    End If
End Sub
Sub WriteToRange(Optional tarWS As Worksheet, Optional tarRng As Range, Optional isWriteRowTag As Boolean = True, Optional isWriteHeads As Boolean = True)
    If tarWS Is Nothing Then Set tarWS = Sheets.Add
    If tarRng Is Nothing Then Set tarRng = tarWS.Range("A1")
    Dim colAdj As Long, rowAdj As Long
    
    If isWriteRowTag Then
        colAdj = 1
    Else
        colAdj = 0
    End If
    If isWriteHeads Then
        rowAdj = 1
    Else
        rowAdj = 0
    End If
    
    If isWriteHeads Then wsInteraction.WriteArrToRow pHeads, tarRng.row, tarRng.Column + colAdj, tarWS
    If isWriteRowTag Then wsInteraction.WriteArrToColumn pRowTags, tarRng.row + rowAdj, tarRng.Column, tarWS
    wsInteraction.WriteArrToRngTwoDim pData, tarRng.row + rowAdj, tarRng.Column + colAdj, tarWS
End Sub


Property Get columnNum(headTag As String) As Long 'get the column number from head tag
    columnNum = GetArrNum(headTag, pHeads)
End Property

Property Get columnsNum(headTags() As String) As Long() 'get the column number from head tag
    Dim i As Long, arr() As Long
    Dim count As Long
    count = 1
    ReDim arr(1 To UBound(headTags) - LBound(headTags) + 1)
    For i = LBound(headTags) To UBound(headTags)
        arr(count) = GetArrNum(headTags(i), pHeads)
        count = count + 1
    Next i
    columnsNum = arr
End Property

Property Get rowNum(rowTag As String) As Long 'get the row number from head tag
    rowNum = GetArrNum(rowTag, pRowTags)
End Property

Property Get data(rowTag As String, headTag As String, Optional isAlldata As Boolean = False) As Variant 'this should be able to set the range
    If isAlldata Then
        data = pData
    Else
        data = pData(GetArrNum(rowTag, pRowTags), GetArrNum(headTag, pHeads))
    End If
End Property


Property Get idata(rowNum As Long, colNum As Long) As Variant 'return the data of a cell (1 data) using the rowNum and ColumnNum
    
    idata = pData(rowNum, colNum)
    
End Property

Property Let idata(rowNum As Long, colNum As Long, val As Variant)
     pData(rowNum, colNum) = val
End Property
Property Get idata_multiple(rowNum As String, colNum As String, Optional is2DArr As Boolean = False) As Variant 'return the data of n cells using the rowNum and ColumnNum. Supporting Slicer Input
    Dim arrSlice_Row As Variant, arrSlice_Col As Variant
    Dim arr As Variant
    arrSlice_Row = Slicer(LBound(pRowTags), UBound(pRowTags), rowNum)
    arrSlice_Col = Slicer(LBound(pHeads), UBound(pHeads), colNum)
    
    If UBound(arrSlice_Col) - LBound(arrSlice_Col) = 0 And UBound(arrSlice_Row) - LBound(arrSlice_Row) = 0 Then
        idata_multiple = pData(rowNum, colNum)
    ElseIf UBound(arrSlice_Col) - LBound(arrSlice_Col) = 0 And Not is2DArr Then
        ReDim arr(UBound(arrSlice_Row) - LBound(arrSlice_Row))
        adj_row = LBound(arrSlice_Row) - LBound(arr)
        For i = LBound(arrSlice_Row) To UBound(arrSlice_Row)
            arr(i - adj_row) = pData(CLng(arrSlice_Row(i)), CLng(colNum))
        Next i
    ElseIf UBound(arrSlice_Row) - LBound(arrSlice_Row) = 0 And Not is2DArr Then
        ReDim arr(UBound(arrSlice_Col) - LBound(arrSlice_Col))
        adj_col = LBound(arrSlice_Col) - LBound(arr)
        For i = LBound(arrSlice_Col) To UBound(arrSlice_Col)
            arr(i - adj_col) = pData(CLng(rowNum), CLng(arrSlice_Col(i)))
        Next i
    Else
        ReDim arr(UBound(arrSlice_Row) - LBound(arrSlice_Row), UBound(arrSlice_Col) - LBound(arrSlice_Col))
        adj_col = LBound(arrSlice_Col) - LBound(arr)
        adj_row = LBound(arrSlice_Row) - LBound(arr)
        For i = LBound(arrSlice_Row) To UBound(arrSlice_Row)
            For j = LBound(arrSlice_Col) To UBound(arrSlice_Col)
                arr(i - adj_row, j - adj_col) = pData(CLng(arrSlice_Row(i)), CLng(arrSlice_Col(j)))
            Next j
        Next i
    End If


    idata_multiple = arr
    
End Property
Property Get head(colNum As Long) As String 'return the head tag of the tables using colNum
    head = pHeads(colNum)
End Property

Property Get Heads() As String()
        Heads = pHeads()

End Property
Property Let Heads(arr() As String) 'to assign head tags to the data
    Dim i As Integer
    Dim diff As Integer
    diff = LBound(arr) - LBound(pHeads)
    For i = LBound(pHeads) To UBound(pHeads)
        pHeads(i) = arr(i + diff)
    Next i
End Property

Property Get heads_slice(Optional colNum As String) As String()
    If colNum = "" Then
        Heads = pHeads()
    Else
        Dim arr() As String, arrSlice As Variant 'the array to contain the result data
        Dim i As Long, j As Long
        Dim adj As Long
    
        arrSlice = Slicer(LBound(pHeads), UBound(pHeads), colNum)
    
        ReDim arr(UBound(arrSlice) - LBound(arrSlice))
        For i = LBound(arrSlice) To UBound(arrSlice)
            arr(i) = head(CLng(arrSlice(i)))
        Next i
        heads_slice = arr
    End If
End Property


Property Get rowTag(rowNum As Long) As String
    rowTag = pRowTags(rowNum)
End Property

Property Get rowTags() As String()
    rowTags = pRowTags
    
    
End Property

Property Let rowTags(arr() As String)
    Dim i As Integer
    Dim diff As Integer
    diff = LBound(arr) - LBound(pRowTags)
    For i = LBound(pRowTags) To UBound(pRowTags)
        pRowTags(i) = arr(i + diff)
    Next i
End Property

Property Get row(rowNum As Long, Optional rowTag As String) As Variant
    Dim arr As Variant
    Dim i As Long
    
    If rowTag <> "" Then
        rowNum = GetArrNum(rowTag, pRowTags)
    End If
    
    ReDim arr(LBound(pHeads) To UBound(pHeads))
    For i = LBound(pHeads) To UBound(pHeads)
        arr(i) = pData(rowNum, i)
    Next i
    
    row = arr
End Property

Property Get iRows(rowNum As String) As Variant 'return the data of all rows of specified columns. Slicer input is allowed
    
    Dim arr As Variant, arrSlice As Variant 'the array to contain the result data
    Dim i As Long, j As Long
    Dim adj As Long
    
    arrSlice = Slicer(LBound(pRowTags), UBound(pRowTags), rowNum)
    
    
    If UBound(arrSlice) - LBound(arrSlice) > 0 Then
        ReDim arr(UBound(arrSlice) - LBound(arrSlice), UBound(pHeads) - LBound(pHeads))
        adj = LBound(arrSlice) - LBound(arr)
        For i = LBound(arrSlice) To UBound(arrSlice)
            For j = LBound(pHeads) To UBound(pHeads)
                arr(i - adj, j - 1) = pData(CLng(arrSlice(i)), j)
            Next j
        Next i
    Else
        ReDim arr(LBound(pHeads) To UBound(pHeads))
        For i = LBound(pHeads) To UBound(pHeads)
            arr(i) = pData(CLng(arrSlice(LBound(arrSlice))), i)
        Next i
    End If

    iRows = arr
End Property

Property Get Columns(headTags() As String) As Variant
    
    Dim arr As Variant
    Dim i As Long, j As Long
    Dim adj As Long

'    If headTags <> "" Then
'        colNum = GetArrNum(headTag, pHeads)
'    End If
    If UBound(headTags) - LBound(headTags) > 0 Then
        ReDim arr(LBound(pRowTags) To UBound(pRowTags), 1 To UBound(headTags) - LBound(headTags) + 1)
        adj = LBound(headTags) - LBound(arr)
        
        For i = LBound(pRowTags) To UBound(pRowTags)
            For j = LBound(headTags) To UBound(headTags)
                colNum = GetArrNum(headTags(j), pHeads)
                arr(i, j - adj) = pData(i, colNum)
            Next j
        Next i
'    Else
'        ReDim arr(UBound(pRowTags) - LBound(pRowTags))
'        colNum = GetArrNum(headTags(LBound(headTags)), pHeads)
'        For i = LBound(pRowTags) To UBound(pRowTags)
'            arr(i) = pData(i, colNum)
'        Next i
    End If
    
    Columns = arr
End Property
Property Get Column(headTag As String) As Variant
    
    Dim arr As Variant
    Dim i As Long, j As Long
    Dim adj As Long

'    If headTags <> "" Then
'        colNum = GetArrNum(headTag, pHeads)
'    End If

        ReDim arr(LBound(pRowTags) To UBound(pRowTags))
        colNum = GetArrNum(headTag, pHeads)
        If colNum > -1 Then
            For i = LBound(pRowTags) To UBound(pRowTags)
                arr(i) = pData(i, colNum)
            Next i
        End If
        
    Column = arr
End Property

Property Let Column(headTag As String, val As Variant)
    
    Dim i As Long, count As Long
    Dim colNum As Long

'    If headTags <> "" Then
'        colNum = GetArrNum(headTag, pHeads)
'    End If
    colNum = Me.columnNum(headTag)
    
    If IsArray(val) Then
        For i = LBound(val) To UBound(val)
            Me.idata(1 + count, colNum) = val(i)
            count = count + 1
        Next i
    Else
        For i = 1 To Me.CountRows
            Me.idata(i, colNum) = val
        Next i
    End If
End Property

'NOT TESTED YET. better done such operation in an arraylist before assign to the dataframe
'Property Let Columns(headTags As Variant, arr As Variant)
'
'    Dim i As Long, j As Long
'
''    If headTags <> "" Then
''        colNum = GetArrNum(headTag, pHeads)
''    End If
'    If UBound(headTags) - LBound(headTags) > 0 Then
'        For j = 0 To UBound(headTags) - LBound(headTags)
'            colNum = GetArrNum(headTags(LBound(headTags) + j), pHeads)
'            For i = 0 To UBound(pRowTags) - LBound(pRowTags)
'                pData(i + LBound(pRowTags), j + LBound(headTags)) = arr(i + LBound(arr, 1), j + LBound(arr, 2))
'            Next i
'        Next j
'    Else
'        colNum = GetArrNum(headTags(LBound(headTags)), pHeads)
'        For i = LBound(pRowTags) To UBound(pRowTags)
'            pData(i + LBound(pRowTags), colNum) = arr(i + LBound(arr))
'        Next i
'    End If
'
'End Property
Property Get iColumn(colNum As Long) As Variant 'return the data of specified columns
    
    Dim arr As Variant
    Dim i As Long, j As Long
    Dim adj As Long

    ReDim arr(LBound(pRowTags) To UBound(pRowTags))
    
    For i = LBound(pRowTags) To UBound(pRowTags)
        arr(i) = pData(i, colNum)
    Next i

    iColumn = arr
End Property

Property Get iColumns(colNum As String) As Variant 'return the data of all rows of specified columns. Slicer input is allowed
    
    Dim arr As Variant, arrSlice As Variant 'the array to contain the result data
    Dim i As Long, j As Long
    Dim adj As Long
    
    arrSlice = Slicer(LBound(pHeads), UBound(pHeads), colNum)
    
    
    If UBound(arrSlice) - LBound(arrSlice) > 0 Then
        ReDim arr(LBound(pRowTags) To UBound(pRowTags), LBound(arrSlice) To UBound(arrSlice))
        adj = LBound(arrSlice) - LBound(arr, 2)
        For i = LBound(pRowTags) To UBound(pRowTags)
            For j = LBound(arrSlice) To UBound(arrSlice)
                arr(i, j - adj) = pData(i, CLng(arrSlice(j)))
            Next j
        Next i
    Else
        ReDim arr(UBound(pRowTags) - LBound(pRowTags))
        For i = LBound(pRowTags) To UBound(pRowTags)
            arr(i) = pData(i, CLng(arrSlice(LBound(arrSlice))))
        Next i
    End If

    iColumns = arr
End Property

'NOT TESTED YET. better done such operation in an arraylist before assign to the dataframe
'Property Let iColumns(colNum As String, arr As Variant)
'    arrSlice As Variant 'the array to contain the result data
'    Dim i As Long, j As Long
'    Dim adj_i As Long, adj_j As Long
'
'    arrSlice = Slicer(LBound(pHeads), UBound(pHeads), colNum)
'
'    If UBound(arrSlice) - LBound(arrSlice) > 0 Then
'        adj_i = LBound(pRowTags) - LBound(arr, 1)
'        adj_j = LBound(arrSlice) - LBound(arr, 2)
'        For i = LBound(pRowTags) To UBound(pRowTags)
'            For j = LBound(arrSlice) To UBound(arrSlice)
'                pData(i, CLng(arrSlice(j))) = arr(i - adj_i, j - adj_j)
'            Next j
'        Next i
'    Else
'        For i = LBound(pRowTags) To UBound(pRowTags)
'            pData(i, CLng(arrSlice(LBound(arrSlice)))) = arr(i - adj_i)
'        Next i
'    End If
'
'End Property

Property Get CountColumns() As Long
    CountColumns = UBound(pData, 2)
End Property

Property Get FindExtremeRow(col_head As Long, Optional isMax As Boolean = True) As Long
    Dim arr As Variant
    Dim i As Long, j As Long
    Dim row_max As Long
    
    row_max = LBound(pRowTags)
    If isMax Then
        For i = LBound(pRowTags) + 1 To UBound(pRowTags)
            If genFuncs.Compare(pData(row_max, col_head), pData(i, col_head), False) Then
                row_max = i
            End If
        Next i
    Else
        For i = LBound(pRowTags) + 1 To UBound(pRowTags)
            If Not genFuncs.Compare(pData(row_max, col_head), pData(i, col_head), False) Then
                row_max = i
            End If
        Next i
    End If

    FindExtremeRow = row_max
End Property

Sub AddColumns()

End Sub

Sub DeleteColumns()

End Sub
Sub AddRows()

End Sub

Sub DeleteRows()

End Sub
Function Filter(col_head As Long, filterBy As Variant) As clsDataFrame
    Dim coll As Object
    Set coll = CreateObject("System.Collections.ArrayList")
    Dim myArr As Variant
    Dim i As Long, rCol As Long
    rCol = col_head
    
    coll.Add pHeads
    If Not IsArray(filterBy) Then filterBy = Array(filterBy)
    
    For i = LBound(pData, 1) To UBound(pData, 1)
        'For k = LBound(filterBy) To UBound(filterBy)
            If genFuncs.IsInArr(pData(i, rCol), filterBy) Then
                coll.Add (Me.row(i))
                'Exit For
            End If
        'Next k
    Next i
    myArr = genFuncs.ArrListTo2DArr(coll)
    Set Filter = New clsDataFrame
    Filter.Init_byArr myArr, True, False

End Function

'Filter by is the array containing the list of strings
Function Filter_str(head As String, filterBy As Variant) As clsDataFrame
    Dim coll As Object
    Set coll = CreateObject("System.Collections.ArrayList")
    Dim myArr As Variant
    Dim i As Long, rCol As Long
    rCol = columnNum(head)
    
    coll.Add pHeads
    
    For i = LBound(pData, 1) To UBound(pData, 1)
        'For k = LBound(filterBy) To UBound(filterBy)
            If genFuncs.IsInArr(pData(i, rCol), filterBy) Then
                coll.Add (Me.row(i))
                'Exit For
            End If
        'Next k
    Next i
    myArr = genFuncs.ArrListTo2DArr(coll)
    Set Filter_str = New clsDataFrame
    Filter_str.Init_byArr myArr, True, False
    'Set Filter_str = Factory.NewclsDataFrame(myArr, True, False)
End Function

Function SplitDF(col_head As Long) As clsDataFrame()
    Dim arr As Variant
    Dim df() As clsDataFrame
    arr = genFuncs.DeDupeOneDimArray(Me.iColumn(col_head))
    ReDim df(UBound(arr) - LBound(arr))
    For i = LBound(arr) To UBound(arr)
        Set df(i) = Me.Filter(col_head, arr(i))
    Next i
    SplitDF = df
End Function

Sub CondenseToSingleCol(Optional headTags As Variant, Optional newHead1 As String, Optional newHead2 As String)
    'This Sub condense the input columns into 1 columns for the table
    Dim arr As Variant, arr2 As Variant
    Dim i As Long
    
    If Not genFuncs.isInitialised(headTags) Then
        'to condense the whole table
        arr = genFuncs.Condense2DArrTo1D(pData)
        arr = genFuncs.Expand1DArrTo2D(arr)
        ReDim arr2(UBound(arr) + 1, 0)
        arr2(0, 0) = pHeads(1)
        For i = 0 To UBound(arr)
            arr2(i + 1, 0) = arr(i, 0)
        Next i
        Me.Init_byArr arr2, True, False
    Else
        If Not IsArray(headTags) Then headTags = Array(headTags)
        'Save the data in an arraylist and convert it back to a dataframe
            Dim arr_notCondense As Variant, arr_condense As Variant
            Dim colNum_notCondense() As Long, colNum_condense() As Long
            Dim n_notCondense As Long, str_notCondense() As String
            Dim counter As Long
            Dim arrL_data As Object, arr_data As Variant, arr_heads() As String
            
            Set arrL_data = CreateObject("System.Collections.ArrayList")
            headTags = genFuncs.CStr_arr(genFuncs.ChangeArrBase(headTags, 1))
            
            colNum_condense = Me.columnsNum(genFuncs.CStr_arr(headTags))
            str_notCondense = genFuncs.CStr_arr(Operator_Not(Me.Heads, headTags))
            colNum_notCondense = Me.columnsNum(str_notCondense)
            n_notCondense = UBound(colNum_notCondense) - LBound(colNum_notCondense) + 1
            
            'Set the heads of the new dataframe
            ReDim arr_heads(1 To n_notCondense + 2)
            str_notCondense = genFuncs.CStr_arr(genFuncs.ChangeArrBase(str_notCondense, 1))

            For i = 1 To n_notCondense
                arr_heads(i) = str_notCondense(i)
            Next i
            arr_heads(n_notCondense + 1) = newHead1
            arr_heads(n_notCondense + 2) = newHead2
            
            'arr_condense = me.idata_multiple(Cstr(i),)
            ReDim arr_data(1 To UBound(arr_heads))
            counter = 1
            For i = 1 To Me.CountRows
                For j = 1 To UBound(colNum_condense)
                    For k = 1 To n_notCondense
                        arr_data(k) = Me.idata(i, colNum_notCondense(k))
                    Next k
                    arr_data(n_notCondense + 1) = headTags(j)
                    arr_data(n_notCondense + 2) = Me.idata(i, colNum_condense(j))
                    arrL_data.Add arr_data
                Next j
            Next i
            arr_data = genFuncs.ArrListTo2DArr(arrL_data)
            Me.Init_byArr arr_data, False, False
            Me.Heads = arr_heads
        End If
        
End Sub

Sub AddDataFrame(df As clsDataFrame, Optional isBelow As Boolean = True)
    Dim newHeads As Variant, newRowTags As Variant
    
    If genFuncs.isInitialised(pData) Then
        
        If isBelow Then
            newRowTags = genFuncs.AddItemsToArr(pRowTags, df.rowTags)
            pRowTags = genFuncs.CStr_arr(newRowTags)
        Else
            newHeads = genFuncs.AddItemsToArr(pHeads, df.Heads)
            pHeads = genFuncs.CStr_arr(newHeads)
        End If
        
        pData = genFuncs.Combine2DArr(pData, df.data("", "", True), isBelow)
    Else
        Me.Init df.data("", "", True), df.Heads, df.rowTags
    End If
End Sub

Function MultiplyRow(rowNum As Long, str_colChange As String, arr As Variant) As clsDataFrame
    Dim i As Long, j As Long, count As Long
    Dim df As clsDataFrame, colNum As Long
    Dim data As Variant, row As Variant
    
    Set df = New clsDataFrame
    row = Me.row(rowNum)
    colNum = Me.columnNum(str_colChange)
    count = LBound(arr)
'    df.Init genFuncs.TransposeArray(genFuncs.Expand1DArrTo2D(Me.row(rowNum), 1)), Me.Heads
'    df.MultiplyDataFrame str_colChange, arr
    ReDim data(1 To UBound(arr) - LBound(arr) + 1, 1 To UBound(pHeads))
    For i = 1 To UBound(data, 1)
        For j = 1 To UBound(data, 2)
            data(i, j) = row(j)
        Next j
        data(i, colNum) = arr(count)
        count = count + 1
    Next i
    df.Init data, pHeads
    Set MultiplyRow = df
End Function

Sub MultiplyDataFrame(Optional str_colChange As String, Optional arr_changeVal As Variant)
    Dim i As Long, j As Long
    Dim df As clsDataFrame, colNum As Long
    Dim val As Variant
    
    Set df = New clsDataFrame
    df.Init Me.data("", "", True), Me.Heads, Me.rowTags
    'colNum = Me.columnNum(str_colChange)
    For i = LBound(arr_changeVal) To UBound(arr_changeVal)
        val = arr_changeVal(i)
        df.Column(str_colChange) = val
        If i = LBound(arr_changeVal) Then
            Me.Init df.data("", "", True), Me.Heads, Me.rowTags
        Else
            Me.AddDataFrame df
        End If
    Next i
End Sub

Property Get CountRows() As Long
    CountRows = UBound(pRowTags) - LBound(pRowTags) + 1
End Property

Private Function GetArrNum(myStr As String, myStrArr() As String) As Long 'to search for certain string in the data
    Dim tempStr As String
    Dim k As Long
    
    For k = LBound(myStrArr) To UBound(myStrArr)
        If myStr = myStrArr(k) Then
            GetArrNum = k
            Exit Function
        End If
    Next k
    GetArrNum = -1
    'MsgBox "No Match of the string to the array"
End Function




Private Function Slicer(lLimit As Long, uLimit As Long, str As String) As Variant
    'this function will return the array of number represented by the input string
    Dim isReturn() As Boolean
    Dim arr As Variant, arr2 As Variant
    Dim returnArr As Variant
    Dim i As Long, j As Long, count As Long
    'Dim isReturn As Variant
    ReDim isReturn(lLimit To uLimit)
    
    For i = lLimit To uLimit
        isReturn(i) = False
    Next i
    count = 0
    
    arr = Split(str, ",")
    
    For i = 0 To UBound(arr)
        If InStr(1, arr(i), ":", vbTextCompare) Then
            arr2 = Split(arr(i), ":")
            If InStr(1, arr2(0), "not", vbTextCompare) Then
                
            Else
                
            End If
            If arr2(0) = "" Then
                For j = lLimit To CLng(arr2(1))
                    isReturn(j) = True
                Next j
            ElseIf arr2(1) = "" Then
                For j = CLng(arr2(0)) To uLimit
                    isReturn(j) = True
                Next j
            Else
                For j = CLng(arr2(0)) To CLng(arr2(1))
                    isReturn(j) = True
                Next j
            End If
        Else
            isReturn(CLng(arr(i))) = True
        End If
    
    Next i
    
    'assign the true into a single array
    For i = lLimit To uLimit
        If isReturn(i) Then
            count = count + 1
        End If
    Next i
    
    ReDim returnArr(count - 1)
    count = 0
    For i = lLimit To uLimit
        If isReturn(i) Then
            returnArr(count) = i
            count = count + 1
        End If
    Next i
    
    Slicer = returnArr
End Function

Private Function Operator_Not(fullList As Variant, arr_not As Variant) As Variant
    Dim i As Long, j As Long
    Dim arrL_all As Object
    Set arrL_all = CreateObject("System.Collections.ArrayList")
    
    For j = LBound(fullList) To UBound(fullList)
        arrL_all.Add CStr(fullList(j))
    Next j
    If Not IsArray(arr_not) Then arr_not = Array(arr_not)

    For i = LBound(arr_not) To UBound(arr_not)
        For j = LBound(fullList) To UBound(fullList)
            If arr_not(i) = arrL_all.Item(j) Then
                arrL_all.RemoveAt j
                Exit For
            End If
        Next j
    Next i
    'ReDim returnArr(arrL_all.count - 1)
    Operator_Not = arrL_all.ToArray
End Function

'Sub Init_byRng(rng As Range, isHvHead As Boolean, isHvRowTag As Boolean)
'    SetArrayTwoDimDynamic rng.Row, rng.Column, xxxxxxxxxxxxxxxx, rng.Worksheet
'    Init_byArr
'End Sub

'Private Sub SetArrayTwoDimDynamic(ByVal fRow As Integer, ByVal fCol As Integer, ByRef myArr As Variant, ByVal isGrpByRow As Boolean, Optional ws As Worksheet)
'    Dim lRow As Long
'    Dim lCol As Long
'    lRow = FindLastRow(fRow, fCol, ws)
'    lCol = FindLastCol(fRow, fCol, ws)
'    If isGrpByRow Then
'        ReDim myArr(lRow - fRow, lCol - fCol)
'    Else: ReDim myArr(lCol - fCol, lRow - fRow)
'    End If
'    With ws
'        For i = 0 To UBound(myArr, 1)
'            For j = 0 To UBound(myArr, 2)
'                If isGrpByRow Then
'                    myArr(i, j) = .Cells(fRow + i, fCol + j)
'                Else
'                    myArr(i, j) = .Cells(fRow + j, fCol + i)
'                End If
'            Next j
'        Next i
'    End With
'End Sub



