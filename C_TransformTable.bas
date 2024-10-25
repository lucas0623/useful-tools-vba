Attribute VB_Name = "C_TransformTable"
Private genFunc As clsGeneralFunctions
Private wsInteract As clsWorksheetsInteraction
Private UIManager As clsUIManager

Sub Init()
    Set genFunc = New clsGeneralFunctions
    Set wsInteract = New clsWorksheetsInteraction
    Set UIManager = New clsUIManager
End Sub

Sub CondenseToCol()
    Init
    
    Dim colRef As Range, colRef_dest As Range
    Dim data_2D As Variant, data_1D As Variant
    Dim i As Long, j As Long
    Dim numX As Long, numY As Long
    Dim count As Long
    
    Set colRef_data = UIManager.AddInputBox_rng(prompt:="PLEASE SELECT THE DATA.", default:=Selection)
    If colRef_data Is Nothing Then EndSub
    
    Set colRef_dest = UIManager.AddInputBox_rng(prompt:="PLEASE SELECT THE DESTINATION FOR OUTPUT.")
    If colRef_dest Is Nothing Then EndSub
    
    data_2D = colRef_data.Value
    numX = UBound(data_2D, 2) - LBound(data_2D, 2) + 1
    numY = UBound(data_2D, 1) - LBound(data_2D, 1) + 1
    ReDim data_1D(numX * numY - 1)
    
    For i = LBound(data_2D, 2) To UBound(data_2D, 2)
        For j = LBound(data_2D, 1) To UBound(data_2D, 1)
            data_1D(count) = data_2D(j, i)
            count = count + 1
        Next j
    Next i
    
    wsInteract.WriteArrToColumn data_1D, colRef_dest.row, colRef_dest.Column, colRef_dest.Worksheet
    
End Sub


Sub CondenseColumnInTable()
    On Error GoTo ErrHandler:
    Init
    
    Dim df_main As clsDataFrame
    Dim ws_main As Worksheet, ws_output As Worksheet
    Dim rng_rowTag As Range, rng_colTag As Range
    Dim rng_output As Range
    
    Dim arr_list() As String, var As Variant
    
    Set rng_rowTag = UIManager.AddInputBox_rng("Select the COLUMN that contains 'reference row tags'", default:=Columns(1))
    If rng_rowTag Is Nothing Then EndSub
    
    Set rng_colTag = UIManager.AddInputBox_rng("Select ROWS that contains 'heads tags'", default:=Rows(1))
    If rng_colTag Is Nothing Then EndSub
    
    Set rng_list = UIManager.AddInputBox_rng("Select the heads tags of the table that represents the columns to be condensed", default:=Range("D1:H1"))
    If rng_list Is Nothing Then EndSub
    
    Set rng_output = UIManager.AddInputBox_rng("Select the target output location", default:=Range("A1"))
    If rng_output Is Nothing Then EndSub
    
    Set ws_main = rng_rowTag.Worksheet
    Set ws_output = rng_output.Worksheet
    
    var = rng_list.Value
    arr_list = genFunc.CStr_arr(genFunc.Condense2DArrTo1D(var))
    
    Set df_main = New clsDataFrame
    df_main.Init_ReadWorksheet ws_main, rng_rowTag.Column, rng_colTag.row
    
    df_main.CondenseToSingleCol arr_list, "new Column", "Condensed Column"
    df_main.WriteToRange ws_output, rng_output
    
    Exit Sub
    
ErrHandler:
    ErrorMsg
    
End Sub

Sub CondenseToTable()
On Error GoTo ErrHandler:
    'This sub condense the data from multiple "3D" table to a table.
    'The tables are read through user specified row_tags and column tags
    Init
    
    Dim rng_temp As Range
    Dim row_tag() As Long, rng_rowTag As Range, rng_output As Range
    Dim col_tag As Long, i As Long
    Dim df As clsDataFrame
    Set df = New clsDataFrame
    
    Set rng_temp = UIManager.AddInputBox_rng("Select the COLUMN that contains 'reference row tags'", default:=Columns(1))
    If rng_temp Is Nothing Then EndSub
    col_tag = rng_temp.Column
    
    Set rng_rowTag = UIManager.AddInputBox_rng("Select ROWS that contains 'heads tags'", default:=Rows(1))
    If rng_rowTag Is Nothing Then EndSub
    Set rng_output = UIManager.AddInputBox_rng("Select the target output location", default:=Range("A1"))
    If rng_output Is Nothing Then EndSub
    
    ReDim row_tag(rng_rowTag.Rows.count - 1)
    For i = 0 To UBound(row_tag)
        row_tag(i) = rng_rowTag.row + i
    Next i
    
    df.Init_ReadWS_3D rng_rowTag.Worksheet, col_tag, row_tag, False
    df.WriteToRange rng_output.Worksheet, rng_output
    Exit Sub
ErrHandler:
    ErrorMsg
End Sub

Sub MultiplyTable()
On Error GoTo ErrHandler:
    Init
    
    Dim df_main As clsDataFrame
    Dim ws_main As Worksheet, ws_output As Worksheet
    
    Dim rng_rowTag As Range, rng_colTag As Range
    Dim rng_output As Range, rng_list As Range
    Dim str_case As String, rng_case As Range
    Dim arr_list() As String, var As Variant
    
    
    Set rng_rowTag = UIManager.AddInputBox_rng("Select the COLUMN that contains 'reference row tags'", default:=Columns(1))
    If rng_rowTag Is Nothing Then EndSub
    
    Set rng_colTag = UIManager.AddInputBox_rng("Select ROWS that contains 'heads tags'", default:=Rows(1))
    If rng_colTag Is Nothing Then EndSub
    
    Set rng_case = UIManager.AddInputBox_rng("Select the heads tag of the table that represents the column of the changing field", default:=Range("G1"))
    If rng_case Is Nothing Then EndSub
    str_case = rng_case.text
    
    Set rng_list = UIManager.AddInputBox_rng("Select the list of changing value", default:=Range("D1:H1"))
    If rng_list Is Nothing Then EndSub
    
    Set rng_output = UIManager.AddInputBox_rng("Select the target output location", default:=Range("A1"))
    If rng_output Is Nothing Then EndSub
    
    Set ws_main = rng_colTag.Worksheet
    Set ws_output = rng_output.Worksheet
    
    
    Set df_main = New clsDataFrame
    df_main.Init_ReadWorksheet ws_main, rng_rowTag.Column, rng_colTag.row
    var = rng_list.Value
    If IsArray(var) Then
        arr_list = genFunc.CStr_arr(genFunc.Condense2DArrTo1D(var))
    Else
        arr_list = genFunc.CStr_arr(Array(var))
    End If
    'arr_list = genFunc.CStr_arr(df_list.Column(tag_list))
    df_main.MultiplyDataFrame str_case, arr_list
    df_main.WriteToRange ws_output, rng_output
    
    Exit Sub
ErrHandler:
    ErrorMsg
End Sub

Sub FilterTable()
    'On Error GoTo ErrHandler:
    Init
    
    Dim df_main As clsDataFrame
    Dim ws_main As Worksheet, ws_output As Worksheet
    
    Dim rng_rowTag As Range, rng_colTag As Range
    Dim rng_output As Range, rng_list As Range
    Dim str_case As String, rng_case As Range
    Dim arr_list() As String, var As Variant
    
    
    Set rng_rowTag = UIManager.AddInputBox_rng("Select the COLUMN that contains 'reference row tags'", default:=Columns(1))
    If rng_rowTag Is Nothing Then EndSub
    
    Set rng_colTag = UIManager.AddInputBox_rng("Select ROWS that contains 'heads tags'", default:=Rows(1))
    If rng_colTag Is Nothing Then EndSub
    
    Set rng_case = UIManager.AddInputBox_rng("Select the heads tag of the table that represents the column of applying filter", default:=Range("G1"))
    If rng_case Is Nothing Then EndSub
    str_case = rng_case.text
    
    Set rng_list = UIManager.AddInputBox_rng("Select the list for string to be filtered out", default:=Range("D1:H1"))
    If rng_list Is Nothing Then EndSub
    
    Set rng_output = UIManager.AddInputBox_rng("Select the target output location", default:=Range("A1"))
    If rng_output Is Nothing Then EndSub
    
    Set ws_main = rng_colTag.Worksheet
    Set ws_output = rng_output.Worksheet
    
    
    Set df_main = New clsDataFrame
    df_main.Init_ReadWorksheet ws_main, rng_rowTag.Column, rng_colTag.row
    var = rng_list.Value
    If IsArray(var) Then
        arr_list = genFunc.CStr_arr(genFunc.Condense2DArrTo1D(var))
    Else
        arr_list = genFunc.CStr_arr(Array(var))
    End If
    'arr_list = genFunc.CStr_arr(df_list.Column(tag_list))
    Set df_main = df_main.Filter(df_main.columnNum(str_case), arr_list)
    df_main.WriteToRange ws_output, rng_output
    
    Exit Sub
ErrHandler:
    ErrorMsg
End Sub

Sub DivideTable()
    'On Error GoTo ErrHandler:
    Init
    
    Dim df_main As clsDataFrame, df_temp As clsDataFrame
    Dim ws_main As Worksheet, ws_output As Worksheet
    
    Dim rng_rowTag As Range, rng_colTag As Range
    Dim rng_output As Range ', rng_list As Range
    Dim str_case As String, rng_case As Range
    Dim arr_list() As String, var As Variant
    Dim prefix As String, suffix As String
    Dim count As Long, fRow As Long, fCol As Long 'first Row/Column for output
    
    Set rng_rowTag = UIManager.AddInputBox_rng("Select the COLUMN that contains 'reference row tags'", default:=Columns(1))
    If rng_rowTag Is Nothing Then EndSub
    
    Set rng_colTag = UIManager.AddInputBox_rng("Select ROWS that contains 'heads tags'", default:=Rows(1))
    If rng_colTag Is Nothing Then EndSub
    
    Set rng_case = UIManager.AddInputBox_rng("Select the heads tag of the table that represents the column of dividing the table", default:=Range("G1"))
    If rng_case Is Nothing Then EndSub
    str_case = rng_case.text
    
    prefix = UIManager.AddInputBox_str("Prefix=")
    If StrPtr(prefix) = 0 Then EndSub
    suffix = UIManager.AddInputBox_str("Suffix=")
    If StrPtr(suffix) = 0 Then EndSub
    'Set rng_list = UIManager.AddInputBox_rng("Select the list for string to be filtered out", default:=Range("D1:H1"))
    'If rng_list Is Nothing Then EndSub
    
    Set rng_output = UIManager.AddInputBox_rng("Select the target output location", default:=Range("A1"))
    If rng_output Is Nothing Then EndSub
    fRow = rng_output.row
    fCol = rng_output.Column
    
    Set ws_main = rng_colTag.Worksheet
    Set ws_output = rng_output.Worksheet
    
    
    Set df_main = New clsDataFrame
    Set df_temp = New clsDataFrame
    df_main.Init_ReadWorksheet ws_main, rng_rowTag.Column, rng_colTag.row
    var = genFunc.DeDupeOneDimArray(df_main.Column(str_case))
    count = 0
    For i = LBound(var) To UBound(var)
'        If IsArray(var) Then
'            arr_list = genFunc.CStr_arr(genFunc.Condense2DArrTo1D(var))
'        Else
'            arr_list = genFunc.CStr_arr(Array(var))
'        End If
        'arr_list = genFunc.CStr_arr(df_list.Column(tag_list))
        prefix = "*USE-STLD,"
        ws_output.Cells(fRow + count, fCol) = prefix & var(i) & suffix
        ws_output.Cells(fRow + count + 1, fCol) = "*Pressure"
        Set df_temp = df_main.Filter(df_main.columnNum(str_case), var(i))
        df_temp.WriteToRange ws_output, ws_output.Cells(fRow + count + 2, fCol), False, False
        count = count + df_temp.CountRows + 2
    Next i
    
    Exit Sub
ErrHandler:
    ErrorMsg
End Sub

Sub FillEmptyCells()
    Dim i As Long, j As Long
    Dim rCol() As Long
    
    ReDim rCol(Selection.Columns.count - 1)
    
    For j = 0 To UBound(rCol)
        rCol(j) = Selection.Column + j
    Next j
    
    For i = Selection.row To Selection.row + Selection.Rows.count - 1
        For j = rCol(0) To rCol(UBound(rCol))
            If Cells(i, j) = "" Then Cells(i, j) = Cells(i - 1, j)
        Next j
    Next i
End Sub

Sub DelEmptyRows()
    Dim rng As Range
    Init
    Set rng = UIManager.AddInputBox_rng(prompt:="PLEASE SELECT THE DATA.", default:=Selection)
    If rng Is Nothing Then EndSub
    
    DelRows rng
End Sub

Sub DelRowsIfEqual()
    Dim rng As Range, val As Variant
    Init
    Set rng = UIManager.AddInputBox_rng(prompt:="PLEASE SELECT THE DATA.", default:=Selection)
    If rng Is Nothing Then EndSub
    
    val = UIManager.AddInputBox_str(prompt:="Delete Rows with Value = ", defaultStr:="0")
    If val = "False" Then EndSub
    DelRows rng, val
End Sub

Private Sub DelRows(rng As Range, Optional str As Variant = "")
    Dim i As Long, j As Long, count As Long
    Dim rCol As Long, fRow As Long
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    fRow = rng.row
    rCol = rng.Column
    Set ws = rng.Worksheet
    For i = rng.row To rng.row + rng.Rows.count - 1
        If ws.Cells(fRow + count, rCol).text = str Then
            ws.Rows(fRow + count).Delete
            count = count - 1
        End If
        count = count + 1
    Next i
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub

Private Sub EndSub()
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    End
End Sub

Private Sub ErrorMsg()
    MsgBox "Something goes wrong. The macro will be terminated. Please check the Input"
End Sub
