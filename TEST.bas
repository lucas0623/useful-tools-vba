Attribute VB_Name = "TEST"
Private genFuncs As clsGeneralFunctions
'Try sync project
Sub TESTING()
    Set genFuncs = New clsGeneralFunctions
    Dim i As Long, var As Variant ', arrL As Object
    Dim df As clsDataFrame, df2 As clsDataFrame
    Dim col_eleList As Long
    Set df = New clsDataFrame
    Set df2 = New clsDataFrame
    'Set arrL = CreateObject("System.Collections.ArrayList")
    'var = Array("1", "2", "3", "A")
    'var = FillNumRange(150, 160)
    
    'arrL.Add CVar(var(1))
    df.Init_ReadWorksheet
    col_eleList = df.columnNum("ElemList")
    For i = 1 To df.CountRows
        var = BreakList(df.idata(i, col_eleList), " ", "to")
        'df2.AddDataFrame (df.MultiplyRow(i, "ElemList", var))
        df2.AddDataFrame df.MultiplyRow(i, "ElemList", var)
    Next i
    
    df2.WriteToRange tarRng:=Range("R4"), isWriteRowTag:=False
End Sub
Private Function BreakList(str As String, _
                        Optional oper_and As String = ",", Optional oper_to As String = ":") As Variant
    'this function will return the array of number represented by the input string
    'Dim isReturn() As Boolean
    Dim arrL As Object 'the arrL that contains the return array
    Dim arr As Variant, arr2 As Variant, arr3 As Variant
    Dim i As Long, j As Long, count As Long
    
    count = 0
    Set arrL = CreateObject("System.Collections.ArrayList")
    
    arr = Split(str, oper_and)
    
    For i = 0 To UBound(arr)
        If InStr(1, arr(i), oper_to, vbTextCompare) Then
            arr2 = Split(arr(i), oper_to)
'            If InStr(1, arr2(0), "not", vbTextCompare) Then
'
'            Else
'
'            End If
            arr3 = genFuncs.CStr_arr(FillNumRange(CLng(arr2(0)), CLng(arr2(1))))
            genFuncs.AddArrToArrList arrL, arr3
        
        Else
            arrL.Add CStr(arr(i))
        End If
    
    Next i
    
    BreakList = arrL.ToArray
End Function

Private Function FillNumRange(lower As Long, upper As Long) As Long()
    Dim arr() As Long, i As Long
    ReDim arr(upper - lower)
    For i = 0 To upper - lower
        arr(i) = lower + i
    Next i
    FillNumRange = arr
End Function
