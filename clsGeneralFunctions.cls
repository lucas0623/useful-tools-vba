VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGeneralFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************
'The module contains some useful functions in macro development
'Arthor: Lucas LEUNG
'Update Log
'28 Nov 2022 - Add Function "Max" and "Min" and "ChangeArrBase", "CreatDictOfObj"
'25 Nov 2022 - Add Function "Condense 2D arr to 1D arr", "expandarrTo2DArr", "Combine2DArr", Modify Function "AddItemToArr"
'21 Nov 2022 - Update join str function
'4 Nov 2022 - Add Functions to convert Variant Array to Dobule/String/Long array
'26 Oct 2022 - Add Functions "getArrDimension" and "Get1DArrFrom2DArr"
'31 Jul 2022 - Add Functions to change varaint array to double/string array

'*******************************************************
'if isRow = true, will search the first row

Function CreateArray(ParamArray arr() As Variant) As Variant
    Dim i As Long, j As Long, arrL As Object
    Dim str2 As String
    Set arrL = CreateObject("System.Collections.ArrayList")
    For i = LBound(arr) To UBound(arr)
        If isInitialised(arr(i)) Then
            For j = LBound(arr(i)) To UBound(arr(i))
                arrL.Add CVar(arr(i)(j))
            Next j
        Else
            arrL.Add arr(i)
        End If
    Next i
    CreateArray = arrL.ToArray
End Function

Function SortArray(arr As Variant, Optional isA2Z As Boolean = True) As Variant
    Dim arrL As Object, i As Long
    Set arrL = CreateObject("System.Collections.ArrayList")
    For i = LBound(arr) To UBound(arr)
        arrL.Add arr(i)
    Next i
    arrL.Sort
    If Not isA2Z Then arrL.Reverse
    SortArray = arrL.ToArray
End Function

Function TransposeArray(MyArray As Variant) As Variant
    Dim x As Long, Y As Long
    Dim maxX As Long, minX As Long
    Dim maxY As Long, minY As Long
    
    Dim tempArr As Variant
    
    'Get Upper and Lower Bounds
    maxX = UBound(MyArray, 1)
    minX = LBound(MyArray, 1)
    maxY = UBound(MyArray, 2)
    minY = LBound(MyArray, 2)
    
    'Create New Temp Array
    ReDim tempArr(minY To maxY, minX To maxX)
    
    'Transpose the Array
    For x = minX To maxX
        For Y = minY To maxY
            tempArr(Y, x) = MyArray(x, Y)
        Next Y
    Next x
    
    'Output Array
    TransposeArray = tempArr
    
End Function

Function isInitialised(ByRef a As Variant) As Boolean
    isInitialised = False
    On Error Resume Next
    isInitialised = IsNumeric(UBound(a))
End Function

Function IsInArr(str As Variant, arr As Variant) As Boolean
    Dim i As Long
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            If str = arr(i) Then
                IsInArr = True
                Exit Function
            End If
        Next i
    Else
        If str = arr Then
            IsInArr = True
            Exit Function
        End If
    End If
    IsInArr = False
End Function

Function InArr(str As Variant, arr As Variant) As Long
    Dim i As Long
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            If str = arr(i) Then
                InArr = i
                Exit Function
            End If
        Next i
    End If
    InArr = -1
End Function

Function Max(arr As Variant) As Double
    Dim i As Long
    Max = arr(LBound(arr))
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(arr(i)) And arr(i) > Max Then Max = arr(i)
    Next i
End Function

Function Min(arr As Variant) As Double
    Dim i As Long
    Min = arr(LBound(arr))
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(arr(i)) And arr(i) < Min Then Min = arr(i)
    Next i
End Function

Function ChangeArrBase(arr As Variant, base As Integer) As Variant
    Dim i As Long, tempArr As Variant
    Dim n As Long, count As Long
    n = UBound(arr) - LBound(arr) + 1
    ReDim tempArr(base To base + n - 1)
    For i = LBound(arr) To UBound(arr)
        tempArr(base + count) = arr(i)
        count = count + 1
    Next i
    ChangeArrBase = tempArr
End Function

Function ChangeObjArrBase(arr As Variant, base As Integer) As Variant

    Dim i As Long, tempArr() As Object
    Dim n As Long, count As Long
    n = UBound(arr) - LBound(arr) + 1
    ReDim tempArr(base To base + n - 1)
    For i = LBound(arr) To UBound(arr)
        Set tempArr(base + count) = arr(i)
        count = count + 1
    Next i
    ChangeObjArrBase = tempArr
End Function

Function CheckDuplicate(arr As Variant) As Boolean
    Dim i As Long, j As Long
    Dim count As Long
    
    For i = LBound(arr) To UBound(arr)
        count = 0
        For j = LBound(arr) To i
            If arr(i) = arr(j) Then
                count = count + 1
            End If
        Next j
        
        If count > 1 Then
            CheckDuplicate = True
            Exit Function
        End If
    Next i
    CheckDuplicate = False
End Function

Function DeDupeOneDimArray(vArray As Variant, Optional isExcludeEmpty As Boolean = True) As Variant
  Dim oDict As Object, i As Long
  Set oDict = CreateObject("Scripting.Dictionary")
  For i = LBound(vArray) To UBound(vArray)
    oDict(vArray(i)) = True
  Next i
  
  On Error Resume Next
  If isExcludeEmpty Then oDict.Remove ("")
  On Error GoTo 0
  
  DeDupeOneDimArray = oDict.keys()
End Function

Function DeDupeTwoDimArray(vArray As Variant, Optional isExcludeEmpty As Boolean = True, Optional excludeArr As Variant) As Variant
  Dim oDict As Object, i As Long, j As Long
  Set oDict = CreateObject("Scripting.Dictionary")
  
  For i = LBound(vArray, 1) To UBound(vArray, 1)
    For j = LBound(vArray, 2) To UBound(vArray, 2)
        oDict(vArray(i, j)) = True
    Next j
  Next i
  
  If isExcludeEmpty Then oDict.Remove ("")
  
  If isInitialised(excludeArr) Then
    For i = LBound(excludeArr) To UBound(excludeArr)
        oDict.Remove (excludeArr(i))
    Next i
  End If
  
  DeDupeTwoDimArray = oDict.keys()
End Function



Sub Filter(ByRef myWS As Worksheet, ByVal ref_firstRow As Integer, ByVal ref_firstCol As Integer, ByVal myCol As Integer, ByRef myArr As Variant)
    If .FilterMode Then
            .ShowAllData
    End If
        
    If myCol < ref_firstCol Then
        MsgBox "Error: checking filter option"
        End
    End If
    
    With myWS

        .Range(.Cells(ref_firstRow, ref_firstCol), .Cells(ref_firstRow, myCol)).AutoFilter , _
        Field:=myCol - ref_firstCol + 1, _
        Criteria1:=myArr, _
        Operator:=xlFilterValues
    End With
    
End Sub
'This sub find the unique data of certain column, and then save it to an array
Sub FindUnique_Visible(ByRef myArr As Variant, rng As Range)
    'sourceCell and targetCell are the reference cell
    
    Sheets.Add Before:=ActiveSheet

        rng.SpecialCells(xlCellTypeVisible).Copy
        Cells(1, 1).PasteSpecial xlPasteValues
        Range("A:A").RemoveDuplicates Columns:=1, Header:=xlNo
        myArr = SetArrayOneDim(1, 1, True)

    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End Sub
'Function GetListBoxArr(listBox As MSForms.listBox) As String()
'    Dim i As Long
'    Dim tempStr() As String
'
'    If listBox.ListCount > 0 Then
'        ReDim tempStr(listBox.ListCount - 1)
'    End If
'
'    For i = 0 To listBox.ListCount - 1
'      tempStr(i) = listBox.List(i)
'    Next i
'
'    GetListBoxArr = tempStr
'End Function


Function ArrListTo2DArr(arrList As Object) As Variant
    Dim arr_row As Variant
    Dim arr_2D As Variant
    Dim arr_list As Variant
    Dim i As Long, j As Long
    Dim is_ini As Boolean
    
    arr_list = arrList.ToArray
    is_ini = False
    
    For i = LBound(arr_list) To UBound(arr_list)
        arr_row = arr_list(i)
        If Not is_ini Then
    
            ReDim arr_2D(LBound(arr_list) To UBound(arr_list), LBound(arr_row) To UBound(arr_row))
            is_ini = True
        End If
        For j = LBound(arr_row) To UBound(arr_row)
            arr_2D(i, j) = arr_row(j)
        Next j
    Next i
    
    ArrListTo2DArr = arr_2D
End Function

'Function DictToArr(dict As Object) As Variant
'    Dim key As Variant
'    Dim tempArr As Variant
'    Dim count As Long
'
'    ReDim tempArr(dict.count - 1)
'    For Each key In dict.Keys
'        Set tempArr(count) = dict(key)
'    Next
'    DictToArr = tempArr
'End Function
Function Compare(ByVal value1 As Double, ByVal value2 As Double, is_ABS As Boolean) As Boolean 'return true if value2 is larger than value1

    If is_ABS = True Then 'turns all the number to positive if is_ABS is true
        
        If value1 < 0 Then
            value1 = -value1
        End If
        
        If value2 < 0 Then
           value2 = -value2
        End If
    End If
        
    If value1 < value2 Then
        Compare = True
        Exit Function
    Else
        Compare = False
        Exit Function
    End If
    
End Function

'Function ListBoxToArray(myListBox As MSForms.listBox) As Variant
'    Dim i As Long, myArr As Variant
'    ReDim myArr(myListBox.ListCount - 1)
'    For i = 0 To myListBox.ListCount - 1
'        myArr(i) = myListBox.Column(0, i)
'    Next i
'    ListBoxToArray = myArr
'End Function




'**************************************************************************************
'Below are some sub to expand the functionality of VBA Dictionary Class
Function AddItemsToArr(arr As Variant, val As Variant) As Variant
    Dim newArr As Variant
    Dim i As Long, n As Long
    Dim count As Long
    If IsArray(val) Then
        n = UBound(val) - LBound(val) + 1
    Else
        n = 1
    End If
    
    If IsArray(arr) Then
        ReDim newArr(LBound(arr) To UBound(arr) + n)
        For i = LBound(arr) To UBound(arr)
            newArr(i) = arr(i)
        Next i
    Else
        ReDim newArr(n)
        newArr(0) = arr
    End If
    
    'add new data to array
    If IsArray(val) Then
        count = LBound(val)
        For i = UBound(arr) + 1 To UBound(newArr)
            newArr(i) = val(count)
            count = count + 1
        Next i
    Else
        newArr(UBound(newArr)) = val
    End If
    
    AddItemsToArr = newArr
End Function

Function Combine2DArr(a As Variant, B As Variant, Optional stacked As Boolean = True) As Variant
    'assumes that A and B are 2-dimensional variant arrays
    'if stacked is true then A is placed on top of B
    'in this case the number of rows must be the same,
    'otherwise they are placed side by side A|B
    'in which case the number of columns are the same
    'LBound can be anything but is assumed to be
    'the same for A and B (in both dimensions)
    'False is returned if a clash

    Dim lb As Long, m_A As Long, n_A As Long
    Dim m_B As Long, n_B As Long
    Dim m As Long, n As Long
    Dim i As Long, j As Long, k As Long
    Dim c As Variant

    If TypeName(a) = "Range" Then a = a.Value
    If TypeName(B) = "Range" Then B = B.Value

    lb = LBound(a, 1)
    m_A = UBound(a, 1)
    n_A = UBound(a, 2)
    m_B = UBound(B, 1)
    n_B = UBound(B, 2)

    If stacked Then
        m = m_A + m_B + 1 - lb
        n = n_A
        If n_B <> n Then
            Combine = False
            Exit Function
        End If
    Else
        m = m_A
        If m_B <> m Then
            Combine = False
            Exit Function
        End If
        n = n_A + n_B + 1 - lb
    End If
    ReDim c(lb To m, lb To n)
    For i = lb To m
        For j = lb To n
            If stacked Then
                If i <= m_A Then
                    c(i, j) = a(i, j)
                Else
                    c(i, j) = B(lb + i - m_A - 1, j)
                End If
            Else
                If j <= n_A Then
                    c(i, j) = a(i, j)
                Else
                    c(i, j) = B(i, lb + j - n_A - 1)
                End If
            End If
        Next j
    Next i
    Combine2DArr = c
End Function


'Sub AddItemsToDict(dict As Object, key As Variant, val As Variant)
'    Dim newArr As Variant
'    Dim i As Long
'
'    If dict.Exists(key) Then
'        dict(key) = AddItemsToArr(dict(key), val)
'    Else
'        dict(key) = val
'    End If
'End Sub

Sub AddItemsToDict(dict As Object, key As Variant, val As Variant)
    Dim newArr As Variant
    Dim i As Long
    
    If dict.Exists(key) Then
        dict(key) = AddItemsToArr(dict(key), val)
    Else
        If Not IsObject(val) Then
            dict(key) = val
        Else
            Set dict(key) = val
        End If
    End If
End Sub

Function CreateDictFromObjs(objs As Variant, prop As String) As Object
    Dim i As Long, dict As Object
    Dim key As Variant
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = LBound(objs) To UBound(objs)
        key = CallByName(objs(i), prop, VbGet)
        AddItemsToDict dict, key, objs(i)
    Next i
    
    Set CreateDictFromObjs = dict
End Function

'Function DictToArr(dict As Object) As Variant
'    Dim key As Variant
'    For Each key In dict.keys
'        If IsArray(dict(key)) Then
'
'        Else
'
'        End If
'    Next key
'End Function

Sub AddArrToArrList(arrList As Object, arr As Variant)
    
    Dim i As Long
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            arrList.Add CVar(arr(i))
        Next i
    Else
        arrList.Add CVar(arr)
    End If
End Sub

Function getArrDimension(var As Variant) As Long
    On Error GoTo Err
    Dim i As Long
    Dim tmp As Long
    i = 0
    Do While True
        i = i + 1
        tmp = UBound(var, i)
    Loop
Err:
    getArrDimension = i - 1
End Function

Function Condense2DArrTo1D(var As Variant) As Variant
    Dim tempArr As Variant, n1 As Long, n2 As Long, n As Long
    Dim count As Long, i As Long, j As Long
    
    n1 = UBound(var, 1) - LBound(var, 1) + 1
    n2 = UBound(var, 2) - LBound(var, 2) + 1
    n = n1 * n2
    ReDim tempArr(n1 * n2 - 1)
    
    For i = LBound(var, 2) To UBound(var, 2)
        For j = LBound(var, 1) To UBound(var, 1)
            tempArr(count) = var(j, i)
            count = count + 1
        Next j
    Next i
    Condense2DArrTo1D = tempArr
End Function

Function Expand1DArrTo2D(var As Variant, Optional base As Long = 0) As Variant
    Dim tempArr As Variant, i As Long
    ReDim tempArr(LBound(var) To UBound(var), base To base)
    For i = LBound(var) To UBound(var)
        tempArr(i, base) = var(i)
    Next i
    Expand1DArrTo2D = tempArr
End Function

'nDim shall be 1 or 2, representing the dimension of the array
'n shall be the rowNum/ Column Num
Function Get1DArrFrom2DArr(var As Variant, nDim As Long, n As Long) As Variant
    Dim tempArr As Variant, i As Long
    Dim loopDim As Long
    If nDim = 1 Then
        loopDim = 2
    ElseIf nDim = 2 Then
        loopDim = 1
    Else
        Debug.Print "Function Get1DArrFrom2DArr: nDim shall be either 1 or 2"
    End If
    
    If getArrDimension(var) > 1 Then
        ReDim tempArr(LBound(var, loopDim) To UBound(var, loopDim))
        
        
            For j = LBound(tempArr) To UBound(tempArr)
                If nDim = 1 Then
                    tempArr(j) = var(n, j)
                Else
                    tempArr(j) = var(j, n)
                End If
            Next j
       
        
        Get1DArrFrom2DArr = tempArr
        
    ElseIf getArrDimension(var) = 1 Then
        Get1DArrFrom2DArr = var
        Exit Function
    End If
    
End Function

Function JoinStr(divider As String, ParamArray str() As Variant) As String
    Dim i As Long
    Dim isInitialised As Boolean
    
    isInitialised = False

    For i = LBound(str) To UBound(str)
        If isInitialised Then
            If Not CStr(str(i)) = "" Then JoinStr = JoinStr & divider & str(i)
        Else
            If Not CStr(str(i)) = "" Then
                JoinStr = str(i)
                isInitialised = True
            End If
        End If
    Next i
    
End Function

Function ArrToStr(var As Variant) As String
    Dim i As Long, myStr As String
    myStr = CStr(LBound(var))
    For i = LBound(var) + 1 To UBound(var)
        myStr = myStr & "," & var(i)
    Next i
End Function

Function StrToArrDbl(myStr As String) As Double()
    Dim str_splitted() As String
    str_splitted = Split(myStr, ",")
    CDbl_arr (str_splitted)
End Function

Function StrToArrLng(myStr As String) As Double()
    Dim str_splitted() As String
    str_splitted = Split(myStr, ",")
    CLng_arr (str_splitted)
End Function

Function CObj_arr(var As Variant) As Variant
    Dim arr() As Object, i As Long
    ReDim arr(LBound(var) To UBound(var))
    For i = LBound(var) To UBound(var)
        Set arr(i) = var(i)
    Next i
    CObj_arr = arr
End Function

Function CDbl_arr(var As Variant) As Double()
    Dim arr() As Double, i As Long
    ReDim arr(LBound(var) To UBound(var))
    For i = LBound(var) To UBound(var)
        arr(i) = CDbl(var(i))
    Next i
    CDbl_arr = arr
End Function

Function CStr_arr(var As Variant) As String()
    Dim arr() As String, i As Long
    ReDim arr(LBound(var) To UBound(var))
    For i = LBound(var) To UBound(var)
        arr(i) = CStr(var(i))
    Next i
    CStr_arr = arr
End Function

Function CLng_arr(var As Variant) As Long()
    Dim arr() As Long, i As Long
    ReDim arr(LBound(var) To UBound(var))
    For i = LBound(var) To UBound(var)
        arr(i) = CLng(var(i))
    Next i
    CLng_arr = arr
End Function
