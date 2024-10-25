VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStringSlicer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Last Update: 18 Aug 2022

'************************************************************************
'this function will return the array of number represented by the input string
'Read String and turn to Array, suitable for
'1. MIDAS Element List, e.g. "36225to36256 36654to36733 37382to37404 37619to37697"
Function Slicer(str As String, Optional delimiter_to As String = "to", _
                Optional delimiter_and As String = " ") As Variant

    'Dim isReturn() As Boolean
    Dim arr As Variant, arr2 As Variant
    Dim returnArr As Variant, arrList As Object
    Dim i As Long, j As Long, count As Long
    
    Set arrList = CreateObject("System.Collections.ArrayList")
    
'    ReDim isReturn(lLimit To uLimit)
'
'    For i = lLimit To uLimit
'        isReturn(i) = False
'    Next i
'    count = 0
    
    arr = Split(str, delimiter_and)
    
    For i = 0 To UBound(arr)
        If InStr(1, arr(i), delimiter_to, vbTextCompare) Then
            arr2 = Split(arr(i), delimiter_to)
                For j = CLng(arr2(0)) To CLng(arr2(1))
                    arrList.Add j
                Next j
        Else
            arrList.Add i
        End If
    
    Next i
    
    'assign the true into a single array
'    For i = lLimit To uLimit
'        If isReturn(i) Then
'            count = count + 1
'        End If
'    Next i
'
'    ReDim returnArr(count - 1)
'    count = 0
'    For i = lLimit To uLimit
'        If isReturn(i) Then
'            returnArr(count) = i
'            count = count + 1
'        End If
'    Next i
    
    Slicer = arrList.ToArray
End Function

'Used in dataframe. Reference is made with Python Pandas
Function Slicer_wLimits(lLimit As Long, uLimit As Long, str As String) As Variant

    Dim isReturn() As Boolean
    Dim arr As Variant, arr2 As Variant
    Dim returnArr As Variant
    Dim i As Long, j As Long, count As Long
    
    ReDim isReturn(lLimit To uLimit)
    
    For i = lLimit To uLimit
        isReturn(i) = False
    Next i
    count = 0
    
    arr = Split(str, ",")
    
    For i = 0 To UBound(arr)
        If InStr(1, arr(i), ":", vbTextCompare) Then
            arr2 = Split(arr(i), ":")
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
