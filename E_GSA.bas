Attribute VB_Name = "E_GSA"
Private UIManager As clsUIManager
Private genFunc As clsGeneralFunctions
Private wsInteract As clsWorksheetsInteraction

Sub Init()
    Set genFunc = New clsGeneralFunctions
    Set wsInteract = New clsWorksheetsInteraction
    Set UIManager = New clsUIManager
End Sub

Sub CreateGSAListFromSectionProp()
    Dim rng_eleID As Range, rng_section As Range, rng_output As Range
    Dim i As Long, key As Variant, tempStr As String, count As Long
    'Dim ws_output As Worksheet, fRow_output As Long, fCol_output As Long
    Dim dict_GSAList As Object, arrL_output As Object
    
    Init
    
    Set dict_GSAList = CreateObject("Scripting.Dictionary")
    Set arrL_output = CreateObject("System.Collections.ArrayList")
    Set rng_eleID = UIManager.AddInputBox_rng("Select the range for element ID")
    Set rng_section = UIManager.AddInputBox_rng("Select the range for element section")
    Set rng_output = UIManager.AddInputBox_rng("Select the range for outputing the GSA list")
    
    For i = 1 To rng_section.Cells.count
        genFunc.AddItemsToDict dict_GSAList, rng_section.Cells(i).text, rng_eleID.Cells(i).text
    Next
    
    count = 0
    
    rng_output.Value = "List Name"
    rng_output.Worksheet.Cells(rng_output.row, rng_output.Column + 1) = "Type"
    rng_output.Worksheet.Cells(rng_output.row, rng_output.Column + 2) = "Definition"
    For Each key In dict_GSAList.keys
        tempStr = ""
        If genFunc.isInitialised(dict_GSAList(key)) Then
            For i = LBound(dict_GSAList(key)) To UBound(dict_GSAList(key))
                tempStr = genFunc.JoinStr(" ", tempStr, CStr(dict_GSAList(key)(i)))
            Next i
        Else
            tempStr = dict_GSAList(key)
        End If
        arrL_output.Add Array(key, "Element", tempStr)
    Next key
    'rng_output = arrL_output.ToArray
    wsInteract.WriteArrToRngTwoDim genFunc.ArrListTo2DArr(arrL_output), rng_output.row + 1, _
                                rng_output.Column, rng_output.Worksheet
    'rng_output.Resize(1, arrL_output.count).Value = arrL_output.ToArray
    
End Sub

Sub BreakDownListInTable()
    Init
    
    Dim ws_main As Worksheet, ws_output As Worksheet
    Dim rng_rowTag As Range, rng_colTag As Range
    Dim rng_output As Range
    Dim i As Long, var As Variant ', arrL As Object
    Dim df As clsDataFrame, df2 As clsDataFrame
    Dim col_eleList As Long
    
    Set rng_rowTag = UIManager.AddInputBox_rng("Select the COLUMN that contains 'reference row tags'", default:=Columns(1))
    If rng_rowTag Is Nothing Then EndSub
    
    Set rng_colTag = UIManager.AddInputBox_rng("Select ROWS that contains 'heads tags'", default:=Rows(1))
    If rng_colTag Is Nothing Then EndSub
    
    Set rng_list = UIManager.AddInputBox_rng("Select the heads tags of the table that represents the columns of the element/node list", default:=Range("D1:H1"))
    If rng_list Is Nothing Then EndSub
    
    Set rng_output = UIManager.AddInputBox_rng("Select the target output location", default:=Range("A1"))
    If rng_output Is Nothing Then EndSub
    
    Set ws_main = rng_rowTag.Worksheet
    Set ws_output = rng_output.Worksheet
    
    Set df = New clsDataFrame
    Set df2 = New clsDataFrame
    df.Init_ReadWorksheet
    col_eleList = df.columnNum(rng_list.text)
    For i = 1 To df.CountRows
        var = BreakList(df.idata(i, col_eleList), " ", "to")
        df2.AddDataFrame df.MultiplyRow(i, rng_list.text, var)
    Next i
    
    df2.WriteToRange ws_output, tarRng:=rng_output, isWriteRowTag:=False
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
            arr3 = genFunc.CStr_arr(FillNumRange(CLng(arr2(0)), CLng(arr2(1))))
            genFunc.AddArrToArrList arrL, arr3
        
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

Private Sub EndSub()
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    End
End Sub
