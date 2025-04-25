Attribute VB_Name = "F_CopyDefinedNames"
'Sub Test()
'    Dim arr As Variant
'    Debug.Print Workbooks.count
'    arr = GetAllSheetsNamesInWB(ActiveWorkbook)
'
'    For i = 0 To UBound(arr)
'        Debug.Print arr(i)
'    Next i
'End Sub
'Sub TestName()
'    Dim name1 As Name, name2 As Name, tempName As Name
'    Set name1 = ActiveWorkbook.Names.Add("ZZZ", "AAA")
'    Set tempName = ActiveWorkbook.Names.Add("ZZZ_", "AAA")
'    name1.Delete
'    Set name2 = ActiveSheet.Names.Add("ZZZ", "AAA")
'    tempName.Delete
'End Sub
Sub WriteDefinedNamesToCell()
    Dim oWS As Worksheet, refRng As Range
    Dim rRow As Long, rCol As Long
    Dim nameList() As String, selectedNames() As String
    
    Set oWS = ActiveSheet
    nameList = GetAllNamesInSheet(oWS)
    If Not isInitialised(nameList) Then
        MsgBox "No Defined Name is found in the Active Sheet"
        Exit Sub
    End If
    
    selectedNames = UFShowArrInList(nameList, "Select Defined Names")
    If Not isInitialised(selectedNames) Then Exit Sub
    
    Set refRng = AddInputBox_rng("Write Names Info at:")
    If refRng Is Nothing Then Exit Sub
    rRow = refRng.row
    rCol = refRng.Column
    Dim i As Long, cName As name, targetWS As Worksheet
    
    For i = 0 To UBound(selectedNames)
        'Debug.Print selectedNames(i)
        Set cName = GetNameObjByString(selectedNames(i), oWS)
        WriteNameToSheet cName, rRow + i, rCol
    Next i
    
End Sub

Sub CopyDefinedNamesToSheet()
    Dim oWS As Worksheet
    Dim nameList() As String, selectedNames() As String
    Dim WBs() As String, selectedWB As String
    Dim WSs() As String, selcetedWS() As String
    
    ' Set the source worksheet to the active sheet
    Set oWS = ActiveSheet
    
    ' Get all defined names in the active sheet
    nameList = GetAllNamesInSheet(oWS)
    If Not isInitialised(nameList) Then
        MsgBox "No Defined Name is found in the Active Sheet"
        Exit Sub
    End If
    
    ' Let user select which names to copy
    selectedNames = UFShowArrInList(nameList, "Select Defined Names")
    If Not isInitialised(selectedNames) Then Exit Sub
    
    ' Let user select the target workbook
    WBs = GetAllOpenedWBNames
    selectedWB = UFShowComboBox(WBs)
    If selectedWB = vbNullString Then Exit Sub
    
    ' Prompt user to choose between specific sheet or global scope
    Dim copyOption As VbMsgBoxResult
    copyOption = MsgBox("Do you want to copy the defined names to a specific sheet or to the global scope?" & vbCrLf & _
                        "Click Yes for specific sheet, No for global scope, Cancel to abort.", vbYesNoCancel)
    If copyOption = vbCancel Then Exit Sub
    
    Dim i As Long, cName As name, targetWS As Worksheet, targetWB As Workbook
    Set targetWB = Workbooks(selectedWB)
    
    If copyOption = vbYes Then
        ' Option 1: Copy to a specific sheet
        WSs = GetAllSheetsNamesInWB(targetWB)
        selcetedWS = UFShowArrInList(WSs, "Select Target Worksheet", False)
        If selcetedWS(0) = vbNullString Then Exit Sub
        Set targetWS = targetWB.Worksheets(selcetedWS(0))
        
        ' Copy each selected name to the target sheet
        For i = 0 To UBound(selectedNames)
            Set cName = GetNameObjByString(selectedNames(i), oWS)
            PasteNameToSheet cName, targetWS, nameList
        Next i
    ElseIf copyOption = vbNo Then
        ' Option 2: Copy to global scope
        For i = 0 To UBound(selectedNames)
            Set cName = GetNameObjByString(selectedNames(i), oWS)
            Dim nameStr As String
            ' Extract the name part without the sheet prefix
            nameStr = Split(cName.name, "!")(1)
            Dim refersTo As String
            refersTo = cName.refersTo
            
            ' Attempt to add the name to the workbook's global scope
            On Error Resume Next
            Dim newName As name
            Set newName = targetWB.Names.Add(name:=nameStr, refersTo:=refersTo)
            If Err.Number = 0 Then
                ' If successful, copy the comment if it exists
                If Not cName.Comment = "" Then
                    newName.Comment = cName.Comment
                End If
            Else
                ' Inform user if the name couldn’t be added (e.g., duplicate name)
                MsgBox "Failed to add name '" & nameStr & "' to the global scope. It may already exist."
                Err.Clear
            End If
            On Error GoTo 0
        Next i
    End If
    
    MsgBox "Selected Names Copied."
End Sub

Sub ChangeDefinedNamesScopeLocalToGlobal()
    Dim oWS As Worksheet, oWB As Workbook
    Dim nameList() As String, selectedNames() As String
    'Dim WBs() As String, selectedWB As String
    
    Set oWS = ActiveSheet
    Set oWB = ActiveWorkbook
    nameList = GetAllNamesInSheet(oWS)
    If Not isInitialised(nameList) Then
        MsgBox "No Defined Name is found in the Active Sheet"
        Exit Sub
    End If
    
    selectedNames = UFShowArrInList(nameList, "Select Defined Names")
    If Not isInitialised(selectedNames) Then Exit Sub
    
'    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False

    Dim i As Long, cName As name, newName As name, tempName As name
    Dim oldNameStr As String, tempNameStr As String, tempNameStrFull As String, newNameStrFull As String
    Dim newNameStr As String, newRefersTo As String, newComment As String
    Dim x As name
    For i = 0 To UBound(selectedNames)
        'Debug.Print selectedNames(i)
        Set cName = GetNameObjByString(selectedNames(i), oWS)
        oldNameStr = cName.name
        newNameStr = Split(cName.name, "!")(1)
        tempNameStr = newNameStr & "_"
        newNameStrFull = "'" & oWB.name & "'!" & newNameStr
        tempNameStrFull = "'" & oWB.name & "'!" & tempNameStr
        newRefersTo = cName.refersTo
        newComment = cName.Comment
        
        Set tempName = oWB.Names.Add(tempNameStr, newRefersTo)
        UpdateChartSeriesReference oWS, oldNameStr, tempNameStrFull

        cName.Delete
        Set newName = oWB.Names.Add(newNameStr, newRefersTo)
        newName.Comment = newComment
        UpdateChartSeriesReference oWS, tempNameStrFull, newNameStrFull
        
        
        'Find all names that refers to the name and change their formula
        For Each x In oWB.Names
            If Not InStr(1, x.refersTo, oldNameStr) = 0 Then
                x.refersTo = Replace(x.refersTo, oldNameStr, newNameStrFull)
            End If
        Next
        tempName.Delete
    Next i
    
    '
'    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True

    MsgBox "Scope Changed."
End Sub

Sub ChangeDefinedNamesScopeGlobalToLocal()
    Dim oWS As Worksheet, oWB As Workbook
    Dim nameList() As String, selectedNames() As String
    'Dim WBs() As String, selectedWB As String
    
    Set oWS = ActiveSheet
    Set oWB = ActiveWorkbook
    nameList = GetAllNamesInSheet(oWB)
    If Not isInitialised(nameList) Then
        MsgBox "No Defined Name is found in the Active Sheet"
        Exit Sub
    End If
    
    selectedNames = UFShowArrInList(nameList, "Select Defined Names")
    If Not isInitialised(selectedNames) Then Exit Sub
    
'    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False

    Dim i As Long, cName As name, newName As name, tempName As name
    Dim oldNameStr As String, tempNameStr As String
    Dim newNameStr As String, newRefersTo As String, newComment As String
    
    For i = 0 To UBound(selectedNames)
        'Debug.Print selectedNames(i)
        
        Set cName = GetNameObjByString(selectedNames(i), oWB)
        oldNameStr = "'" & oWB.name & "'!" & cName.name
        
        If Not InStr(1, oWS.name, " ") = 0 Then
            newNameStr = "'" & oWS.name & "'!" & cName.name
            tempNameStr = "'" & oWS.name & "'!" & cName.name & "_"
        Else
            newNameStr = oWS.name & "!" & cName.name
            tempNameStr = oWS.name & "!" & cName.name & "_"
        End If
        
        newRefersTo = cName.refersTo
        newComment = cName.Comment
        
        Set tempName = oWS.Names.Add(newNameStr & "_", newRefersTo)
        UpdateChartSeriesReference oWS, oldNameStr, tempNameStr
        
        cName.Delete
        Set newName = oWS.Names.Add(newNameStr, newRefersTo)
        newName.Comment = newComment
        UpdateChartSeriesReference oWS, tempNameStr, newNameStr
        
        'Find all names that refers to the name and change their formula
        For Each x In oWB.Names
            If Not InStr(1, x.refersTo, oldNameStr) = 0 Then
'                Do Until InStr(1, x.RefersTo, oldNameStr) = 0
'                    x.RefersTo = Replace(x.RefersTo, oWS.Name & "!", oWB.Name & "!")
'                Loop
                
                'Do Until InStr(1, x.RefersTo, oldNameStr) = 0
                    x.refersTo = Replace(x.refersTo, oldNameStr, newNameStr)
                'Loop
            End If
        Next
        tempName.Delete
    Next i

'    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True

    MsgBox "Scope Changed."
End Sub

Private Function GetAllOpenedWBNames() As String()
    Dim wb As Workbook
    Dim wbNameColl As New Collection, wbNameArr() As String
    
    For Each wb In Workbooks
        'If Not UCase(Split(wb.Name, ".")(1)) = "XLSB" Then
            wbNameColl.Add wb.name
        'End If
    Next
    wbNameArr = CStr_arr(CollToArr1D(wbNameColl))
    GetAllOpenedWBNames = wbNameArr
End Function

Private Function GetAllSheetsNamesInWB(wb As Workbook) As String()
    Dim sheetsName() As String
    Dim cSheet As Worksheet, count As Integer
    
    ReDim sheetsName(wb.Sheets.count - 1)
    
    For Each cSheet In wb.Sheets
        sheetsName(count) = cSheet.name
        count = count + 1
    Next
    
    GetAllSheetsNamesInWB = sheetsName
End Function

Private Function GetAllNamesInSheet(obj As Object) As String()
    Dim nameList As New Collection
    Dim x As name
    If TypeOf obj Is Worksheet Then
        For Each x In obj.Names
            If InStr(1, x.name, obj.name) > 0 Then
                'Debug.Print Split(x.name, "!")(1)
                nameList.Add Split(x.name, "!")(1)
            End If
        Next x
    ElseIf TypeOf obj Is Workbook Then
        For Each x In obj.Names
            If InStr(1, x.name, ".") = 0 And InStr(1, x.name, "!") = 0 Then
                'Debug.Print Split(x.name, "!")(1)
                nameList.Add x.name
            End If
        Next x
    Else
        MsgBox "Wrong Input Object Type"
    End If
    If nameList.count = 0 Then Exit Function
    GetAllNamesInSheet = CStr_arr(CollToArr1D(nameList))
End Function

Private Function GetNameObjByString(str As String, obj As Object) As name
    Dim retName As name
    Dim x As name
    If TypeOf obj Is Worksheet Then
        For Each x In obj.Names
            If InStr(1, x.name, obj.name) > 0 Then
                'Debug.Print Split(x.name, "!")(1)
                If InStr(1, x.name, str) > 0 Then
                    Set retName = x
                    Exit For
                End If
            End If
        Next x
    ElseIf TypeOf obj Is Workbook Then
        For Each x In obj.Names
            If x.name = str Then
                Set retName = x
                Exit For
            End If
        Next x
    Else
        MsgBox "Wrong Input Object Type"
    End If
    Set GetNameObjByString = retName
End Function

Private Function UFShowArrInList(arr As Variant, listboxTitle As String, Optional isMultiSelect As Boolean = True) As String()
'UF to get user input
    Dim uf As UFBasic
    Dim listbox1 As MSForms.listbox
    
    Set uf = New UFBasic
    uf.Initialize 200, True
    uf.TitleBarCaption = "Copy Defined Names"
    
    uf.AddSelectionBox_Empty listbox1, listboxTitle
    listbox1.List = arr
    If Not isMultiSelect Then listbox1.MultiSelect = fmMultiSelectSingle
    
    uf.AdjustHeight
    uf.RepositionOkAndCancelButtonsToCenter
    uf.Show
    
    If Not uf.CloseState = 0 Then Exit Function
    
    Dim retStr() As String
    retStr = ListBoxSelectedToArray(listbox1)
    UFShowArrInList = retStr
End Function

Private Sub UpdateChartSeriesReference(sht As Worksheet, oldStr As String, newStr As String)
    
    Dim chtObj As ChartObject, cht As Chart
    Dim srsColl As SeriesCollection, srs As Series
    Set sht = ActiveSheet
    For Each chtObj In sht.ChartObjects
        Debug.Print (chtObj.name)
        Set cht = chtObj.Chart
        Set srsColl = cht.SeriesCollection
        For Each srs In srsColl
            Debug.Print srs.formula
            srs.formula = Replace(srs.formula, oldStr, newStr)
            Debug.Print srs.formula
        'Debug.Print srs.YValues
        Next
    Next chtObj
    Debug.Print ("Complete")
End Sub
'Private Function UFShowComboBoxWithArrInList(comboBoxArr As Variant) As String
'    Dim uf As UFBasic
'    Dim comboBox As MSForms.comboBox, listbox As MSForms.listbox
'
'    Set uf = New UFBasic
'    uf.Initialize 350, True
'    uf.TitleBarCaption = "Copy Defined Names"
'
'    uf.AddComboBox comboBoxArr, comboBox, "Select Workbook"
'    uf.AddSelectionBox_Empty listbox, "Select Target Worksheet"
'    listbox.List = GetAllSheetsNamesInWB(Workbooks(comboBox.Value))
'    listbox.MultiSelect = fmMultiSelectSingle
'
'    uf.AdjustHeight
'    uf.RepositionOkAndCancelButtonsToCenter
'    uf.Show
'
'    If Not uf.CloseState = 0 Then Exit Function
'
'    Dim retStr As String
'    retStr = listbox1.Value
'    UFShowComboBoxWithArrInList = retStr
'End Function

Private Function UFShowComboBox(comboBoxArr As Variant) As String
    Dim uf As UFBasic
    Dim comboBox As MSForms.comboBox
    
    Set uf = New UFBasic
    uf.Initialize 400, True
    uf.TitleBarCaption = "Copy Defined Names"
    
    uf.AddComboBox comboBoxArr, comboBox, "Select Workbook"
        
    uf.AdjustHeight
    uf.RepositionOkAndCancelButtonsToCenter
    uf.Show
    
    If Not uf.CloseState = 0 Then Exit Function
    
    Dim retStr As String
    retStr = comboBox.Value
    UFShowComboBox = retStr
End Function
Private Sub PasteNameToSheet(oName As name, targetSheet As Object, nameList() As String)
    Dim newName As name
    Dim newNameStr As String, newRefersTo As String
    newNameStr = Split(oName.name, "!")(1)
    newRefersTo = oName.refersTo
    
    Do Until InStr(1, newRefersTo, oName.Parent.name & "!") = 0
        newRefersTo = Replace(newRefersTo, oName.Parent.name & "!", targetSheet.name & "!")
    Loop
    
    Do Until InStr(1, newRefersTo, "'" & oName.Parent.name & "'!") = 0
        newRefersTo = Replace(newRefersTo, "'" & oName.Parent.name & "'!", "'" & targetSheet.name & "'!")
    Loop
    Set newName = targetSheet.Names.Add(name:=newNameStr, refersTo:=newRefersTo)
    newName.Comment = oName.Comment
    
    'check if the name refers to another defined name. If yes, paste that name once more
    Dim i As Long
    For i = LBound(nameList) To UBound(nameList)
        If InStr(1, newRefersTo, nameList(i)) > 0 And Not nameList(i) = newNameStr Then
            PasteNameToSheet GetNameObjByString(nameList(i), ActiveSheet), targetSheet, nameList
        End If
    Next i
End Sub

Private Sub WriteNameToSheet(oName As name, row As Long, col As Long)
    'Dim newName As Name
    'Set newName = targetSheet.Names.Add(Name:=Split(oName.Name, "!")(1), RefersTo:=oName.RefersTo)
    Cells(row, col) = Split(oName.name, "!")(1)
    Cells(row, col + 2) = "'" & oName.refersTo
    Cells(row, col + 1) = oName.Comment
End Sub

'**************************************************************************************************************************************
Private Function CollToArr1D(coll As Collection) As Variant
    Dim i As Long, arr As Variant
    ReDim arr(coll.count - 1)
    For i = 1 To coll.count
        arr(i - 1) = coll(i)
    Next i
    CollToArr1D = arr
End Function

Private Function CStr_arr(var As Variant) As String()
    Dim arr() As String, i As Long
    If IsArray(var) Then
        ReDim arr(LBound(var) To UBound(var))
        For i = LBound(var) To UBound(var)
            arr(i) = CStr(var(i))
        Next i
    Else
        ReDim arr(0)
        arr(0) = CStr(var)
    End If
    CStr_arr = arr
End Function

Private Function isInitialised(ByRef a As Variant) As Boolean
'This sub check if an ARRAY is initialized.
    isInitialised = False
    On Error GoTo ErrHandler
    If IsArray(a) Then
        If Not UBound(a) = -1 Then
            isInitialised = True
        End If
    ElseIf Not a = vbNullString Then
        isInitialised = True
    End If

    Exit Function
ErrHandler:
    isInitialised = False
End Function

Private Function ListBoxSelectedToArray(myListBox As MSForms.listbox) As String()
    Dim i As Long, myArr As Variant, myColl As New Collection
    If Not myListBox.ListCount = 0 Then
        ReDim myArr(myListBox.ListCount - 1)
        For i = 0 To myListBox.ListCount - 1
            If myListBox.Selected(i) Then
                myColl.Add myListBox.Column(0, i)
            End If
        Next i
        myArr = CollToArr1D(myColl)
        ListBoxSelectedToArray = CStr_arr(myArr)
    End If
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

Private Function IsInArr(str As Variant, arr As Variant, Optional isCaseSensitive As Boolean = True) As Boolean
    Dim i As Long
    If Not isCaseSensitive Then str = UCase(str)
    
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            If Not isCaseSensitive Then arr(i) = UCase(arr(i))
            If str = arr(i) Then
                IsInArr = True
                Exit Function
            End If
        Next i
    Else
        If Not isCaseSensitive Then arr = UCase(arr)
        If str = arr Then
            IsInArr = True
            Exit Function
        End If
    End If
    IsInArr = False
End Function


