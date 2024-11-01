Attribute VB_Name = "ExportProject"
'@Folder("VBAProject")

Private UI As New clsUIManager

Sub ExportAllVBAComponents()
    Dim vbComp As Object
    Dim vbProj As Object
    Dim exportPath As String
    Dim fileName As String
    Dim dict As Object
    Dim selectedName As String
    Dim selectedItem As Object
    
    ' Get the dictionary of open workbooks and add-ins
    Set dict = GetOpenWorkbooksAndAddins()
    
    ' Get the names of all open workbooks and add-ins
    Dim names() As String
    names = GetDictionaryKeys(dict)
    
    ' Show the combo box to select a workbook or add-in
    selectedName = UFShowComboBox(names)
    If selectedName = "" Then Exit Sub ' User cancelled the selection
    
    ' Get the selected workbook or add-in object
    Set selectedItem = dict(selectedName)
    
    ' Set the export path
    exportPath = UI.GetFolderPath
    
    ' Ensure the path ends with a backslash
    If Right(exportPath, 1) <> "\" Then exportPath = exportPath & "\"
    
    ' Loop through each component in the selected VBA project
    If TypeName(selectedItem) = "Workbook" Then
        Set vbProj = selectedItem.VBProject
        
    ElseIf TypeName(selectedItem) = "AddIn" Then
        Set vbProj = Workbooks(selectedItem.Name).VBProject
    Else
        MsgBox "Selected item is not a valid workbook or add-in."
        Exit Sub
    End If
    
    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case 1 ' Standard Module
                fileName = exportPath & vbComp.Name & ".bas"
                vbComp.Export fileName
            Case 2 ' Class Module
                fileName = exportPath & vbComp.Name & ".cls"
                vbComp.Export fileName
            Case 3 ' UserForm
                fileName = exportPath & vbComp.Name & ".frm"
                vbComp.Export fileName
            Case Else
                ' Handle other types if necessary
        End Select
    Next vbComp

    MsgBox "Export completed!"
End Sub

Private Function GetDictionaryKeys(dict As Object) As String()
    Dim names() As String
    Dim i As Integer
    Dim key As Variant
    
    ReDim names(dict.count - 1)
    i = 0
    For Each key In dict.Keys
        names(i) = key
        i = i + 1
    Next key
    
    GetDictionaryKeys = names
End Function

Private Function UFShowComboBox(comboBoxArr As Variant) As String
    Dim uf As UFBasic
    Dim comboBox As msforms.comboBox
    
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

Private Function GetOpenWorkbooksAndAddins() As Object
    Dim wb As Workbook
    Dim ai As AddIn
    Dim dict As Object
    
    ' Create a new dictionary using late binding
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Loop through all open workbooks and add them to the dictionary
    For Each wb In Application.Workbooks
        dict.Add wb.Name, wb
    Next wb
    
    ' Loop through all add-ins and add them to the dictionary if they are installed
    For Each ai In Application.AddIns
        If ai.Installed Then
            dict.Add ai.Name, ai
        End If
    Next ai
    
    ' Return the dictionary
    Set GetOpenWorkbooksAndAddins = dict
End Function
