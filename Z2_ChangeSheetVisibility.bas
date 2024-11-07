Attribute VB_Name = "Z2_ChangeSheetVisibility"
Sub ManageSheetVisibilityInteractive()

    ' Kinen Ma
    ' 2024-11-07
    ' Version 1.0
    ' Interactivelly change sheet visibility, including very hidden sheet

    Dim ws As Worksheet
    Dim sheetList As String
    Dim sheetIndexInput As String
    Dim visibilityChoice As String
    Dim visibility As XlSheetVisibility
    Dim sheetIndexes As Variant
    Dim i As Integer
    Dim currentVisibility As String
    Dim maxIndexLength As Integer
    Dim maxNameLength As Integer

    ' Determine max lengths for index and sheet name
    For Each ws In ThisWorkbook.Worksheets
        If Len(ws.Name) > maxNameLength Then maxNameLength = Len(ws.Name)
    Next ws
    maxIndexLength = Len(CStr(ThisWorkbook.Worksheets.count))
    
    ' Create header for the message box
    sheetList = "Index" & space(Application.Max(0, maxIndexLength + 2 - Len("Index"))) & _
                "Sheet Name" & space(Application.Max(0, maxNameLength + 2 - Len("Sheet Name"))) & _
                "Visibility" & vbCrLf
    sheetList = sheetList & String(maxIndexLength + maxNameLength + 20, "-") & vbCrLf

    ' Enumerate sheets with visibility status
    For i = 1 To ThisWorkbook.Worksheets.count
        Set ws = ThisWorkbook.Worksheets(i)
        Select Case ws.Visible
            Case xlSheetVisible
                currentVisibility = "Visible"
            Case xlSheetHidden
                currentVisibility = "Hidden"
            Case xlSheetVeryHidden
                currentVisibility = "Very Hidden"
        End Select
        
        ' Append sheet details to the list
        sheetList = sheetList & i & space(Application.Max(0, maxIndexLength + 2 - Len(CStr(i)))) & _
                    ws.Name & space(Application.Max(0, maxNameLength + 2 - Len(ws.Name))) & _
                    currentVisibility & vbCrLf
    Next i
    
    ' Combine sheet list display and sheet selection
    sheetIndexInput = inputBox("Sheets in this workbook:" & vbCrLf & sheetList & vbCrLf & _
                               "Enter the number(s) of the sheets you want to modify, separated by commas:", "Select Sheet Numbers")
    
    If sheetIndexInput = "" Then
        MsgBox "No sheets selected.", vbExclamation
        Exit Sub
    End If
    
    ' Split the input into an array of sheet indexes
    sheetIndexes = Split(sheetIndexInput, ",")
    
    ' Get the visibility choice from the user
    visibilityChoice = inputBox("Enter the visibility option for the selected sheets:" & vbCrLf & "1. Visible" & vbCrLf & "2. Hidden" & vbCrLf & "3. Very Hidden", "Set Visibility")
    
    Select Case visibilityChoice
        Case "1"
            visibility = xlSheetVisible
        Case "2"
            visibility = xlSheetHidden
        Case "3"
            visibility = xlSheetVeryHidden
        Case Else
            MsgBox "Invalid choice.", vbExclamation
            Exit Sub
    End Select
    
    ' Change the visibility of the selected sheets
    For i = LBound(sheetIndexes) To UBound(sheetIndexes)
        Dim sheetIndex As Integer
        sheetIndex = val(Trim(sheetIndexes(i)))
        
        If sheetIndex > 0 And sheetIndex <= ThisWorkbook.Worksheets.count Then
            Set ws = ThisWorkbook.Worksheets(sheetIndex)
            ws.Visible = visibility
        Else
            MsgBox "Invalid sheet number: " & sheetIndexes(i), vbExclamation
        End If
    Next i
    
    MsgBox "The visibility of the selected sheets has been changed.", vbInformation
End Sub

