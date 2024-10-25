Attribute VB_Name = "ExportProject"
'@Folder("VBAProject")
Sub ExportAllVBAComponents()
    Dim vbComp As Object
    Dim vbProj As Object
    Dim exportPath As String
    Dim fileName As String
    
    ' Set the export path
    exportPath = "C:\Users\lucasleung\OneDrive\12 Engineering\01 Resource Lib\04 Useful Excel Macro\Tools Addin\Useful Tools VBA Modules" ' Change this to your desired folder path
    
    ' Ensure the path ends with a backslash
    If Right(exportPath, 1) <> "\" Then exportPath = exportPath & "\"
    
    ' Loop through each component in the VBA project
    Set vbProj = ThisWorkbook.VBProject
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
