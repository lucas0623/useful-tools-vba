Attribute VB_Name = "D_FileManager"
Private UI As clsUIManager
Private genFunc As clsGeneralFunctions
Private wsInteract As clsWorksheetsInteraction

Sub GetAllFilesInFolder()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Set UI = New clsUIManager
    Set genFunc = New clsGeneralFunctions
    Set wsInteract = New clsWorksheetsInteraction
    
    Dim oFSO As Object
    Dim folderPath As String
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    Dim rng_output As Range
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
     
     folderPath = UI.GetFolderPath
    If folderPath = vbNullString Then Exit Sub
    Set oFolder = oFSO.GetFolder(folderPath)
    
    Set rng_output = UI.AddInputBox_rng("Select the location for output")
    If rng_output Is Nothing Then Exit Sub
    
    For Each oFile In oFolder.Files
     
        rng_output.Worksheet.Cells(rng_output.row + i, rng_output.Column) = oFile.Name
        rng_output.Worksheet.Cells(rng_output.row + i, rng_output.Column + 1) = oFile.Type
        i = i + 1
     
    Next oFile
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub

