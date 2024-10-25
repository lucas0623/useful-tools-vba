VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUIManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'This class contains Sub and Funcitons for Creating Userforms/ other UIs, Getting User Input and Loading save input
'Arthor: Lucas LEUNG
'Update Log:
'21 Nov 2022 - ver1.0

'***************************************************************
'Get User Input Through Userform
Function AddInputBox_rng(prompt As String, Optional title As String, Optional default As Range) As Range
    'Return "Nothing" if "cancel" is clicked
    If default Is Nothing Then
        Set default = Range("D1")
    End If
    On Error GoTo ErrHandler:
        Set AddInputBox_rng = Application.inputBox(prompt:=prompt, title:=title, default:=default.Address, Type:=8)
    On Error GoTo 0
    
    Exit Function
ErrHandler:
    Set AddInputBox_rng = Nothing

End Function
Function AddInputBox_str(prompt As String, Optional title As String, Optional defaultStr As String = "") As String
    
    'Return "False" if "cancel" is clicked
    AddInputBox_str = Application.inputBox(prompt:=prompt, title:=title, default:=defaultStr, Type:=2)
    
End Function

Function GetFilePath(fileType As String, Optional myTitle As String = "File Path") As String

    Dim strFileToOpen As String
    
    strFileToOpen = Application.GetOpenFilename _
    (title:=myTitle, _
    FileFilter:="File Type *" & fileType & "* (*" & fileType & "*),")
    
    If strFileToOpen = "False" Then
        'MsgBox "No file selected.", vbExclamation, "Sorry!"
        Application.ScreenUpdating = True
        Application.Calculation = xlAutomatic
        End
    Else
        GetFilePath = strFileToOpen
    End If
    
End Function

Function GetFolderPath() As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ActiveWorkbook.Path
        .title = "Please choose a folder"
        .AllowMultiSelect = False
    
        If .Show = -1 Then
            GetFolderPath = .SelectedItems(1) '=-1 if action button is clicked, =0 if cancel button is clicked
        End If
    End With

End Function
