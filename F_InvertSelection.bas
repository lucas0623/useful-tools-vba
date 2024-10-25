Attribute VB_Name = "F_InvertSelection"
Sub Absolute()
    Dim cell As Range
    
    For Each cell In Selection
            On Error Resume Next
        If cell.Value < 0 Then cell.Value = -cell.Value
    Next cell
End Sub

Sub Invert()
    Dim cell As Range
    
    For Each cell In Selection
            On Error Resume Next
            cell.Value = -cell.Value
    Next cell
End Sub
