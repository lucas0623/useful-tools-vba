Attribute VB_Name = "F_ConvertFormulaRefType"
Sub ConverFormulaReferences()
'Updateby20140603
Dim rng As Range
Dim WorkRng As Range
Dim xName As Name
Dim xIndex As Integer
Dim UIManager As clsUIManager
Set UIManager = New clsUIManager
On Error Resume Next

    Set WorkRng = UIManager.AddInputBox_rng("Range", default:=Selection)
    If WorkRng Is Nothing Then End

    
    xIndex = Application.inputBox("Change formulas to?" & Chr(13) & Chr(13) _
    & "Absolute = 1" & Chr(13) _
    & "Row absolute = 2" & Chr(13) _
    & "Column absolute = 3" & Chr(13) _
    & "Relative = 4", "", 1, Type:=1)
    
    If xIndex = 0 Then End
    
    Set WorkRng = WorkRng.SpecialCells(xlCellTypeFormulas)
    For Each rng In WorkRng
        rng.formula = Application.ConvertFormula(rng.formula, XlReferenceStyle.xlA1, XlReferenceStyle.xlA1, xIndex)
    Next
End Sub

