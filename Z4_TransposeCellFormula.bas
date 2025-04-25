Attribute VB_Name = "Z4_TransposeCellFormula"
Sub TransposeCellFormula()

    'Kinen Ma 2024-11-19
    'Initial setup of function

    Dim rng As Range
    Dim transposedFormulas As Variant
    Dim formulas As Variant
    
    ' Set the range to the current selection
    Set rng = Selection
    
    ' Check if the selection is not empty
    If rng Is Nothing Then
        MsgBox "Please select a range first."
        Exit Sub
    End If
    
    ' Check if the selection consists of only one cell
    If rng.Cells.count = 1 Then
        MsgBox "Please select multiple cells."
        Exit Sub
    End If
    
    ' Get the formulas from the selection
    formulas = rng.formula
    
    ' Transpose the formulas
    transposedFormulas = Application.Transpose(formulas)
    
    ' Clear the original selection
    rng.ClearContents
    
    ' Determine the size of the transposed array
    Dim rows As Long, cols As Long
    If IsArray(transposedFormulas) Then
        On Error Resume Next
        rows = UBound(transposedFormulas, 1)
        cols = UBound(transposedFormulas, 2)
        
        ' Handle cases where the array might be one-dimensional
        If Err.Number <> 0 Then
            Err.Clear
            rows = 1
            cols = UBound(transposedFormulas)
        End If
        On Error GoTo 0
    Else
        MsgBox "Error in transposing formulas. Please try a different range."
        Exit Sub
    End If
    
    ' Determine the new range based on transposed size
    Dim newRng As Range
    Set newRng = rng.Resize(rows, cols)
    
    ' Apply the transposed formulas to the new range
    newRng.formula = transposedFormulas
    
    MsgBox "Formulas transposed successfully."
End Sub
