Attribute VB_Name = "Z_Functions"
Function FormulaConcat(rng As Range, Optional round_num As Integer = -2, Optional include_equal_sign = False) As Variant
'Kinen Ma, 2023-12-22, Version 1.0
'
'More reliable than FormulaValue
'cross referened to other spreadsheet or excel does not work yet
'need to fix if Round(text) results in an error


    Dim text As String
    Dim res As String
    Dim block As String
    Dim quote As Boolean
    Dim x As String
    Dim x2 As String
    Dim operators As String
    Dim count As Integer
    
    count = 0
    
    
    operators = "|+|-|*|/|!|^|(|)| |,|=|.|#|&|<|>|"
    text = rng.formula
    text = Replace(text, "$", "")
    res = "=CONCAT("
    block = ""
    quote = False
    
    Dim i As Integer
    If Mid(text, 1, 1) = "=" Then
        If include_equal_sign = False Then
            i = 2 'to include the = sign at the beninging or not :)
        Else
            i = 1
        End If
    Else
        i = 1
    End If
    
    While i < Len(text) + 1 And count < 10000
        count = count + 1
        x = UCase(Mid(text, i, 1))
        If quote Then 'in a string quote, skip
            If x = """" Then
                quote = Not quote
                block = block + "''"
                res = Helper_FormulatoConcat_BlockToRes(block, res, True) 'call BlockToRes to clear the block with end of string quote
                block = ""
            Else
                block = block + (Mid(text, i, 1))
            End If
            i = i + 1
        ElseIf x = """" Then
            block = block + "''"
            i = i + 1
            quote = Not quote
        ElseIf Not x Like "*[A-Z]*" Then
            i = i + 1
            'res = res + x
            block = block + x
        Else
            
            Dim check As Integer
            Dim n As Integer
            check = 0
            n = 1
            
            res = Helper_FormulatoConcat_BlockToRes(block, res, True) 'call BlockToRes to clear the block
            block = ""
            
            Dim while2 As Boolean
            while2 = True
            While i + n < Len(text) + 1 And while2 And count < 10000
                count = count + 1
                x2 = UCase(Mid(text, i + n, 1))
                If x2 Like "*[A-Z]*" And check = 0 Then
                    n = n + 1
                ElseIf x2 Like "*[0-9]*" And check = 0 Then
                    n = n + 1
                    check = 1
                ElseIf x2 Like "*[0-9]*" And check = 1 Then
                    n = n + 1
                    check = 1
                ElseIf x2 Like "*[A-Z]*" Then
                    n = n + 1
                    check = -1
                ElseIf InStr(1, operators, "|" & x2 & "|", vbTextCompare) Then 'if this char is an operator
                    If check = 1 Then
                        Dim val As String
                        cell_address = Mid(text, i, n)
                        If IsNumeric(Range(cell_address).Value) Then
                            'rounded_cell_address = "Round(" + cell_address + ",1)" 'Round to 1 decimal place
                            rounded_cell_address = Helper_FormulaConcat_RoundNum(cell_address, round_num)
                        Else
                            rounded_cell_address = cell_address
                        End If
                        res = Helper_FormulatoConcat_BlockToRes(rounded_cell_address, res, False) 'call BlockToRes to add this cell reference
                    Else
                        res = Helper_FormulatoConcat_BlockToRes(Mid(text, i, n), res, True) 'call BlockToRes to clear this texts
                    End If
                    i = i + n
                    n = 0
                    while2 = False 'exit this inner loop
                End If
                
                If i + n = Len(text) + 1 And check = 1 Then 'for the end of the string
                    Dim val2 As String
                    val2 = Mid(text, i, n)
                    If IsNumeric(Range(val2).Value) Then
                        'rounded_cell_address2 = "Round(" + val2 + ",1)" 'Round to 1 decimal place
                        rounded_cell_address2 = Helper_FormulaConcat_RoundNum(val2, round_num)
                    Else
                        rounded_cell_address2 = val2
                    End If
                    res = Helper_FormulatoConcat_BlockToRes(rounded_cell_address2, res, False) 'call BlockToRes to add this cell reference
                    
                    i = i + n
                ElseIf i + n = Len(text) + 1 Then
                    res = Helper_FormulatoConcat_BlockToRes(Mid(text, i, n), res, True) 'call BlockToRes to append the text
                    i = i + n
                End If
            Wend
        End If
    Wend
    
    If block <> "" Then
        res = Helper_FormulatoConcat_BlockToRes(block, res, True)
    End If
    
    res = left(res, Len(res) - 1) + ")"
    FormulaConcat = res '***RESULT***
 
    
    
    'Merge consecutive strings
    res_merged = res
    
    Dim left_idx As Integer
    Dim lengths As Integer
    quote = False
    i = 9
    
    
    While i < Len(res_merged) + 1 And count < 10000
        count = count + 1
        x = Mid(res_merged, i, 1)
        x_prev = Mid(res_merged, i - 1, 1)
        x_next = Mid(res_merged, i + 1, 1)
        If x = """" And quote = False Then
            left_idx = i
            quote = True
        ElseIf x = """" And quote = True Then
            quote = False
        ElseIf x = "," And quote = False And x_prev = """" And x_next = """" Then
            res_merged = left(res_merged, i - 2) + Right(res_merged, Len(res_merged) - i - 1)
            i = i - 2
            quote = True
        End If
        i = i + 1
    Wend
    
    FormulaConcat = res_merged '***RESULT, merged consecutive strings***
    
End Function

Private Function Helper_FormulaConcat_RoundNum(cell_address, round_num) As Variant

    Dim d As Double
    d = Range(cell_address).Value
    
    If round_num = -1 Then
        Helper_FormulaConcat_RoundNum = cell_address
    ElseIf round_num = -2 Then
        If d < 2 Then
            Helper_FormulaConcat_RoundNum = "Round(" + cell_address + ",2)"
        ElseIf d < 100 Then
            Helper_FormulaConcat_RoundNum = "Round(" + cell_address + ",1)"
        Else
            Helper_FormulaConcat_RoundNum = "Round(" + cell_address + ",0)"
        End If
    Else
        Helper_FormulaConcat_RoundNum = "Round(" + cell_address + "," + CStr(round_num) + ")"
    End If

End Function


Private Function Helper_FormulatoConcat_BlockToRes(block, res, with_quote As Boolean) As Variant
    
    If with_quote Then
        Helper_FormulatoConcat_BlockToRes = res + """" + block + """" + ","
    Else
        Helper_FormulatoConcat_BlockToRes = res + block + ","
    End If

End Function

Sub FormulaConcat_to_CellFormula()
    
    Dim ran As Range
    Set ran = ActiveSheet.UsedRange
    
    Dim c As Range
    Dim firstAddress As String

    With ran
        Set c = .Find("=FormulaConcat(", LookIn:=xlFormulas)
        If Not c Is Nothing Then
            firstAddress = c.Address
            Do
                c.formula = c.Value
                Set c = .FindNext(c)
            Loop While Not c Is Nothing
        End If
    End With
    
End Sub
