Attribute VB_Name = "Z6_NamedRanges"
Option Explicit

Sub ListExternalNamedRangesWithUsage()
    ' Declare variables
    Dim nm As name
    Dim ws As Worksheet
    Dim outputRange As Range
    Dim externalOutputRange As Range
    Dim rowOffset As Long
    Dim extRowOffset As Long
    Dim name As String
    Dim refersTo As String
    Dim reExternal As Object
    Dim reQuote As Object
    Dim dictFormulas As Object ' Store cleaned formulas
    Dim dictUsage As Object   ' Map named ranges to usage cells
    Dim dictExternal As Object ' Store cells with external references
    Dim formulaCells As Range
    Dim cell As Range
    Dim key As Variant
    
    ' Set starting points
    Set outputRange = ActiveCell ' Named ranges table
    Set externalOutputRange = outputRange.Offset(0, 4) ' External formulas table, 4 columns to the right
    rowOffset = 0
    extRowOffset = 0
    
    ' Initialize regex objects
    Set reExternal = CreateObject("VBScript.RegExp")
    reExternal.pattern = "\[[^\]]+\.(xls|xlsx|xlsm)[^\]]*\]" ' Updated to match .xls, .xlsx, .xlsm
    reExternal.IgnoreCase = True
    
    Set reQuote = CreateObject("VBScript.RegExp")
    reQuote.pattern = """[^""]*"""
    reQuote.Global = True
    
    ' Initialize dictionaries
    Set dictFormulas = CreateObject("Scripting.Dictionary")
    Set dictUsage = CreateObject("Scripting.Dictionary")
    Set dictExternal = CreateObject("Scripting.Dictionary")
    
    ' Add headers for named ranges table
    With outputRange
        .Value = "Name"
        .Offset(0, 1).Value = "Refers To"
        .Offset(0, 2).Value = "Used In Cells"
    End With
    rowOffset = 1
    
    ' Add headers for external formulas table
    With externalOutputRange
        .Value = "Cell Address"
        .Offset(0, 1).Value = "Formula"
    End With
    extRowOffset = 1
    
    ' Step 1: Collect all formulas and process them once
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        Set formulaCells = ws.Cells.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        
        If Not formulaCells Is Nothing Then
            For Each cell In formulaCells
                Dim fullFormula As String
                Dim cleanFormula As String
                fullFormula = cell.formula
                cleanFormula = reQuote.Replace(fullFormula, "")
                
                ' Store formula in dictionary
                key = ws.name & "!" & cell.Address
                dictFormulas(key) = cleanFormula
                
                ' Check for external references in full formula
                If reExternal.TEST(fullFormula) Then
                    dictExternal(key) = fullFormula
                End If
            Next cell
        End If
    Next ws
    
    ' Step 2: Build named range usage dictionary
    Dim reName As Object
    Set reName = CreateObject("VBScript.RegExp")
    reName.IgnoreCase = True
    
    For Each nm In ActiveWorkbook.Names
        name = nm.name
        refersTo = nm.refersTo
        
        ' Only process external named ranges
        If InStr(1, refersTo, "[") > 0 And reExternal.TEST(refersTo) Then
            reName.pattern = "\b" & name & "\b"
            Dim usedCells As String
            usedCells = ""
            
            For Each key In dictFormulas.Keys
                Dim formula As String
                formula = dictFormulas(key)
                
                ' Quick check to skip LET formulas with this name as a variable
                If left(UCase(formula), 5) = "=LET(" Then
                    Dim argsStart As Long
                    Dim argsEnd As Long
                    argsStart = InStr(1, formula, "(") + 1
                    argsEnd = InStrRev(formula, ")")
                    If argsEnd > argsStart Then
                        Dim letArgs As String
                        letArgs = Mid(formula, argsStart, argsEnd - argsStart)
                        Dim args() As String
                        args = Split(letArgs, ",")
                        Dim i As Long
                        Dim letVarNames As String
                        letVarNames = ""
                        
                        For i = 0 To UBound(args) - 1 Step 2
                            letVarNames = letVarNames & "," & Trim(UCase(args(i)))
                        Next i
                        
                        If InStr(1, "," & letVarNames & ",", "," & UCase(name) & ",") = 0 Then
                            ' Check the final expression (last argument)
                            Dim expression As String
                            expression = Trim(args(UBound(args)))
                            If reName.TEST(expression) Then
                                usedCells = usedCells & IIf(usedCells = "", key, ", " & key)
                            End If
                        End If
                    End If
                Else
                    ' Non-LET formula
                    If reName.TEST(formula) Then
                        usedCells = usedCells & IIf(usedCells = "", key, ", " & key)
                    End If
                End If
            Next key
            
            ' Output named range details
            With outputRange.Offset(rowOffset, 0)
                .Value = name
                .Offset(0, 1).Value = "'" & refersTo
                .Offset(0, 2).Value = usedCells
            End With
            rowOffset = rowOffset + 1
        End If
    Next nm
    
    ' Step 3: Output external formula references
    For Each key In dictExternal.Keys
        With externalOutputRange.Offset(extRowOffset, 0)
            .Value = key
            .Offset(0, 1).Value = "'" & dictExternal(key)
        End With
        extRowOffset = extRowOffset + 1
    Next key
End Sub
