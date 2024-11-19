Attribute VB_Name = "A_RibbonControl"
Public Sub ProcessRibbon_UsefulTools(Control As IRibbonControl)
    Select Case Control.ID
        'call different macro based on button name pressed
        'Format Table Tab
        Case "btn_AddBotBorder"
            B_FormatTable.AddBorder_Hori
        Case "btn_ClrBotBorder"
            B_FormatTable.ClearBorder_Hori
        Case "btn_ClrRightBorder"
            B_FormatTable.ClearBorder_Right
        Case "btn_BoldMax"
            B_FormatTable.BoldMaxOfSameGroup
        Case "btn_MergeSameIdentity"
            B_FormatTable.MergeCellsInSameGroup
            
        Case "btn_AddBotBorderAtPageBreak"
            AddThickBottomLineAtPageBreak
        Case "btn_SetupFooter"
            F_SetFooter.SetFooter
            
        'GSA Tab
'        Case "btn_CreateGSAListFromSection"
            
        'Defined Names Tab
        Case "btn_CopyDefinedNames"
            F_CopyDefinedNames.CopyDefinedNamesToSheet
        Case "btn_ChangeNamesScopeLocalToGlobal"
            F_CopyDefinedNames.ChangeDefinedNamesScopeLocalToGlobal
        Case "btn_ChangeNamesScopeGlobalToLocal"
            F_CopyDefinedNames.ChangeDefinedNamesScopeGlobalToLocal
            
            
        'File System Tab
        Case "btn_GetFilesInFolder"
        D_FileManager.GetAllFilesInFolder
        
        Case "btn_ChangeFileName"
        
'        'Transform Table Tab
'        Case "btn_CondenseTableToCol"
'            C_TransformTable.CondenseToCol
'        Case "btn_CondenseColumnsInTable"
'            C_TransformTable.CondenseColumnInTable
'        Case "btn_Condense3DTableTo2D"
'            C_TransformTable.CondenseToTable
'        Case "btn_MultiplyTable"
'            C_TransformTable.MultiplyTable
'        Case "btn_FilterTable"
'            C_TransformTable.FilterTable
'        Case "btn_DelEmtpyRows"
'            C_TransformTable.DelEmptyRows
'        Case "btn_DelRowsIfEqual"
'            C_TransformTable.DelRowsIfEqual
'        Case "btn_FillEmtpyRows"
'            C_TransformTable.FillEmptyCells
            
        'Miscellaneous Tab
        Case "btn_ShowCellCalculationStep"
            Z_Functions.FormulaConcat_to_CellFormula
        Case "btn_ConvertCellRef"
            ConverFormulaReferences
        Case "btn_ChangeSheetVisibility"
            Z2_ChangeSheetVisibility.ManageSheetVisibilityInteractive
        Case "btn_MathAutoCorrect"
            Z3_MathAutoCorrect.AutoCorrectMathSymbols
        Case "btn_TransposeCellFormula"
            Z4_TransposeCellFormula.TransposeCellFormula
        
'        Case "btn_CalSoilLoad"
'            F_CalSoilLoad.CalSoilLoad
'        Case "btn_ABS"
'            F_InvertSelection.Absolute
'        Case "btn_Invert"
'            F_InvertSelection.Invert
        
        
            
        'Graph Tab
        Case "btn_Graph1"
            G_Graph.FormatGraph1
        Case "btn_UpdateChartReference"
            UpdateChartSeries.UpdateChartsToCurrentWorksheet
        Case "btn_UpdateAxisSpacing"
            UpdateChartSeries.UpdateChartPlottingExtent
        Case "btn_SwapAxis"
            SwapPlotData.SwapPlotDataXY
        'Info
        Case "btn_help"
            A_VersionControl.Help
        Case "btn_version"
            A_VersionControl.ShowVersion
            
    End Select
End Sub

