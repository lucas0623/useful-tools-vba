Attribute VB_Name = "F_AddThkBottomLine"
Sub AddThickBottomLineAtPageBreak()
    Dim myWS As Worksheet
    Set myWS = ActiveSheet
    Dim myRow As Long
    Debug.Print myWS.PageSetup.PrintArea

    For i = 1 To myWS.HPageBreaks.count
        myRow = myWS.HPageBreaks(i).Location.row - 1
        With Rows(myRow).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = 1
        End With
    Next
    
End Sub


