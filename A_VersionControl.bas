Attribute VB_Name = "A_VersionControl"
Sub ShowVersion()

    MsgBox "Stable Version 1.3.1 (19 Nov 2024)" & Chr(10) & _
    Chr(10) & _
    "https://github.com/lucas0623/Useful-Tools-VBA-Modules/tree/main" & Chr(10) & _
    Chr(10) & _
    "Arthor: Lucas LEUNG & Kinen MA"
End Sub

Sub Help()
    Set objShell = CreateObject("Wscript.Shell")
    objShell.Run ("https://connecthkuhk-my.sharepoint.com/:f:/g/personal/u3506883_connect_hku_hk/Ek6CdqMdGw9JrQmIQDcTw64BLIIzDJqqhVe6RFYAPt9Myw?e=4plXGR")
End Sub
'Update History
'****************************************************
'Version 1.2.0(02 Oct 2024)
'Added functions for udpating chart series and swapping graph x-y axis.
'updated the function of FormulaConcat

'Version 1.1.1(19 Jan 2024)
'Updated the 'Named Range'>'Change Scope' function. Now the reference of graph series will be changed as well.

'Version 1.0.3 (12 Jan 2023)
'Add Functions for setting up the footer of the page

'Version 1.0.2 (7 Dec 2022)
'Add Functions and updated UI
'But the Functions not links to the button yet

'Version 1.0.1 (30 Nov 2022)
'Add functions: "Condense Columns in Table", "Multiply table"

'Version 1.0.0 (25 Nov 2022)
