Attribute VB_Name = "F_SetFooter"
Sub SetFooter()
 ActiveSheet.PageSetup.RightHeader = "Page &P of &N"
 ActiveSheet.PageSetup.RightFooter = "&""Arial""&8" & "Printed at &D &T" & Chr(10) & "  &Z&F"
 'ActiveSheet.PageSetup.RightFooter.Text.Font = 8
End Sub
