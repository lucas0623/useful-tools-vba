Attribute VB_Name = "F_CalSoilLoad"
Private hSoil As Double
Private dWater As Double
Private sWeight As Double
Private fAngle As Double
Private Kvalue_Identifier As Integer
Private fAngle_rad As Double
Private slopeAngle_deg As Double
Private slopeAngle_rad As Double

Private k As Double
Private SoilForce_total As Double 'in kN/m width
Private SoilForce_aboveWater As Double 'in kN/m width
Private SoilForce_belowWater As Double 'in kN/m width
Private soilPres_waterTable As Double
Private soilPres_bottom As Double
Private title As String
Private isWriteResult As Boolean

Sub CalSoilLoad()
    
    'Get user Input
    title = "Soil Load Calculator v1.2 20200804"
    On Error Resume Next
    hSoil = Application.inputBox("Total Soil Height (in m)", title, Type:=1)
    If hSoil < 0 Then
        MsgBox "Please input positive value!!!!!"
        Exit Sub
    ElseIf hSoil = 0 Then
        End
    End If
    
    dWater = Application.inputBox("Water Table depth below ground level (in m) (Input depth > soil height if water table is not appeared.)", title, hSoil, Type:=1)
    If dWater < 0 Then
        MsgBox "Please input positive value!!!!!"
        Exit Sub
    ElseIf dWater = 0 Then
        End
    End If
    
    sWeight = Application.inputBox("Soil Weight (in kN/m3)", title, 19, Type:=1)
    If sWeight < 0 Then
        MsgBox "Please input positive value!!!!!"
        Exit Sub
    ElseIf sWeight = 0 Then
        End
    End If
    
    fAngle = Application.inputBox("Friction Angle?", title, 33)
    If fAngle < 0 Then
        MsgBox "Please input positive value!!!!!"
        Exit Sub
    ElseIf fAngle = 0 Then
        End
    End If
    
    slopeAngle_deg = Application.inputBox("Slope Angle?(>0, used only when K0 is chosen)", title, 0)
    If slopeAngle_deg < 0 Then
        MsgBox "Please input positive value!!!!!"
        Exit Sub

    End If
    slopeAngle_rad = Application.WorksheetFunction.Radians(slopeAngle_deg)
    
    On Error GoTo 0
    
    'Calculation of K value
    Kvalue_Identifier = inputBox("1 for Ka, 2 for K0, 3 for Kp", , 2)
    
    fAngle_rad = Application.WorksheetFunction.Radians(fAngle)
    If Kvalue_Identifier = 1 Then
        k = (1 - Sin(fAngle_rad)) / (1 + Sin(fAngle_rad)) 'Active Pressure
    ElseIf Kvalue_Identifier = 2 Then
        k = (1 - Sin(fAngle_rad)) * (1 + Sin(slopeAngle_rad)) 'K0
    ElseIf Kvalue_Identifier = 3 Then
        k = (1 + Sin(fAngle_rad)) / (1 - Sin(fAngle_rad)) 'Kp
    Else
        MsgBox ("Wrong Input!!")
        Exit Sub
    End If
    
    'consider the slope of ground for K0 calculation only
    If Not Kvalue_Identifier = 2 Then
        slopeAngle_rad = 0
    End If
    
    'Calculate Soil Pressure and Soil Force
    soilPres_waterTable = sWeight * dWater * k * Cos(slopeAngle_rad)
    SoilForce_aboveWater = soilPres_waterTable * dWater / 2
    
    If dWater < hSoil Then
        soilPres_bottom = soilPres_waterTable + (sWeight - 10) * (hSoil - dWater) * k * Cos(slopeAngle_rad)
        SoilForce_belowWater = (soilPres_waterTable + soilPres_bottom) * (hSoil - dWater) / 2
    Else
        soilPres_bottom = soilPres_waterTable
        SoilForce_belowWater = 0
    End If
    
    SoilForce_total = SoilForce_aboveWater + SoilForce_belowWater
    
    'MsgBox ("Soil Pressure at Water Table Level = " & soilPres_waterTable & "kN/m2")
    'MsgBox ("Soil Pressure at Bottom Level = " & soilPres_bottom & "kN/m2")
    MsgBox ("Total Soil Force = " & Format(SoilForce_total, "0.00") & " kN/m")
    
    'Write Result
    isWriteResult = MsgBox("Do you want to write the result to the active cell? (17 rows x 3 columns would be required)", vbYesNo, title)
    If isWriteResult Then
        WriteResult
    Else
        Exit Sub
    End If
    
    
End Sub

Private Sub WriteResult()
    Dim rRow As Long
    Dim rCol As Long
    Dim cRow As Long
    Dim cCol As Long
    
    rRow = ActiveCell.row
    rCol = ActiveCell.Column
    
    cRow = rRow
    Cells(rRow, rCol) = "User Input"
    Cells(rRow, rCol).Font.Bold = True
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "Soil Height"
    Cells(cRow, rCol + 1) = hSoil
    Cells(cRow, rCol + 2) = "m"
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "Water Table Depth"
    Cells(cRow, rCol + 1) = dWater
    Cells(cRow, rCol + 2) = "m"
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "Friction Angle"
    Cells(cRow, rCol + 1) = fAngle
    Cells(cRow, rCol + 2) = "degree"
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "Slope Angle (used for K0 only)"
    Cells(cRow, rCol + 1) = slopeAngle_deg
    Cells(cRow, rCol + 2) = "degree"
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "Ka"
    Cells(cRow, rCol + 1) = (1 - Sin(fAngle_rad)) / (1 + Sin(fAngle_rad))
    Cells(cRow, rCol + 2) = "[Ka = (1 - Sin(fAngle)) / (1 + Sin(fAngle))]"
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "K0"
    Cells(cRow, rCol + 1) = (1 - Sin(fAngle_rad)) * (1 + Sin(slopeAngle_rad))
    Cells(cRow, rCol + 2) = "[K0 = (1 - Sin(fAngle))* (1 + Sin(slopeAngle_rad)]"
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "Kp"
    Cells(cRow, rCol + 1) = (1 + Sin(fAngle_rad)) / (1 - Sin(fAngle_rad))
    Cells(cRow, rCol + 2) = "[Kp = (1 + Sin(fAngle)) / (1 - Sin(fAngle))]"
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "Soil Weight"
    Cells(cRow, rCol + 1) = sWeight
    Cells(cRow, rCol + 2) = "kN/m3"
    
    cRow = cRow + 1
    
    cRow = cRow + 1
    If Kvalue_Identifier = 1 Then
        Cells(cRow, rCol) = "Result - Active Soil Load" 'Active Pressure
    ElseIf Kvalue_Identifier = 2 Then
        Cells(cRow, rCol) = "Result - Soil Load in Equilibrium State" 'K0
    Else
        Cells(cRow, rCol) = "Result - Passive Soil Load" 'Kp
    End If
    Cells(cRow, rCol).Font.Bold = True
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "Soil Pressure at Water Table Level"
    Cells(cRow, rCol + 1) = soilPres_waterTable
    Cells(cRow, rCol + 2) = "kN/m2"
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "Max Soil Pressure"
    Cells(cRow, rCol + 1) = soilPres_bottom
    Cells(cRow, rCol + 2) = "kN/m2"
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "Total Soil Force"
    Cells(cRow, rCol + 1) = SoilForce_total
    Cells(cRow, rCol + 2) = "kN/m width"
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "Depth of Center of Force"
    Cells(cRow, rCol + 1) = (dWater * 2 / 3 * soilPres_waterTable * dWater / 2 _
                            + ((hSoil - dWater) * 2 / 3 + dWater) * (soilPres_bottom - soilPres_waterTable) * (hSoil - dWater) / 2 _
                            + ((hSoil - dWater) / 2 + dWater) * soilPres_waterTable * (hSoil - dWater)) _
                            / (soilPres_waterTable * dWater / 2 + (soilPres_bottom - soilPres_waterTable) * (hSoil - dWater) / 2 + soilPres_waterTable * (hSoil - dWater))
    Cells(cRow, rCol + 2) = "m"
    
    cRow = cRow + 1
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "Result - Water Load"
    Cells(cRow, rCol).Font.Bold = True
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "Max Water Pressure"
    Cells(cRow, rCol + 1) = (hSoil - dWater) * 10
    Cells(cRow, rCol + 2) = "kN/m2"
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "Total Water Force"
    Cells(cRow, rCol + 1) = (hSoil - dWater) ^ 2 * 10 / 2
    Cells(cRow, rCol + 2) = "kN/m width"
    
    cRow = cRow + 1
    Cells(cRow, rCol) = "Depth of Center of Force"
    Cells(cRow, rCol + 1) = dWater + (hSoil - dWater) * 2 / 3
    Cells(cRow, rCol + 2) = "m"
End Sub
