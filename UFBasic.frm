VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFBasic 
   Caption         =   "Information"
   ClientHeight    =   3.06450e5
   ClientLeft      =   4950
   ClientTop       =   19560
   ClientWidth     =   2.45505e5
   OleObjectBlob   =   "UFBasic.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFBasic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Userform.GeneralForm")

'********************************************************
'This is the Generalized Userform for the used in different modules

'Arthor: Lucas LEUNG
'Update Log
'Aug 2023 - Initial
'*******************************************************

Option Explicit

'For Hiding Default Tilte Bar and using own title bar
Private Const WM_NCLBUTTONDOWN = &HA1&
Private Const HTCAPTION = 2&
'Private Const WS_BORDER = &H800000
Private Declare PtrSafe Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLongPtr Lib "User32" Alias "SetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "User32" (ByVal HWND As Long) As Long
Private Declare PtrSafe Sub ReleaseCapture Lib "User32" ()
Private Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private gHWND As Long
Private Const GWL_STYLE As Long = (-16)


'For
Private mbIsButtonAtCenter As Boolean
Private eventHandlerCollection As New Collection
Private pHoriSpace As Long, pVertSpace As Long
Private cTopPos As Double, pBotMargin As Double
Private pCloseMode As Integer '0 when OK button is pressed, -1 when cancel button is pressed

'******************************************************************************************
'***************************For Customized Title Bar***************************************
'******************************************************************************************
Private Sub HandleDragMove(HWND As Long)
    Call ReleaseCapture
    Call SendMessage(HWND, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub LabelX_Click()
    pCloseMode = -1
    Me.Hide
End Sub

Private Sub MyTitleBar_Caption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If Button = 1 Then HandleDragMove gHWND
End Sub


'*****************************************************************************************************************
'***************************Button Hover Control/ OK Cancel Button Control****************************************
'*****************************************************************************************************************
Private Sub OKButtonInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
'PURPOSE: Make OK Button appear Green when hovered on

  CancelButtonInactive.Visible = True
  OKButtonInactive.Visible = False

End Sub
Private Sub CancelButtonInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button appear Green when hovered on
    CancelButtonInactive.Visible = False
    OKButtonInactive.Visible = True
End Sub

Private Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status
    CancelButtonInactive.Visible = True
    OKButtonInactive.Visible = True
End Sub

Private Sub CancelButton_Click()
    pCloseMode = -1
    Me.Hide
End Sub

Private Sub OKButton_Click()
    pCloseMode = 0
    Me.Hide
End Sub

Public Property Get OKButtonImage() As msforms.Image
    Set OKButtonImage = Me.OKButton
End Property

Public Property Get CancelButtonImage() As msforms.Image
    Set CancelButtonImage = Me.CancelButton
End Property
'*****************************************************************************************************************
'***************************Initialization & Close Behaviour******************************************************
'*****************************************************************************************************************
Private Sub UserForm_Initialize()

    'Remove default title bar and create own window
    Dim frm As Long
    Dim wHandle As Long
    wHandle = FindWindow(vbNullString, Me.caption)
    frm = GetWindowLong(wHandle, GWL_STYLE)
    'frm = frm And WS_BORDER

    SetWindowLongPtr wHandle, -16, 0
    DrawMenuBar wHandle
    gHWND = wHandle
    
    'Other Initializetion
    'Me.width = 300
    Me.height = 300
    pHoriSpace = 10
    pVertSpace = 10
    pBotMargin = 10
    cTopPos = MyTitleBar_Border.height + 10
End Sub

Private Sub Userform_Activate()
    MyTitleBar_Caption.width = Me.width
    MyTitleBar_Border.width = Me.width
    
    'Position t
    Reposition_labelX
    If mbIsButtonAtCenter Then
        RepositionOkAndCancelButtonsToCenter
    Else
        RepositionOkAndCancelButtonsToRight
    End If
    RepositionOkAndCancelButtonsToCenter
    RepositionUF
End Sub

Public Sub Initialize(Optional width As Double = 300, _
                        Optional isButtonAtCenter As Boolean = True)
    Me.width = width
    mbIsButtonAtCenter = isButtonAtCenter
End Sub

'*************************************************************************************************************************
'*************************************Basic Methods***********************************************************************
'*************************************************************************************************************************
Private Sub Reposition_labelX()
    With LabelX
        .top = 2
        .left = Me.width - 20
    End With
End Sub

Public Sub RepositionOkAndCancelButtonsToCenter()
    With OKButton
        .top = Me.height - pBotMargin - .height
        .left = (Me.width / 2 - .width) / 2
    End With
    With OKButtonInactive
        .top = Me.height - pBotMargin - .height
        .left = (Me.width / 2 - .width) / 2
    End With
    With CancelButton
        .top = Me.height - pBotMargin - .height
        .left = Me.width / 2 + (Me.width / 2 - .width) / 2
    End With
    With CancelButtonInactive
        .top = Me.height - pBotMargin - .height
        .left = Me.width / 2 + (Me.width / 2 - .width) / 2
    End With
End Sub

Public Sub RepositionOkAndCancelButtonsToRight()
    With OKButton
        .top = Me.height - pBotMargin - .height
        .left = Me.width - pHoriSpace * 2 - .width * 2
    End With
    With OKButtonInactive
        .top = Me.height - pBotMargin - .height
        .left = Me.width - pHoriSpace * 2 - .width * 2
    End With
    With CancelButton
        .top = Me.height - pBotMargin - .height
        .left = Me.width - pHoriSpace - .width
    End With
    With CancelButtonInactive
        .top = Me.height - pBotMargin - .height
        .left = Me.width - pHoriSpace - .width
    End With
    
    'Debug.Print OKButton.top
End Sub

Private Sub RepositionUF()
   Me.top = Application.top + (Application.UsableHeight / 2) - (Me.height / 2)
   Me.left = Application.left + (Application.UsableWidth / 2) - (Me.width / 2)
End Sub

Property Get CloseState() As Integer
    CloseState = pCloseMode
End Property
Property Let TitleBarCaption(str As String)
    MyTitleBar_Caption.caption = str
End Property

'*************************************************************************************************************************
'******************************Building UI Elements***********************************************************************
'*************************************************************************************************************************
Public Sub AddSpace(Optional space As Double = 10)
    cTopPos = cTopPos + space
End Sub
Public Sub AddLabel(caption As String, Optional tipText As String, Optional labelWidth As Double, _
                    Optional isBold As Boolean = True, Optional isUnderline As Boolean = False, Optional isAutoSize As Boolean = False, _
                    Optional fontSize As Double = 10)
    Dim label As msforms.label
    Dim checkbox As msforms.checkbox
    
    If Not caption = vbNullString Then
        Set label = Me.Controls.Add("Forms.Label.1")
        With label
            .height = 12
            .top = cTopPos
            .left = pHoriSpace
            .caption = caption
            If labelWidth = 0 Then
                .width = (Me.width - pHoriSpace * 2)
            Else
                .width = labelWidth
            End If
            .AutoSize = isAutoSize
            .ControlTipText = tipText
            .Font.Bold = isBold
            .Font.Underline = isUnderline
            .Font.size = fontSize
        End With
    End If
    
    cTopPos = cTopPos + label.height + 5
    
End Sub
Public Function AddButton(Optional caption As String = "Button", _
                    Optional width As Double = 50, Optional height As Double = 15, _
                    Optional top As Double = 10, Optional left As Double = 10, Optional pageNum As Integer = -1) As msforms.CommandButton
    Dim btn As msforms.CommandButton

    Set btn = Me.Controls.Add("Forms.CommandButton.1")
    
    With btn
        .height = height 'Application.Max(label.Height, 15)
        .width = width
        .top = top
        .left = left
        .caption = caption
    End With
        
    Set AddButton = btn
    
End Function

Public Sub AddCheckBox(reCheckbox As msforms.checkbox, Optional title As String = "Please Select", _
                Optional Description As String, Optional isCheck As Boolean = False, Optional tipText As String, _
                Optional labelWidth As Double)
    Dim label As msforms.label
    Dim checkbox As msforms.checkbox
    
    If Not title = vbNullString Then
        Set label = Me.Controls.Add("Forms.Label.1")
        With label
            .top = cTopPos
            .left = pHoriSpace
            .caption = title
            If labelWidth = 0 Then
                .width = (Me.width - pHoriSpace) / 2 - pHoriSpace
            Else
                .width = labelWidth
            End If
            .AutoSize = True
            .ControlTipText = tipText
        End With
    End If
    
    Set checkbox = Me.Controls.Add("Forms.Checkbox.1")
    With checkbox
        .height = 15 'Application.Max(label.Height, 15)
        .width = Me.width - label.width - pHoriSpace * 2
        .top = cTopPos
        If labelWidth = 0 Then
            .left = (Me.width - pHoriSpace) / 2 + pHoriSpace
        Else
            .left = labelWidth + pHoriSpace
        End If
        .Value = isCheck
        .caption = Description
        
    End With
        
    Set reCheckbox = checkbox
    cTopPos = cTopPos + checkbox.height + 10
    
End Sub

Sub AddCheckBox_double(reCheckbox As msforms.checkbox, reCheckbox2 As msforms.checkbox, _
                        Optional title As String = "Please Select", Optional description1 As String, _
                        Optional description2 As String, Optional isCheck1 As Boolean = False, Optional isCheck2 As Boolean = False, _
                        Optional tipText As String)
    
    Dim label As msforms.label, Label2 As msforms.label
    Dim checkbox As msforms.checkbox, CheckBox2 As msforms.checkbox
    
    AddCheckBox checkbox, title, description1, isCheck1, tipText
    
    cTopPos = cTopPos - checkbox.height - 10
    
    Set CheckBox2 = Me.Controls.Add("Forms.Checkbox.1")
    With CheckBox2
        .height = 15 'Application.Max(label.Height, 15)
        .width = (Me.width - pHoriSpace) / 4
        .top = cTopPos
        .left = 3 * (Me.width - pHoriSpace) / 4 '(Me.Width - pHoriSpace) / 2 + pHoriSpace
        .Value = isCheck2
        .caption = description2
    End With
    

    Set reCheckbox = checkbox
    Set reCheckbox2 = CheckBox2
    
    cTopPos = cTopPos + checkbox.height + 10
    
End Sub

Public Sub AddComboBox_Empty(reComboBox As msforms.comboBox, Optional title As String = "Please Select", Optional tipText As String)
    Dim label As msforms.label
    Dim comboBox As msforms.comboBox
    
    Set label = Me.Controls.Add("Forms.Label.1")
    With label
        .top = cTopPos
        .left = pHoriSpace
        .caption = title
        .width = (Me.width - pHoriSpace) / 2 - pHoriSpace
        .AutoSize = True
        .ControlTipText = tipText
        'AdjustLabelHeight label
        'Debug.Print "Label Height = " & .Height
    End With
    
    Set comboBox = Me.Controls.Add("Forms.ComboBox.1")
    With comboBox
        .height = Application.Max(label.height, 18)
        .width = (Me.width - pHoriSpace) / 2 - pHoriSpace * 2
        .top = cTopPos
        .left = (Me.width - pHoriSpace) / 2 + pHoriSpace
    End With
    
    If comboBox.height > label.height Then
        label.top = label.top + (comboBox.height - label.height) / 2
    End If
    
    Set reComboBox = comboBox
    cTopPos = cTopPos + comboBox.height + 10
End Sub

Sub AddComboBox(arr As Variant, reComboBox As msforms.comboBox, Optional title As String = "Please Select", _
                Optional defaultVal As Variant = vbNullString, Optional tipText As String)
    AddComboBox_Empty reComboBox, title, tipText:=tipText
    reComboBox.List = arr
    
    If defaultVal = vbNullString Then
        reComboBox.Value = arr(0)
    Else: reComboBox.Value = defaultVal
    End If
   
End Sub

Sub AddInputBox(reTextBox As msforms.Textbox, Optional title As String = "Please Input", Optional def_value As Variant)
    Dim label As msforms.label
    Dim Textbox As msforms.Textbox
    
    Set label = Me.Controls.Add("Forms.Label.1")
    With label
        .top = cTopPos
        .left = pHoriSpace
        .caption = title
        .width = (Me.width - pHoriSpace) / 2 - pHoriSpace
        .AutoSize = True
    End With
    
    Set Textbox = Me.Controls.Add("Forms.TextBox.1")
    With Textbox
        .height = Application.Max(label.height, 15)
        .width = (Me.width - pHoriSpace) / 2 - pHoriSpace * 2
        .top = cTopPos
        .left = (Me.width - pHoriSpace) / 2 + pHoriSpace
        .Value = def_value
    End With
    
    Set reTextBox = Textbox
    cTopPos = cTopPos + Textbox.height + 10
End Sub

'Sub AddRngInputBox(reTextBox As MSForms.Textbox, Optional title As String = "Please Input", Optional def_value As Variant)
'    Dim label As MSForms.label
'    Dim Textbox As MSForms.Textbox
'
'    AddInputBox Textbox, title, def_value
'    Textbox.DropButtonStyle = fmDropButtonStyleReduce
'    Textbox.ShowDropButtonWhen = fmShowDropButtonWhenAlways
'
'    'Dim obj As EventUFRngInput
'    'Set obj = New EventUFRngInput
'    obj.Initialize Me, Textbox
'    eventHandlerCollection.Add obj
'
'    Set reTextBox = Textbox
'
'End Sub
'Sub AddRngInputBox2(reTextBox As MSForms.Textbox, Optional Title As String = "Please Input", Optional def_value As Variant)
'    Dim label As MSForms.label
'    Dim Textbox As MSForms.Textbox
'    Dim obj As ModelessRefEdit
'
'    Set obj = New ModelessRefEdit
'    Set label = Me.Controls.Add("Forms.Label.1")
'    With label
'        .top = cTopPos
'        .left = pHoriSpace
'        .caption = Title
'        .width = (Me.width - pHoriSpace) / 2 - pHoriSpace
'        .AutoSize = True
'    End With
'
'
'    Set Textbox = Me.Controls.Add("Forms.TextBox.1")
'    With Textbox
'        .height = Application.Max(label.height, 15)
'        .width = (Me.width - pHoriSpace) / 2 - pHoriSpace * 2
'        .top = cTopPos
'        .left = (Me.width - pHoriSpace) / 2 + pHoriSpace
'        .value = def_value
'        .DropButtonStyle = fmDropButtonStyleReduce
'        .ShowDropButtonWhen = fmShowDropButtonWhenAlways
'    End With
'    Set obj.UF = Me
'    obj.addRefEditBox Textbox
'    eventHandlerCollection.Add obj
'
'    Set reTextBox = Textbox
'    cTopPos = cTopPos + Textbox.height + 10
'End Sub
'Sub AddRefEdit(reRefEdit As refEdit.refEdit, Optional Title As String = "Please Input", Optional def_value As Variant)
'
'    Dim label As MSForms.label
'    Dim refEdit As refEdit.refEdit
'
'    Set label = Me.Controls.Add("Forms.Label.1")
'    With label
'        .top = cTopPos
'        .left = pHoriSpace
'        .caption = Title
'        .width = (Me.width - pHoriSpace) / 2 - pHoriSpace
'        .AutoSize = True
'    End With
'
'    Set refEdit = Me.Controls.Add("RefEdit.Ctrl")
'    With refEdit
'        .height = Application.Max(label.height, 30)
'        .width = (Me.width - pHoriSpace) / 2 - pHoriSpace * 2
'        .top = cTopPos
'        .left = (Me.width - pHoriSpace) / 2 + pHoriSpace
'        .MultiLine = True
'        .value = def_value
'    End With
'
'    Set reRefEdit = refEdit
'    cTopPos = cTopPos + refEdit.height + 10
'End Sub
Public Sub AddSelectionBox_Empty(reListBox As msforms.listbox, Optional title As String = "listbox 1", _
                        Optional height_LB As Long = 140, Optional width_LB As Long)
        
    'add the items
    Dim listbox1 As msforms.listbox
    Dim label As msforms.label
    
    If width_LB = 0 Then
        width_LB = Me.width - pHoriSpace * 2
    End If
    
    Set label = Me.Controls.Add("Forms.Label.1")
    With label
        .top = cTopPos
        .left = pHoriSpace
        .caption = title
        .width = width_LB + 5
        .AutoSize = True
    End With
    
    Set listbox1 = Me.Controls.Add("Forms.ListBox.1")
    With listbox1
        .height = height_LB
        .width = width_LB
        .top = cTopPos + label.height
        .left = pHoriSpace
        .MultiSelect = fmMultiSelectExtended
    End With
    
    Set reListBox = listbox1

    cTopPos = cTopPos + height_LB + label.height + 10
    
End Sub

'Public Sub AddSelectionBoxMulti_Empty(reListBox As MSForms.listBox, Optional title As String = "listbox 1", _
'                        Optional height_LB As Long = 140, Optional width_LB As Long, _
'                        Optional width_cmdBtm As Long = 50, _
'                        Optional is_reListBox2 As Boolean = False, Optional reListBox2 As MSForms.listBox, _
'                        Optional title2 As String = "listbox 2", _
'                        Optional isCreateFrame As Boolean = False, Optional frameTitle As String = "", _
'                        Optional reCmdBtn1 As MSForms.CommandButton, Optional reCmdBtn2 As MSForms.CommandButton)
'
'    'add the items
'    Dim newFrame As MSForms.frame
'    Dim listbox1 As MSForms.listBox, ListBox2 As MSForms.listBox
'    Dim cmdB1 As MSForms.CommandButton, cmdB2 As MSForms.CommandButton
'    Dim label As MSForms.label, Label2 As MSForms.label
'
'    Dim btnEvent As EventSelectionBoxMulti
'    Dim framePosX As Double, framePosY As Double
'
'    If width_LB = 0 Then
'        width_LB = Me.width / 2 - pHoriSpace * 2 - width_cmdBtm / 2
'    End If
'
'    If isCreateFrame Then
'        Set newFrame = Me.Controls.Add("forms.frame.1", "TEST", True)
'        framePosX = 10
'        framePosY = 10
'
'        With newFrame
'            .caption = frameTitle
'            .top = cTopPos
'            .left = pHoriSpace
'            .width = 2 * width_LB + 4 * pHoriSpace + width_cmdBtm + framePosX
'            .height = height_LB + framePosY + 15
'            .Font.Bold = True
'        End With
'    End If
'
'    Set label = Me.Controls.Add("Forms.Label.1")
'    With label
'        .top = cTopPos + framePosY
'        .left = pHoriSpace + framePosX
'        .caption = title
'        .width = width_LB + 5
'        .AutoSize = True
'    End With
'
'    Set listbox1 = Me.Controls.Add("Forms.ListBox.1")
'    With listbox1
'        .height = height_LB
'        .width = width_LB
'        .top = cTopPos + label.height + framePosY
'        .left = pHoriSpace + framePosX
'        .MultiSelect = fmMultiSelectExtended
'    End With
'
'    Set cmdB1 = Me.Controls.Add("Forms.CommandButton.1")
'    With cmdB1
'        .caption = "->"
'        .height = height_LB / 4
'        .width = width_cmdBtm
'        .top = listbox1.top + height_LB / 2 - 10 - .height
'        '.Top = cTopPos + label.Height + height_LB / 2 - 10 - .Height
'        .left = listbox1.left + listbox1.width + pHoriSpace
'    End With
'    Set reCmdBtn1 = cmdB1
'
'    Set cmdB2 = Me.Controls.Add("Forms.CommandButton.1")
'    With cmdB2
'        .caption = "<-"
'        .height = height_LB / 4
'        .width = width_cmdBtm
'        .top = listbox1.top + height_LB / 2 + 10
'        .left = listbox1.left + listbox1.width + pHoriSpace
'    End With
'    Set reCmdBtn2 = cmdB2
'
'    Set Label2 = Me.Controls.Add("Forms.Label.1")
'    With Label2
'        .top = cTopPos + framePosY
'        .left = cmdB1.left + cmdB1.width + pHoriSpace + framePosX
'        .caption = title2
'        .width = width_LB
'        .AutoSize = True
'    End With
'
'    Set ListBox2 = Me.Controls.Add("Forms.ListBox.1")
'    With ListBox2
'        .height = height_LB
'        .width = width_LB
'        .top = cTopPos + label.height + framePosY
'        .left = cmdB1.left + cmdB1.width + pHoriSpace + framePosX
'        .MultiSelect = fmMultiSelectExtended
'    End With
'
''    Set btnEvent = New EventSelectionBoxMulti
''    btnEvent.Init Me, ListBox1, ListBox2, cmdB1, cmdB2
''    eventHandlerCollection.Add btnEvent
'    'mColButtons.Add btnEvent
'
'    Set reListBox = listbox1
'    If is_reListBox2 Then
'        Set reListBox2 = ListBox2
'    End If
'
'    'Resize windows
'    'Me.Width = Application.Max(Me.Width, width_cmdBtm + width_LB * 2 + pHoriSpace * 5 + framePosX * 3)
''    Me.height = cTopPos + def_botMargin + height_LB + label.height + 15
'
'    cTopPos = cTopPos + height_LB + label.height + 10 + framePosY * 2
'
'    If isCreateFrame Then
'        newFrame.Visible = True
'    End If
'
'End Sub

'Sub AddSelectionBoxMulti(arr As Variant, reListBox As MSForms.listBox, Optional title As String = "listbox 1", _
'                        Optional title2 As String = "listbox 2", Optional height_LB As Long = 140, _
'                        Optional width_LB As Long, Optional width_cmdBtm As Long = 50, _
'                        Optional is_reListBox2 As Boolean = False, Optional reListBox2 As MSForms.listBox, _
'                        Optional isCreateFrame As Boolean = False, Optional frameTitle As String = "", _
'                        Optional reCmdBtn1 As MSForms.CommandButton, Optional reCmdBtn2 As MSForms.CommandButton)
'
'    'add the items
'    Dim newFrame As MSForms.frame
'    Dim listbox1 As MSForms.listBox, ListBox2 As MSForms.listBox
'    Dim cmdB1 As MSForms.CommandButton, cmdB2 As MSForms.CommandButton
'    Dim label As MSForms.label, Label2 As MSForms.label
'
'    Dim btnEvent As EventSelectionBoxMulti
'    Dim framePosX As Double, framePosY As Double
'
'If width_LB = 0 Then
'        width_LB = Me.width / 2 - pHoriSpace * 2.5 - width_cmdBtm / 2
'    End If
'
'    If isCreateFrame Then
'        Set newFrame = Me.Controls.Add("forms.frame.1", "TEST", True)
'        framePosX = 10
'        framePosY = 10
'
'        With newFrame
'            .caption = frameTitle
'            .top = cTopPos
'            .left = pHoriSpace
'            .width = 2 * width_LB + 4 * pHoriSpace + width_cmdBtm + framePosX
'            .height = height_LB + framePosY + 15
'            .Font.Bold = True
'        End With
'    End If
'
'    Set label = Me.Controls.Add("Forms.Label.1")
'    With label
'        .top = cTopPos + framePosY
'        .left = pHoriSpace + framePosX
'        .caption = title
'        .width = width_LB + 5
'        .AutoSize = True
'    End With
'
'    Set listbox1 = Me.Controls.Add("Forms.ListBox.1")
'    With listbox1
'        .height = height_LB
'        .width = width_LB
'        .top = cTopPos + label.height + framePosY
'        .left = pHoriSpace + framePosX
'        .MultiSelect = fmMultiSelectExtended
'        .List = arr
'    End With
'
'    Set cmdB1 = Me.Controls.Add("Forms.CommandButton.1")
'    With cmdB1
'        .caption = "->"
'        .height = height_LB / 4
'        .width = width_cmdBtm
'        .top = listbox1.top + height_LB / 2 - 10 - .height
'        '.Top = cTopPos + label.Height + height_LB / 2 - 10 - .Height
'        .left = listbox1.left + listbox1.width + pHoriSpace
'    End With
'    Set reCmdBtn1 = cmdB1
'
'    Set cmdB2 = Me.Controls.Add("Forms.CommandButton.1")
'    With cmdB2
'        .caption = "<-"
'        .height = height_LB / 4
'        .width = width_cmdBtm
'        .top = listbox1.top + height_LB / 2 + 10
'        .left = listbox1.left + listbox1.width + pHoriSpace
'    End With
'    Set reCmdBtn2 = cmdB2
'
'    Set Label2 = Me.Controls.Add("Forms.Label.1")
'    With Label2
'        .top = cTopPos + framePosY
'        .left = cmdB1.left + cmdB1.width + pHoriSpace + framePosX
'        .caption = title2
'        .width = width_LB
'        .AutoSize = True
'    End With
'
'    Set ListBox2 = Me.Controls.Add("Forms.ListBox.1")
'    With ListBox2
'        .height = height_LB
'        .width = width_LB
'        .top = cTopPos + label.height + framePosY
'        .left = cmdB1.left + cmdB1.width + pHoriSpace + framePosX
'        .MultiSelect = fmMultiSelectExtended
'    End With
'
''    Set btnEvent = New EventSelectionBoxMulti
''    btnEvent.Init Me, ListBox1, ListBox2, cmdB1, cmdB2
''    eventHandlerCollection.Add btnEvent
'    'mColButtons.Add btnEvent
'
'    Set reListBox = listbox1
'    If is_reListBox2 Then
'        Set reListBox2 = ListBox2
'    End If
'
'    'Resize windows
'    'Me.width = Application.Max(Me.width, width_cmdBtm + width_LB * 2 + pHoriSpace * 5 + framePosX * 3)
''    Me.height = cTopPos + def_botMargin + height_LB + label.height + 15
'
'    cTopPos = cTopPos + height_LB + label.height + 10 + framePosY * 2
'
'    If isCreateFrame Then
'        newFrame.Visible = True
'    End If
'
'End Sub

'Sub AddMultiColumnsListBox(arr As Variant, arr_header As Variant, Optional ControlCB As MSForms.ComboBox, _
'                        Optional ControlCB2 As MSForms.ComboBox, Optional dict As Object, _
'                        Optional reListBox As MSForms.listBox, Optional title As String = "listbox 1", _
'                        Optional height_LB As Long = 140, Optional width_LB As Long = 260, Optional colWidth As String, _
'                        Optional isCreateFrame As Boolean = False, Optional frameTitle As String = "", Optional isSortList As Boolean = True)
'    Dim newFrame As MSForms.frame
'    Dim listbox1 As MSForms.listBox
'    'Dim cmdB1 As MSForms.CommandButton, cmdB2 As MSForms.CommandButton
'    Dim label As MSForms.label, Label2 As MSForms.label
'    Dim i As Long
'    'Dim btnEvent As clsSelectionBoxMultiEvent
'    Dim framePosX As Double, framePosY As Double
'    Dim comboBoxEvent As EventComboBoxToListBox
'
'    Set label = Me.Controls.Add("Forms.Label.1")
'    With label
'        .height = 16
'        .top = cTopPos + framePosY
'        .left = pHoriSpace + framePosX
'        .caption = title
'        .width = width_LB + 5
'        .AutoSize = False
'        .Font.Bold = True
'        .Font.Underline = False
'        .Font.size = 10
'
'    End With
'
'    'Set The headers
'    Dim lb_header As MSForms.listBox, header As Variant
'    ReDim header(0, LBound(arr_header) To UBound(arr_header))
'    For i = LBound(arr_header) To UBound(arr_header)
'        header(0, i) = arr_header(i)
'    Next i
'
'    Set lb_header = Me.Controls.Add("Forms.ListBox.1")
'    With lb_header
'        .height = 16
'        .width = width_LB
'        .top = cTopPos + label.height + framePosY
'        .left = pHoriSpace + framePosX
'        .ColumnCount = UBound(arr, 2) - LBound(arr, 2) + 1
'        '.MultiSelect = fmMultiSelectExtended
'        .List = header
'        .SpecialEffect = fmSpecialEffectFlat
'        .BackColor = RGB(200, 200, 200)
'        .Font.Bold = True
'        .Font.size = 8
'        .Locked = True
'        .ColumnWidths = colWidth
'    End With
'
'    Set listbox1 = Me.Controls.Add("Forms.ListBox.1")
'    With listbox1
'        .height = height_LB
'        .width = width_LB
'        .top = cTopPos + label.height + framePosY + lb_header.height
'        .left = pHoriSpace + framePosX
'        .ColumnCount = UBound(arr, 2) - LBound(arr, 2) + 1
'        '.MultiSelect = fmMultiSelectExtended
'        .List = arr
'        .ColumnWidths = colWidth
'    End With
'
'    If Not ControlCB Is Nothing Then
'        Set comboBoxEvent = New EventComboBoxToListBox
'        comboBoxEvent.Init Me, listbox1, ControlCB, dict, ControlCB2
'        eventHandlerCollection.Add comboBoxEvent
'    End If
'    'Resize windows
'    Me.width = Application.Max(Me.width, width_LB + pHoriSpace * 3 + framePosX * 3)
'    'Me.height = cTopPos + def_botMargin + height_LB + label.height + 15
'
'    cTopPos = cTopPos + height_LB + label.height + 10 + framePosY * 2 + 15
'End Sub

'Sub AddComboBoxVisibilityControl(cb_master As MSForms.ComboBox, _
'                        cb_slave() As Object, val_visible As String)
'    Dim btnEvent As EventComboBoxControl
'    Set btnEvent = New EventComboBoxControl
'    btnEvent.Init Me, cb_master, cb_slave, val_visible
'    eventHandlerCollection.Add btnEvent
'End Sub

Public Sub AdjustHeight()
    Me.height = cTopPos + pBotMargin + OKButton.height + pVertSpace
End Sub
'Private Sub AdjustLabelHeight(lb As MSForms.label)
'    Dim minHeight As Double
'    minHeight = 18
'    If lb.Height < minHeight Then
'        lb.AutoSize = False
'        lb.Height = minHeight
'    End If
'End Sub


