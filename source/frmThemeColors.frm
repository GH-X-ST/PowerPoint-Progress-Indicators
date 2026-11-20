Attribute VB_Name = "frmThemeColors"
Attribute VB_Base = "0{335DF451-35AA-4695-B13A-943F97EFA782}{876C7287-2645-471B-91AA-27C04B306655}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' ------------------------------------------------------
' Name: frmThemeColors
' Kind: Form
' Purpose: show a colorswatch. colors come from the currently active theme. colors can be clicked on, entered by RGB value or by sliders
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ------------------------------------------------------
Option Explicit
Dim mdlClrSelected As Long

Private Sub cmdCancel_Click()
10        frmThemeColors.Hide
End Sub

Private Sub cmdColor0_Click()
10        SetColorFromButton cmdColor0
End Sub

Private Sub cmdColor1_Click()
10        SetColorFromButton cmdColor1
End Sub

Private Sub cmdColor2_Click()
10        SetColorFromButton cmdColor2
End Sub

Private Sub cmdColor3_Click()
10        SetColorFromButton cmdColor3
End Sub

Private Sub cmdColor4_Click()
10        SetColorFromButton cmdColor4
End Sub

Private Sub cmdColor5_Click()
10        SetColorFromButton cmdColor5
End Sub

Private Sub cmdColor6_Click()
10        SetColorFromButton cmdColor6
End Sub

Private Sub cmdColor7_Click()
10        SetColorFromButton cmdColor7
End Sub

Private Sub cmdColor8_Click()
10        SetColorFromButton cmdColor8
End Sub

Private Sub cmdColor9_Click()
10        SetColorFromButton cmdColor9
End Sub

Private Sub cmdColor10_Click()
10        SetColorFromButton cmdColor10
End Sub

Private Sub cmdColor11_Click()
10        SetColorFromButton cmdColor11
End Sub

Private Sub cmdColor12_Click()
10        SetColorFromButton cmdColor12
End Sub

Private Sub cmdColor13_Click()
10        SetColorFromButton cmdColor13
End Sub

Private Sub cmdColor14_Click()
10        SetColorFromButton cmdColor14
End Sub

Private Sub cmdColor15_Click()
10        SetColorFromButton cmdColor15
End Sub

Private Sub cmdColor16_Click()
10        SetColorFromButton cmdColor16
End Sub

Private Sub cmdColor17_Click()
10        SetColorFromButton cmdColor17
End Sub

Private Sub cmdColor18_Click()
10        SetColorFromButton cmdColor18
End Sub

Private Sub cmdColor19_Click()
10        SetColorFromButton cmdColor19
End Sub

Private Sub cmdColor20_Click()
10        SetColorFromButton cmdColor20
End Sub

Private Sub cmdColor21_Click()
10        SetColorFromButton cmdColor21
End Sub

Private Sub cmdColor22_Click()
10        SetColorFromButton cmdColor22
End Sub

Private Sub cmdColor23_Click()
10        SetColorFromButton cmdColor23
End Sub

Private Sub cmdColor24_Click()
10        SetColorFromButton cmdColor24
End Sub

Private Sub cmdColor25_Click()
10        SetColorFromButton cmdColor25
End Sub

Private Sub cmdColor26_Click()
10        SetColorFromButton cmdColor26
End Sub

Private Sub cmdColor27_Click()
10        SetColorFromButton cmdColor27
End Sub

Private Sub cmdColor28_Click()
10        SetColorFromButton cmdColor28
End Sub

Private Sub cmdColor29_Click()
10        SetColorFromButton cmdColor29
End Sub

Private Sub cmdColor30_Click()
10        SetColorFromButton cmdColor30
End Sub

Private Sub cmdColor31_Click()
10        SetColorFromButton cmdColor31
End Sub

Private Sub cmdColor32_Click()
10        SetColorFromButton cmdColor32
End Sub

Private Sub cmdColor33_Click()
10        SetColorFromButton cmdColor33
End Sub

Private Sub cmdColor34_Click()
10        SetColorFromButton cmdColor34
End Sub

Private Sub cmdColor35_Click()
10        SetColorFromButton cmdColor35
End Sub

Private Sub cmdColor36_Click()
10        SetColorFromButton cmdColor36
End Sub

Private Sub cmdColor37_Click()
10        SetColorFromButton cmdColor37
End Sub

Private Sub cmdColor38_Click()
10        SetColorFromButton cmdColor38
End Sub

Private Sub cmdColor39_Click()
10        SetColorFromButton cmdColor39
End Sub

Private Sub cmdColor40_Click()
10        SetColorFromButton cmdColor40
End Sub

Private Sub cmdColor41_Click()
10        SetColorFromButton cmdColor41
End Sub

Private Sub cmdColor42_Click()
10        SetColorFromButton cmdColor42
End Sub

Private Sub cmdColor43_Click()
10        SetColorFromButton cmdColor43
End Sub

Private Sub cmdColor44_Click()
10        SetColorFromButton cmdColor44
End Sub

Private Sub cmdColor45_Click()
10        SetColorFromButton cmdColor45
End Sub

Private Sub cmdColor46_Click()
10        SetColorFromButton cmdColor46
End Sub

Private Sub cmdColor47_Click()
10        SetColorFromButton cmdColor47
End Sub

Private Sub cmdColor48_Click()
10        SetColorFromButton cmdColor48
End Sub

Private Sub cmdColor49_Click()
10        SetColorFromButton cmdColor49
End Sub

Private Sub cmdColor50_Click()
10        SetColorFromButton cmdColor50
End Sub

Private Sub cmdColor51_Click()
10        SetColorFromButton cmdColor51
End Sub

Private Sub cmdColor52_Click()
10        SetColorFromButton cmdColor52
End Sub

Private Sub cmdColor53_Click()
10        SetColorFromButton cmdColor53
End Sub

Private Sub cmdColor54_Click()
10        SetColorFromButton cmdColor54
End Sub

Private Sub cmdColor55_Click()
10        SetColorFromButton cmdColor55
End Sub

Private Sub cmdColor56_Click()
10        SetColorFromButton cmdColor56
End Sub

Private Sub cmdColor57_Click()
10        SetColorFromButton cmdColor57
End Sub

Private Sub cmdColor58_Click()
10        SetColorFromButton cmdColor58
End Sub

Private Sub cmdColor59_Click()
10        SetColorFromButton cmdColor59
End Sub

Private Sub cmdColor60_Click()
10        SetColorFromButton cmdColor60
End Sub

Private Sub cmdColor61_Click()
10        SetColorFromButton cmdColor61
End Sub

Private Sub cmdColor62_Click()
10        SetColorFromButton cmdColor62
End Sub

Private Sub cmdColor63_Click()
10        SetColorFromButton cmdColor63
End Sub

Private Sub cmdColor64_Click()
10        SetColorFromButton cmdColor64
End Sub

Private Sub cmdColor65_Click()
10        SetColorFromButton cmdColor65
End Sub

Private Sub cmdColor66_Click()
10        SetColorFromButton cmdColor66
End Sub

Private Sub cmdColor67_Click()
10        SetColorFromButton cmdColor67
End Sub

Private Sub cmdColor68_Click()
10        SetColorFromButton cmdColor68
End Sub

Private Sub cmdColor69_Click()
10        SetColorFromButton cmdColor69
End Sub

Private Sub cmdColor70_Click()
10        SetColorFromButton cmdColor70
End Sub

Private Sub cmdColor71_Click()
10        SetColorFromButton cmdColor71
End Sub


Sub ColorWriteBack()
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Sub ColorWriteBack"         'Debug: Write active procedure to log
          Dim intRed As Integer
          Dim intGreen As Integer
          Dim intBlue As Integer
          
30        intRed = Me.dotSeenRt
40        intGreen = Me.dotSeenGt
50        intBlue = Me.dotSeenBt
          
60        optionsFromRibbon Me.lblColorSelected.Tag, CStr(intRed) & ";" & CStr(intGreen) & ";" & CStr(intBlue)
          
exit_proc:
70        Exit Sub
          
ErrorHandler:
80        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
90        Select Case Err.Number
              Case Else
100               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "ColorWriteBack"
110               Resume Next
120               Resume
              
130       End Select
End Sub
    
Sub ShowColor()
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Sub ShowColor"         'Debug: Write active procedure to log
          Dim strRGBArray() As String
30        Me.lblColorSelected.BackColor = mdlClrSelected

40        If gblbolDebug = True Then LogInfo "Color: " & Color_to_RGB(mdlClrSelected)
50        strRGBArray = Split(Color_to_RGB(mdlClrSelected), ";")

60        Me.dotSeenR.Value = Val(strRGBArray(0))
70        Me.dotSeenG.Value = Val(strRGBArray(1))
80        Me.dotSeenB.Value = Val(strRGBArray(2))

exit_proc:
90        Exit Sub
          
ErrorHandler:
100       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
110       Select Case Err.Number
              Case Else
120               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl
130               Resume Next
140               Resume
              
150       End Select
End Sub

' ----------------------------------------------------------------
' Procedure Name: SetColorFromButton
' Purpose: Universal action for a pressed button: Set the color of the pressed button
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter oButton (label): color to be used (background of the LABEL as it seems no backgroundcolor for buttons on MAC)
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub SetColorFromButton(ByVal oButton As label)
10        If gblbolDebug = True Then LogInfo "Sub SetColorFromButton (ByVal oButton As label)"         'Debug: Write active procedure to log
          'Sub SetColorFromButton(ByVal oButton As CommandButton)
20        mdlClrSelected = oButton.BackColor
30        ShowColor
End Sub


' ----------------------------------------------------------------
' Procedure Name: RecolorNewColor
' Purpose: color has been chosen, set variable, recolor RGB fields
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub RecolorNewColor()
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Sub RecolorNewColor"         'Debug: Write active procedure to log
          '    Me.dotSeenRt.Value = Me.dotSeenR & ";" & Me.dotSeenG & ";" & Me.dotSeenB
          
30        Me.dotSeenR.ControlTipText = Me.dotSeenR.Value
40        Me.dotSeenG.ControlTipText = Me.dotSeenG.Value
50        Me.dotSeenB.ControlTipText = Me.dotSeenB.Value
          
60        Me.dotSeenRt = Me.dotSeenR.Value
70        Me.dotSeenGt = Me.dotSeenG.Value
80        Me.dotSeenBt = Me.dotSeenB.Value
          
90        Me.lblColorSelected.BackColor = RGB(Me.dotSeenR, Me.dotSeenG, Me.dotSeenB)
          
exitProc:
100       Exit Sub

ErrorHandler:
110       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
120       Select Case Err.Number
          
              Case Else
130               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "frmThemeColors_cmdOK_Click"
140               Resume Next
150               Resume
              
160       End Select
          
End Sub


' ----------------------------------------------------------------
' Procedure Name: cmdOK_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Private Sub cmdOK_Click()
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Priv Sub frmThemeColors_cmdOK_Click"         'Debug: Write active procedure to log
              
30        ColorWriteBack
40        frmThemeColors.Hide
          
exit_proc:
50        Exit Sub
          
ErrorHandler:
60        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
70        Select Case Err.Number
          
              Case Else
80                MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "frmThemeColors_cmdOK_Click"
90                Resume Next
100               Resume
              
110       End Select
End Sub

' ----------------------------------------------------------------
' Procedure Name: dotSeenB_Change
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Private Sub dotSeenB_Change()
10        On Error Resume Next
20        If gblbolDebug = True Then LogInfo "Priv Sub dotSeenB_Change"         'Debug: Write active procedure to log
30        RecolorNewColor
End Sub

' ----------------------------------------------------------------
' Procedure Name: dotSeenBt_Change
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Private Sub dotSeenBt_Change()
10        On Error Resume Next
20        If gblbolDebug = True Then LogInfo "Private Sub dotSeenBt_Change()"         'Debug: Write active procedure to log
              
          Dim i As Integer
30        i = Round((Rnd() * 255), 0) 'if user enters stupid value, just set a random one
40        i = Me.dotSeenBt
50        Me.dotSeenB = i
60        Me.dotSeenBt = i
70        Me.dotSeenBt.BackColor = RGB(0, 0, i)
End Sub

' ----------------------------------------------------------------
' Procedure Name: dotSeenG_Change
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Private Sub dotSeenG_Change()
10        On Error Resume Next
20        If gblbolDebug = True Then LogInfo "Priv Sub dotSeenG_Change"         'Debug: Write active procedure to log
30        RecolorNewColor
End Sub

' ----------------------------------------------------------------
' Procedure Name: dotSeenGt_Change
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Private Sub dotSeenGt_Change()
10        On Error Resume Next
20        If gblbolDebug = True Then LogInfo "Priv Sub dotSeenGt_Change"         'Debug: Write active procedure to log
              
          Dim i As Integer
30        i = Round((Rnd() * 255), 0) 'if user enters stupid value, just set a random one
40        i = Me.dotSeenGt
50        Me.dotSeenG = i
60        Me.dotSeenGt = i
70        Me.dotSeenGt.BackColor = RGB(0, i, 0)
End Sub

' ----------------------------------------------------------------
' Procedure Name: dotSeenR_Change
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Private Sub dotSeenR_Change()
10        On Error Resume Next
20        If gblbolDebug = True Then LogInfo "Priv Sub dotSeenR_Change"         'Debug: Write active procedure to log
30        RecolorNewColor
End Sub

' ----------------------------------------------------------------
' Procedure Name: dotSeenRt_Change
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Private Sub dotSeenRt_Change()
10        On Error Resume Next
20        If gblbolDebug = True Then LogInfo "Priv Sub dotSeenRt_Change"         'Debug: Write active procedure to log
              
          Dim i As Integer
30        i = Round((Rnd() * 255), 0) 'if user enters stupid value, just set a random one
40        i = Me.dotSeenRt
50        Me.dotSeenR = i
60        Me.dotSeenRt = i
70        Me.dotSeenRt.BackColor = RGB(i, 0, 0)
End Sub



Private Sub UserForm_Initialize()
    PI_SetFormFont Me
    Me.Caption = PI_Zh("900962E9989C8272")
    Me.lblColorSelected.Caption = PI_Zh("65B0989C8272")
    Me.lblColorOld.Caption = PI_Zh("539F989C8272")
    Me.cmdOK.Caption = PI_Zh("786E5B9A")
    Me.cmdCancel.Caption = PI_Zh("53D66D88")
    Me.lblColorSelected.ControlTipText = PI_Zh("988489C865B07684989C8272")
    Me.lblColorOld.ControlTipText = PI_Zh("539F67657684989C8272")
    Me.cmdOK.ControlTipText = PI_Zh("5E94752862409009989C8272")
    Me.cmdCancel.ControlTipText = PI_Zh("4FDD7559539F67657684989C8272")
    Me.lblRed1.ControlTipText = PI_Zh("7EA28272901A9053002000280030002D0032003500350029")
    Me.lblG1.ControlTipText = PI_Zh("7EFF8272901A9053002000280030002D0032003500350029")
    Me.lblB1.ControlTipText = PI_Zh("84DD8272901A9053002000280030002D0032003500350029")
    Me.dotSeenRt.ControlTipText = PI_Zh("7EA28272901A9053768453418FDB5236503C002000280030002D0032003500350029")
    Me.dotSeenGt.ControlTipText = PI_Zh("7EFF8272901A9053768453418FDB5236503C002000280030002D0032003500350029")
    Me.dotSeenBt.ControlTipText = PI_Zh("84DD8272901A9053768453418FDB5236503C002000280030002D0032003500350029")
End Sub
