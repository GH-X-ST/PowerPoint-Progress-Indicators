Attribute VB_Name = "frmCustomImages"
Attribute VB_Base = "0{3579B89F-22EF-4722-8E62-007FCED22505}{BF2A354D-30C9-4244-8DB4-1C3473C2C9F4}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub btnClose_Click()
10        If Me.chkUseCustomShapes = True Then
20            gblshowCustomImage = 1
30            gblCustomImage_Seen = Me.txtImageSeen
40            gblCustomImage_Active = Me.txtImageActive
50            gblCustomImage_UnSeen = Me.txtImageUnSeen
60        Else
70            gblshowCustomImage = 0
80        End If
90        frmCustomImages.Hide
End Sub

Private Sub chkUseCustomShapes_Click()
10        If Me.chkUseCustomShapes = True Then
20            If (Me.txtImageActive <> "") And (Me.txtImageSeen <> "") And (Me.txtImageUnSeen <> "") Then
                  ' alles ok
30            Else
40                MsgBox PI_Zh("8BF75206522B900962E95DF28BB230015F53524D548C672A8BB24E0979CD72B66001768456FE72473002"), vbCritical, PI_Zh("56FE7247672A90099F50")
50                Me.chkUseCustomShapes = False
60                gblshowCustomImage = 0
70            End If
80        End If
End Sub

Private Sub cmdChooseActive_Click()
10        Me.txtImageActive = PickFile
End Sub

Private Sub cmdChooseSeen_Click()
10        Me.txtImageSeen = PickFile
End Sub

Private Sub cmdChooseUnSeen_Click()
10        Me.txtImageUnSeen = PickFile
End Sub


Private Sub UserForm_Initialize()
    PI_SetFormFont Me
    Me.Caption = PI_Zh("81EA5B9A4E498FDB5EA656FE7247")
    Me.lblHint.Caption = PI_Zh("4F7F752865B96CD5FF1A000D000A0031002E00205206522B900962E95DF28BB230015F53524D548C672A8BB24E0979CD72B66001768400200050004E0047002062160020004A00500047002056FE7247" & _
                        "3002000D000A0032002E002052FE9009201C4F7F752881EA5B9A4E4956FE7247201D3002000D000A0033002E002070B951FB201C517395ED201D3002000D000A0034002E002091CD65B0751F62108FDB" & _
                        "5EA667613002000D000A000D000A6CE8610FFF1A000D000A91CD65B0751F621065F6FF0C56FE724765874EF65FC5987B4ECD80FD8BBF95EE3002000D000A56FE72478DEF5F848BBE7F6E4E0D968F6F14" & _
                        "793A65877A3F4FDD5B583002000D000A53D66D8852FE9009540E4F1A6062590D4E4B524D76845F6272B6FF0C4F464E0B62C95217886853EF80FD4ECD663E793A81EA5B9A4E4956FE72473002000D000A" & _
                        "000D000A8BF7900962E95C3A5BF8548C65874EF64F5379EF8F835C0F768456FE724730026BCF987590FD4F1A590D5236591A4E2A56FE6807FF0C8F83592756FE724753EF80FD660E663E964D4F4E901F" & _
                        "5EA63002")
    Me.btnClose.Caption = PI_Zh("517395ED")
    Me.chkUseCustomShapes.Caption = PI_Zh("4F7F752881EA5B9A4E4956FE7247")
    Me.lblImageSeen.Caption = PI_Zh("5DF28BB298759762FF1A")
    Me.Label1.Caption = PI_Zh("5F53524D98759762FF1A")
    Me.Label2.Caption = PI_Zh("672A8BB298759762FF1A")
    Me.Label3.Caption = PI_Zh("540472B6600156FE7247")
    Me.cmdChooseSeen.Caption = PI_Zh("900962E9002E002E002E")
    Me.cmdChooseUnSeen.Caption = PI_Zh("900962E9002E002E002E")
    Me.cmdChooseActive.Caption = PI_Zh("900962E9002E002E002E")
End Sub
