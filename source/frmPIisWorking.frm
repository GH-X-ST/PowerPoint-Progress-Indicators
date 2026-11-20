Attribute VB_Name = "frmPIisWorking"
Attribute VB_Base = "0{D9C332EF-DF73-406A-8769-3D72E34EDD6A}{3F3F079F-CDAD-44BA-A05A-7C993927004D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' ------------------------------------------------------
' Name: frmPIisWorking
' Kind: Form
' Purpose: notify user, that PI is being created (as this can take a while)
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2020-03-06
' ------------------------------------------------------
Option Explicit

Private Sub cmdCancelCreation_Click()
          'cmdCancelCreation.Caption = "Creation will stop!"
10        gblCancelCreate = 1
20        DoEvents
End Sub

Private Sub lblCheckNewVersion_Click()
10        On Error Resume Next
20        ActivePresentation.FollowHyperlink "https://url.olaf-noehring.de/progressindicator"
End Sub

Private Sub UserForm_Initialize()
    PI_SetFormFont Me
    Me.Caption = PI_Zh("6B635728751F62108FDB5EA66761")
    Me.lblInfo.Caption = PI_Zh("6B6357284E3A5E7B706F7247751F62108FDB5EA66761FF0C8BF77A0D50193002")
    Me.lblStateXofY.Caption = PI_Zh("6B635728751F6210002E002E002E")
    Me.lblInfoAutoclose.Caption = PI_Zh("751F62107ED3675F540EFF0C6B647A9753E34F1A81EA52A8517395ED3002")
    Me.cmdCancelCreation.Caption = PI_Zh("505C6B62751F6210")
    Me.Label1.Caption = PI_Zh("98848BA15B8C621065F695F44F1A9010987566F465B03002987597628D8A591AFF0C4F308BA1901A5E388D8A7A335B9AFF1B969085CF9875548C8DF38FC798754F1A5F7154CD4F308BA13002")
    Me.lblCheckNewVersion.Caption = PI_Zh("70B951FB8BBF95EE002000500072006F0067007200650073007300200049006E00640069006300610074006F00720020539F72484F5C80057F517AD9FF1A000D000A00680074007400700073003A002F" & _
                        "002F00750072006C002E006F006C00610066002D006E006F0065006800720069006E0067002E00640065002F00700072006F006700720065007300730069006E00640069006300610074006F0072")
    Me.lblProgressbar.Caption = PI_Zh("")
    Me.lblProgressbar.Font.Name = "Courier New"
    Me.cmdCancelCreation.ControlTipText = PI_Zh("5C0657285F53524D98755B8C6210540E505C6B62751F62103002")
End Sub
