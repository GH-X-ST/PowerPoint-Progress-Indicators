Attribute VB_Name = "frmHelp"
Attribute VB_Base = "0{0E425515-2C97-4B24-AA7C-042467A83EB5}{FC9A7876-2B18-4EE6-825D-764826226905}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' ------------------------------------------------------
' Name: frmHelp
' Kind: Form
' Purpose: offer some help text and URLs
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ------------------------------------------------------
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: cbxDebugMode_Click
' Purpose: Enable Debug Mode: Creates a logfile with a lot of information
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-27
' ----------------------------------------------------------------
Private Sub cbxDebugMode_Click()
          'On Error Resume Next
10        gblbolDebug = cbxDebugMode
          
20        If gblbolDebug = True Then
30            LogInfo "Debugmode was enabled"         'Debug: Write active procedure to log
40            LogInfo "Progress Indicator " & gblProgramVersion
50        Else
60            LogInfo "Debugmode was disabled"         'Debug: Write active procedure to log
70        End If
          
End Sub


Private Sub cmdWebCheck_Click()
          Dim varVersion As Variant

          Dim strSrcVersion As String
          Dim strSrcPI As String
          Dim strQuestion As String

10        strSrcPI = "https://url.olaf-noehring.de/progressindicator"
20        strSrcVersion = "https://url.olaf-noehring.de/pilatestversion"
30        strQuestion = PI_Zh("70B951FB201C662F201D8BBF95EE7248672C4FE1606F00200028") & strSrcVersion & PI_Zh("0029FF0C68C067E54F5C800553D15E037684539F7248662F5426670966F465B03002") & vbCrLf & PI_Zh("598253D173B065B07248672CFF0C53EF524D5F800020") & strSrcPI & PI_Zh("00204E0B8F7D539F72483002539F72484E0D5305542B6B644E2D65875B9A5236724876844FEE65393002")

40        If gblProgramVersionBase & "" = "" Then Initialize

50        Select Case MsgBox(strQuestion, vbYesNo + vbQuestion + vbDefaultButton1, PI_Zh("68C067E5539F724866F465B0FF1F"))
              Case vbYes
60                varVersion = ReadURLFile(strSrcVersion)
70                If Not IsNull(varVersion) And IsNumeric(varVersion) Then
80                    If Val(varVersion) > Val(gblProgramVersionBase) Then
90                        strQuestion = PI_Zh("68C06D4B5230539F724865B07248672C00200028") & varVersion & PI_Zh("002930025F53524D539F724857FA78407248672C4E3A0020") & gblProgramVersionBase & PI_Zh("300270B951FB201C662F201D8BBF95EE0020") & strSrcPI & PI_Zh("00204E0B8F7D539F72483002539F72484E0D5305542B6B644E2D65875B9A5236724876844FEE65393002")
100                       Select Case MsgBox(strQuestion, vbYesNo + vbQuestion + vbDefaultButton1, PI_Zh("62535F004F5C80057F517AD9FF1F"))
                              Case vbYes
110                               ActivePresentation.FollowHyperlink Address:=strSrcPI
120                           Case vbNo
                                  '
130                       End Select
140                   Else
150                       MsgBox PI_Zh("6CA1670968C06D4B523066F465B07684539F724830026B6468C067E54E0D5305542B4E2D65875B9A52367248768466F465B03002"), vbOKOnly, PI_Zh("672A53D173B0539F724866F465B0")
160                   End If
170               Else
180                   strQuestion = PI_Zh("68C067E566F465B059318D2530025F53524D539F724857FA78407248672C4E3A0020") & gblProgramVersionBase & PI_Zh("30028BF78BBF95EE0020") & strSrcPI & PI_Zh("0020624B52A868C067E5FF0C62167A0D540E91CD8BD53002")
190                   MsgBox strQuestion, vbOKOnly + vbCritical, PI_Zh("68C067E566F465B059318D25")
200               End If
210           Case vbNo
                  'do nothing
220       End Select
End Sub

Private Sub Label14_Click()
10        On Error Resume Next
20        ActivePresentation.FollowHyperlink "https://url.olaf-noehring.de/progressindicator"
End Sub

Private Sub Label3_Click()
10        On Error Resume Next
20        ActivePresentation.FollowHyperlink "https://github.com/adammwilson/progress_dots"
End Sub

' ----------------------------------------------------------------
' Procedure Name: Label4_Click
' Purpose: url to how to make ppt sections
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Private Sub Label4_Click()
10        On Error Resume Next
20        ActivePresentation.FollowHyperlink "https://support.office.com/en-us/article/Organize-your-slides-into-sections-DE4BF162-E9CC-4F58-B64A-7AB09443B9F8"
End Sub

Private Sub lblAuthor3_Click()
10        On Error Resume Next
20        ActivePresentation.FollowHyperlink "https://url.olaf-noehring.de/progressindicator"
End Sub

' ----------------------------------------------------------------
' Procedure Name: lblIcons_Click
' Purpose: url to icon source
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Private Sub lblIcons_Click()
10        On Error Resume Next
20        ActivePresentation.FollowHyperlink "http://www.fatcow.com/free-icons"
End Sub

Private Sub lblOriginal_Click()
10        On Error Resume Next
20        ActivePresentation.FollowHyperlink "http://ed.loper.org/projects/progress_dots/"
End Sub

Private Sub UserForm_Initialize()
    PI_SetFormFont Me
    Me.Caption = PI_Zh("6F148BB28FDB5EA667610020002D00205E2E52A9")
    Me.lblAuthors.Caption = PI_Zh("539F72484F5C80050020002870B951FB53EF62535F0094FE63A50029")
    Me.lblOriginal.Caption = PI_Zh("0032003000300037FF1A004500640067006100720020004C006F00700065007200207F165199539F59CB7248672CFF1A0068007400740070003A002F002F00650064002E006C006F007000650072002E" & _
                        "006F00720067002F00700072006F006A0065006300740073002F00700072006F00670072006500730073005F0064006F00740073002F")
    Me.Label3.Caption = PI_Zh("0032003000310036FF1A004100640061006D002000570069006C0073006F006E00205C06987976EE53D15E0352300020004700690074004800750062FF1A00680074007400700073003A002F002F0067" & _
                        "00690074006800750062002E0063006F006D002F006100640061006D006D00770069006C0073006F006E002F00700072006F00670072006500730073005F0064006F00740073")
    Me.lblAuthor3.Caption = PI_Zh("0032003000310037002D0032003000320033FF1A004F006C006100660020004E00F6006800720069006E0067002062695C55529F80FD5E76652F6301002000570069006E0064006F007700733001004D" & _
                        "0061006300204EE553CA002000330032002F0036003400204F4D0020004F0066006600690063006530024F5C80057F517AD9FF1A00680074007400700073003A002F002F0064006100740065006E0062" & _
                        "0061006E006B002D00700072006F006A0065006B0074002E00640065002F")
    Me.lblHelp01.Caption = PI_Zh("572800200050006F0077006500720050006F0069006E007400204E2D5EFA7ACB201C8282201D540EFF0C63D24EF64F1A5C068282540D79F075284F5C7AE0828268079898300264CD4F5C8BF4660EFF1A")
    Me.Label5.Caption = PI_Zh("201C8DF38FC7201D4F1A7ED95F53524D98756DFB52A068078BB0FF0C4F7F5B834E0D663E793A8FDB5EA66761FF1B201C6CBF7528201D4F1A590D52364E0A4E005F2067096548987576848FDB5EA6FF1B" & _
                        "201C6B635E38201D4F1A79FB96648FD94E9B68078BB030026B6464CD4F5C4E0D4F1A969085CF5E7B706F72473002")
    Me.Label6.Caption = PI_Zh("9ED88BA44E0D4E3A00200050006F0077006500720050006F0069006E00740020969085CF9875751F62108FDB5EA656FE6807FF0C53EF5728201C63929664201D7EC44E2D66F465393002")
    Me.Label7.Caption = PI_Zh("201C63929664201D7EC453EF5206522B63A75236969085CF98754EE553CA201C8DF38FC7002F6CBF7528201D68078BB0662F5426751F65483002")
    Me.Label8.Caption = PI_Zh("5F00542F201C81EA52A866F465B0201D540EFF0C66F465398BBE7F6E4F1A91CD65B0751F62108FDB5EA667613002201C988489C8201D4F1A514879FB966463075B9A830356F45185768465E78FDB5EA6" & _
                        "6761FF0C53EA91CD5EFA5C1191CF987597623002988489C8540E53CA589E5220300191CD639298759762540EFF0C8BF770B951FB201C751F6210201D66F465B0516890E863075B9A987597623002")
    Me.Label9.Caption = PI_Zh("529F80FD533A56FE6807676581EA00200046006100740043006F0077FF0C91C7752800200043007200650061007400690076006500200043006F006D006D006F006E0073002000410074007400720069" & _
                        "0062007500740069006F006E00200033002E003000208BB853EFFF1A")
    Me.cbxDebugMode.Caption = PI_Zh("8BB05F558C038BD565E55FD7")
    Me.Label10.Caption = PI_Zh("4E0D542B5E7B706F724776847AE082824E0D4F1A663E793A3002")
    Me.lblSecJokDesc.Caption = PI_Zh("4EE54E0B72796B8A7AE08282540D79F04F1A4FDD755953604F4D548C56FE6807FF0C4F464E0D663E793A7AE08282540D79F0FF1A")
    Me.Label11.Caption = PI_Zh("5F53524D7248672C")
    Me.Label12.Caption = PI_Zh("518D6B2162535F00201C66F4591A5B574F53002E002E002E201D524DFF0C8BF75148900962E953E64E0079CD9ED88BA45B574F5330028BFB53D65B574F5365F653EF80FD77ED6682542F52A80020004D" & _
                        "006900630072006F0073006F0066007400200057006F007200643002")
    Me.Label13.Caption = PI_Zh("4EE54E0B72796B8A7AE082824E2D7684516890E89875976290FD4E0D663E793A7AE0828268079898548C8FDB5EA656FE6807FF1A")
    Me.cmdWebCheck.Caption = PI_Zh("68C067E5539F724866F465B0")
End Sub
