Attribute VB_Name = "frmFontChooser"
Attribute VB_Base = "0{EF5A8B79-5A45-4FDB-B98A-9960E9DF2947}{31B0A1ED-D513-4CAC-8F45-227905F522D2}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' ------------------------------------------------------
' Name: frmFontChooser
' Kind: Form
' Purpose: Form that reads installed fonts from word and shows dropdownlist for user
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ------------------------------------------------------
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: baseFontsChooser
' Purpose: clear all fonts in the drop-down list. add some base fonts on top of the list
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub baseFontsChooser()
10        On Error Resume Next
20        With frmFontChooser.fontName
30            .Clear
40            .AddItem ("Arial")
50            .AddItem ("Calibri")
60            .AddItem ("Courier")
70            .AddItem ("Times New Roman")
80            .AddItem (PI_Zh("66F4591A5B574F53002E002E002E"))
90            .Value = "Calibri"
100       End With
End Sub

' ----------------------------------------------------------------
' Procedure Name: cmdCancel_Click
' Purpose: close form, set font back to something installed (calibri)
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Private Sub cmdCancel_Click()
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Private Sub cmdCancel_Click()"         'Debug: Write active procedure to log
30        gblFontnameRibbonActivated = gblstrMessageMore
40        frmFontChooser.Hide
          
exit_proc:
50        Exit Sub
          
ErrorHandler:
60        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
70        Select Case Err.Number
          
              Case Else
80                MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "frmFontChooser_cmdCancel_Click"
90                Resume Next
100               Resume
              
110       End Select
          
End Sub

' ----------------------------------------------------------------
' Procedure Name: cmdOK_Click
' Purpose: update font-variable, close form
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Private Sub cmdOK_Click()
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Private Sub frmFontChooser cmdOK_Click"         'Debug: Write active procedure to log
              
30        gblFontnameRibbonActivated = Me.fontName
                  If gblFontnameRibbonActivated = PI_Zh("66F4591A5B574F53002E002E002E") Then gblFontnameRibbonActivated = gblstrMessageMore
40        gobjRibbon.InvalidateControl "cmbFontName"
50        frmFontChooser.Hide

exit_proc:
60        Exit Sub
          
ErrorHandler:
70        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
80        Select Case Err.Number
              Case 91 'can not invalidate control
90                Resume Next
                  
100           Case Else
110               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "frmFontChooser_cmdOK_Click"
120               Resume Next
130               Resume
              
140       End Select
End Sub

' ----------------------------------------------------------------
' Procedure Name: fontName_Change
' Purpose: react to changed font
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Private Sub fontName_Change()
10        On Error Resume Next
20        Me.lblDemoTextFontSeen.Font = Me.fontName
End Sub

' ----------------------------------------------------------------
' Procedure Name: ListFontsChooser
' Purpose: read installed fonts, sort them a-z and fill dropdown box
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Olaf Noehring ( https://www.datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Public Sub ListFontsChooser()
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Private Sub ListFontsChooser"         'Debug: Write active procedure to log
              
          'src: https://stackoverflow.com/questions/32080762/get-a-list-of-all-fonts-in-vba-excel-2010
          '    Dim strFonts(1000) As String  '-> (1000) removed .... before: read max 1000 different fonts!
          Dim strFonts() As String

          Dim wd As Object, fontID As Variant
          Dim vTemp  As Variant

          Dim i As Integer
          Dim j As Integer
          Dim k As Integer
          Dim m As Integer
          
30        Me.fontName.Clear
40        baseFontsChooser   'add basefonts first
          
50        m = 0
60        On Error Resume Next
70        Set wd = CreateObject("Word.Application")
80        For Each fontID In wd.Fontnames
90            m = m + 1
              '        If m < 1000 Then
              '            strFonts(m) = fontID
              '        End If
              '        'Call me.fontName.AddItem(fontID)
              '        If gblbolDebug = True Then LogInfo "#" & m & " " & fontID
100       Next

110       ReDim strFonts(1 To m) As String
120       Me.lblFntsFound.Caption = PI_Zh("5DF2627E52305B574F53FF1A") & m
130       m = 0
140       For Each fontID In wd.Fontnames
150           m = m + 1
160           strFonts(m) = fontID
              'Call me.fontName.AddItem(fontID)
170           If gblbolDebug = True Then LogInfo "#" & m & " " & fontID
180       Next
          
190       wd.Quit
200       Set wd = Nothing

          '       If m > 1000 Then
          '           MsgBox "You have more than 1000 fonts installed (" & m & "). Only the first (a-z) fonts will be listed.", vbOKOnly Or vbInformation Or vbSystemModal, "Too many fonts!"
          '           m = 1000
          '       End If
          
          
210       On Error GoTo ErrorHandler
          'BubbleSort Array:
          '------------------
          
220       For j = m - 1 To 1 Step -1
              ' Alle links davon liegenden Zeichen auf richtige Sortierung
              ' der jeweiligen Nachfolger überprüfen:
230           If strFonts(j) <> "" Then
240               For i = 1 To j
250                   If gblbolDebug = True Then LogInfo strFonts(i)
                      ' Ist das aktuelle Element seinem Nachfolger gegenüber korrekt sortiert?
260                   If strFonts(i) > strFonts(i + 1) Then
                          ' Element und seinen Nachfolger vertauschen.
270                       vTemp = strFonts(i)
280                       strFonts(i) = strFonts(i + 1)
290                       strFonts(i + 1) = vTemp
300                   End If
310               Next i
320           End If
330       Next j
          '------------------

340       For k = 1 To m
350           Me.fontName.AddItem (strFonts(k))
360           If gblbolDebug = True Then LogInfo "F" & k & " " & fontID
370       Next k
380       Me.fontName.Value = "Calibri"
          
390       If gblbolDebug = True Then LogInfo "Fonts installed: " & m
          
exit_proc:
400       Exit Sub
          
ErrorHandler:
410       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
420       Select Case Err.Number
          
              Case Else
430               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl
440               Resume Next
450               Resume
              
460       End Select
End Sub


Private Sub UserForm_Initialize()
    PI_SetFormFont Me
    Me.Caption = PI_Zh("900962E97AE082825B574F53")
    Me.cmdOK.Caption = PI_Zh("786E5B9A")
    Me.cmdCancel.Caption = PI_Zh("53D66D88")
    Me.lblDemoTextFontSeen.Caption = PI_Zh("7AE082825B574F53988489C80020004100620067007900200031003200330039")
    Me.lblDesc01.Caption = PI_Zh("4ECE5DF25B8988C55B574F534E2D900962E930028BFB53D65B574F5365F653EF80FD77ED668262535F005E76517395ED0020004D006900630072006F0073006F0066007400200057006F007200643002")
    Me.lblFntsFound.Caption = PI_Zh("6B6357288BFB53D65B574F53002E002E002E")
End Sub
