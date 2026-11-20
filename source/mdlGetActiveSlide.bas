Attribute VB_Name = "mdlGetActiveSlide"
' ------------------------------------------------------
' Name: mdlGetActiveSlide
' Kind: Modul
' Purpose: Get the active slide from matching window
' Author: http://www.officeoneonline.com/vba.html
' Date: 2017-06-21
' ------------------------------------------------------
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: GetActiveSlide
' Purpose: Returns active slide
' Procedure Kind: Function
' Procedure Access: Public
' Parameter Wnd (Object):
' Return Type: slide
' Author: http://www.officeoneonline.com/vba.html
' Date: 2017-06-21
' ----------------------------------------------------------------
Function GetActiveSlide(ByVal Wnd As Object) As slide
          'base sources from
          ' http://www.officeoneonline.com/vba.html
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Function GetActiveSlide(ByVal Wnd As Object) As slide"         'Debug: Write active procedure to log
              
          Dim Sld As slide

30        Select Case TypeName(Wnd)
              Case "DocumentWindow"
40                Set Sld = Wnd.Selection.SlideRange(1)
50                Set GetActiveSlide = Sld
                  
60            Case "SlideShowWindow"
70                Set Sld = Wnd.View.slide
80                Set GetActiveSlide = Sld
                  
90            Case Else
                  'Beep
                  '            Sld = 0
                  
100       End Select
          
exit_proc:
110       Exit Function
          
ErrorHandler:
120       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log

130       Select Case Err.description
              Case "Selection.SlideRange : Invalid request.  SlideRange cannot be constructed from a Master."
140               MsgBox PI_Zh("8BF7900051FA5E7B706F72476BCD724889C656FEFF0C5728666E901A7F168F9189C656FE4E2D4F7F75288FDB5EA667613002"), vbCritical Or vbSystemModal, PI_Zh("5F53524D89C656FE65E06CD54F7F75288FDB5EA66761")
150               GetActiveSlide = 0
160               Resume exit_proc
170               Resume


180           Case Else
190               Select Case Err.Number

                      Case -2147188160, 438
200                       MsgBox PI_Zh("65E06CD583B753D65F53524D98757801FF0C5C065207636252307B2C00200031002098753002") & vbCrLf & PI_Zh("598253CD590D51FA73B06B6463D0793AFF0C8BF74FDD755995198BEF4FE1606F5E7680547CFB7EF462A480053002"), vbCritical Or vbSystemModal, PI_Zh("65E06CD583B753D65F53524D98757801")
210                       ActivePresentation.Slides(1).Select        'Which method is saver on MAC?
                          'Application.ActiveWindow.View.GotoSlide 1  'Which method is saver on MAC?
220                       Resume
230                       Resume exit_proc
              
240                   Case Else
250                       MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "GetActiveSlide"
260                       Resume Next
270                       Resume
              
280               End Select
       
290       End Select
       
End Function

' ----------------------------------------------------------------
' Procedure Name: TestDesignWindow
' Purpose: testing
' Procedure Kind: Sub
' Procedure Access: Public
' Author: http://www.officeoneonline.com/vba.html
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub TestDesignWindow()
          'base sources from
          ' http://www.officeoneonline.com/vba.html
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Sub TestDesignWindow"         'Debug: Write active procedure to log
              
30        MsgBox PI_Zh("5F53524D98757801FF1A") + _
              Str(GetActiveSlide(ActiveWindow).SlideIndex)
exit_proc:
40        Exit Sub
          
ErrorHandler:
50        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
60        Select Case Err.Number

              
              Case Else
70                MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl
80                Resume Next
90                Resume
              
100       End Select
End Sub

' ----------------------------------------------------------------
' Procedure Name: TestSlideShowWindow
' Purpose: testing
' Procedure Kind: Sub
' Procedure Access: Public
' Author: http://www.officeoneonline.com/vba.html
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub TestSlideShowWindow()
          'base sources from
          ' http://www.officeoneonline.com/vba.html
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Sub TestSlideShowWindow"         'Debug: Write active procedure to log
              
30        If SlideShowWindows.Count > 0 Then
40            MsgBox PI_Zh("5F53524D98757801FF1A") + _
                  Str(GetActiveSlide(SlideShowWindows(1)).SlideIndex)
50        End If
exit_proc:
60        Exit Sub
          
ErrorHandler:
70        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
80        Select Case Err.Number
          
              Case Else
90                MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl
100               Resume Next
110               Resume
              
120       End Select
End Sub

