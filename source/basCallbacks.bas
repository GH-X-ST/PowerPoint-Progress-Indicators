Attribute VB_Name = "basCallbacks"
' ------------------------------------------------------
' Name: basCallbacks
' Kind: Modul
' Purpose: Handles Ribbon-Actions. Basecode created by ribboncreator2016
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ------------------------------------------------------
Option Explicit

Public gobjRibbon As IRibbonUI
Public bolEnabled As Boolean    ' Used in Callback "getEnabled"
' Further informations in Callback "getEnabled"
                               
Public bolVisible As Boolean    ' Used in Callback "getVisible"
' More information in Callback "getVisible
' Further informations in Callback "getVisible
' For Sample Callback "GetContent"
Public Type ItemsVal
    id As String
    label As String
    imageMso As String
End Type



' DynamicMenu
Sub GetContent(control As IRibbonControl, ByRef XMLString)
10        If gblbolDebug = True Then LogInfo "Sub GetContent (control As IRibbonControl, ByRef XMLString)"         'Debug: Write active procedure to log
          ' Sample for a Ribbon XML "getContent" Callback
          ' See also http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Callbacks:dynamicMenu_-_getContent
          '     and: http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Ribbon_XML___Controls:Dynamic_Menu
          ' Beispiel fuer einen Ribbon XML - "getContent" Callback
          ' Siehe auch: http://www.accessribbon.de/index.php?Access_-_Ribbons:Callbacks:dynamicMenu_-_getContent
          '       und : http://www.accessribbon.de/?Access_-_Ribbons:Ribbon_XML___Controls:Dynamic_Menu
20        Select Case control.id
              
              Case Else
30                XMLString = getXMLForDynamicMenu()
40        End Select
       
End Sub

Sub GetDescription(control As IRibbonControl, ByRef description)
10        If gblbolDebug = True Then LogInfo "Sub GetDescription (control As IRibbonControl, ByRef description)"         'Debug: Write active procedure to log
          ' Callbackname in XML File "getDescription"
          ' To set the property "description" to a Ribbon Control
20        Select Case control.id
              
              Case Else
30                description = "*getDescription*"
40        End Select
End Sub

Public Sub GetEnabled(control As IRibbonControl, ByRef enabled)
          If gblbIndicatorUpdateRunning Then
              enabled = False
              Exit Sub
          End If
10        If gblbolDebug = True Then LogInfo "Sub GetEnabled (control As IRibbonControl, ByRef enabled)"         'Debug: Write active procedure to log
          ' Callbackname in XML File "getEnabled"
          
          ' To set the property "enabled" to a Ribbon Control
          ' For further information see: http://www.accessribbon.de/en/index.php?Downloads:12
          ' Setzen der Enabled Eigenschaft eines Ribbon Controls
          ' Weitere Informationen: http://www.accessribbon.de/index.php?Downloads:12
20        Select Case control.id
              
              Case "ddcGeneralProgressbarEvenSpace"
30                enabled = True
                  
40            Case "ddcIgnoreIndicator"
50                enabled = True
                  
60            Case "ddcIgnoreSections"
70                enabled = True
                  
80            Case "btnGeneralProgressbarCreate"
                  ' Button
                  ' In Tab:   tabProgressDots
                  ' In Group: grpGeneral
90                enabled = True
                  
100           Case "btnGeneralProgressbarDelete"
                  ' Button
                  ' In Tab:   tabProgressDots
                  ' In Group: grpGeneral
110               enabled = True
                  
120           Case "btnGeneralProgressbarHelp"
                  ' Button
                  ' In Tab:   tabProgressDots
                  ' In Group: grpGeneral
130               enabled = True
                  
140           Case "btnGeneralProgressbarHideTrue"
                  ' Button
                  ' In Tab:   tabProgressDots
                  ' In Group: grpSlide
150               enabled = True
                  
160           Case "btnGeneralProgressbarHideFalse"
                  ' Button
                  ' In Tab:   tabProgressDots
                  ' In Group: grpSlide
170               enabled = True
                  
180           Case "cmbFontName"
                  ' ComboBox
                  ' In Tab:   tabProgressDots
                  ' In Group: grpFont
190               enabled = True
                  
200           Case "ddcIndicatorType"
                  ' Dropdown
                  ' In Tab:   tabProgressDots
                  ' In Group: grpIndicator
210               enabled = True
                  
220           Case "cbxFontColorUseBG"
230               enabled = True
       
240           Case "cbxFontColorUseBG_Current"
250               enabled = True
       
260           Case "cbxFontColorUseBG_UnSeen"
270               enabled = True
        
        
280           Case "cbxIndicatorFill_Seen"
290               enabled = True
        
300           Case "cbxIndicatorFill_Current"
310               enabled = True
        
320           Case "cbxIndicatorFill_UnSeen"
330               enabled = True


                  
340           Case "cbxSectionplacement"
350               enabled = True
                  
360           Case "cbxGeneralProgressbarDirect"
370               enabled = True
              
380           Case "cbxIgnoreIgnored"
                  ' Checkbox
                  ' In Tab:   tabProgressDots
                  ' In Group: grpIgnore
390               enabled = True
                  
400           Case "cbxIgnoreHidden"
                  ' Checkbox
                  ' In Tab:   tabProgressDots
                  ' In Group: grpIgnore
410               enabled = True
                  
                  
420           Case Else
430               enabled = True
440       End Select
End Sub

Sub GetLabel(control As IRibbonControl, ByRef label)
10        If gblbolDebug = True Then LogInfo "Sub GetLabel(control As IRibbonControl, ByRef label)"         'Debug: Write active procedure to log
          ' Callbackname in XML File "getLabel"
          ' To set the property "label" to a Ribbon Control
20        Select Case control.id
              
              Case Else
30                label = "*getLabel*"
40        End Select
End Sub

Sub GetPressedCheckBox(control As IRibbonControl, ByRef bolReturn)
10        If gblbolDebug = True Then LogInfo "Sub GetPressedCheckBox(control As IRibbonControl, ByRef bolReturn)"         'Debug: Write active procedure to log
          ' Callbackname in XML File "GetPressedCheckBox"
          
          ' Callback for checkbox
          ' indicates how the control is displayed

          
          Dim i As Integer

          ''debug for breaking
          'MsgBox "checkbox 1"
          'MsgBox "checkbox 2"
          
20        Select Case control.id
                      
                      
              Case "cbxFontColorUseBG"
30                i = gblintUseTextBGCol

40            Case "cbxFontColorUseBG_Current"
50                i = gblintUseTextBGCol_Current
       
60            Case "cbxFontColorUseBG_UnSeen"
70                i = gblintUseTextBGCol_UnSeen
        
        
80            Case "cbxIndicatorFill_Seen"
90                i = gblintUseShapeBG_Seen
        
100           Case "cbxIndicatorFill_Current"
110               i = gblintUseShapeBG_Current
        
120           Case "cbxIndicatorFill_UnSeen"
130               i = gblintUseShapeBG_UnSeen
                
                
140           Case "cbxSectionplacement"
150               i = gblintSectionPlacement
                  
160           Case "cbxGeneralProgressbarDirect"
170               i = gblintGeneralProgressbarDirect
                  
180           Case "cbxIgnoreIgnored"
190               i = gblintSlideIgnoreSkipped
                           
200           Case "cbxIgnoreHidden"
210               i = gblintSlideIgnoreHidden
                  
                  
220           Case Else
230               If getTheValue(control.Tag, "DefaultValue") = "1" Then
240                   bolReturn = True
250               Else
260                   bolReturn = False
270               End If
                  
                  
280       End Select
          
          'to set according to variables from theOptions
290       If i = 1 Then
300           bolReturn = True
310       Else
320           bolReturn = False
330       End If
          
End Sub

Sub GetPressedTglButton(control As IRibbonControl, ByRef pressed)
10        If gblbolDebug = True Then LogInfo "Sub GetPressedTglButton(control As IRibbonControl, ByRef pressed)"         'Debug: Write active procedure to log
          ' Callbackname in XML File "getPressed"
          ' Callback for an Access ToogleButton Control. Indicates how the control is displayed
20        Select Case control.id
        
              Case Else
30                If getTheValue(control.Tag, "DefaultValue") = "1" Then
40                    pressed = True
50                Else
60                    pressed = False
70                End If
80        End Select
End Sub

Sub GetScreentip(control As IRibbonControl, ByRef screentip)
10        If gblbolDebug = True Then LogInfo "Sub GetScreentip(control As IRibbonControl, ByRef screentip)"         'Debug: Write active procedure to log
          ' Callbackname in XML File "getScreentip"
          ' To set the property "screentip" to a Ribbon Control
20        Select Case control.id
              
              Case Else
30                screentip = "*getScreentip*"
40        End Select
End Sub

Sub GetSelectedItemIndexDropDown(control As IRibbonControl, ByRef index)
10        If gblbolDebug = True Then LogInfo "Sub GetSelectedItemIndexDropDown(control As IRibbonControl, ByRef index)"         'Debug: Write active procedure to log
          ' Callbackname in XML File "GetSelectedItemIndexDropDown"
          
          ' Callback getSelectedItemIndex (DropDown)
          
          Dim varIndex As Variant
20        varIndex = getTheValue(control.Tag, "DefaultValue")
          
          ''debug for breaking
          'MsgBox "indicatortype 1"
          'MsgBox "indicatortype 2"
          
30        If IsNumeric(varIndex) Then
40            Select Case control.id
                  
                  Case "ddcIndicatorType"
                      ' In Tab:   tabProgressDots
                      ' In Group: grpIndicator
                      'index = gblintIndicatorTypeID   'varIndex
50                    index = gblintIndicatorTypeID

                      
60                Case Else
70                    index = varIndex
80            End Select
90        End If
End Sub

Sub GetSelectedItemIndexGallery(control As IRibbonControl, ByRef index)
10        If gblbolDebug = True Then LogInfo "Sub GetSelectedItemIndexGallery(control As IRibbonControl, ByRef index)"         'Debug: Write active procedure to log
          ' Callbackname in XML File "GetSelectedItemIndexGallery"
          
          ' Callback getSelectedItemIndex (Gallery)
          
          Dim varIndex As Variant
20        varIndex = getTheValue(control.Tag, "DefaultValue")
          
30        If IsNumeric(varIndex) Then
40            Select Case control.id
                  
                  Case Else
50                    index = varIndex
60            End Select
70        End If
          
End Sub

Sub GetSupertip(control As IRibbonControl, ByRef supertip)
10        If gblbolDebug = True Then LogInfo "Sub GetSupertip(control As IRibbonControl, ByRef supertip)"         'Debug: Write active procedure to log
          ' Callbackname in XML File "getSupertip"
          ' To set the property "supertip" to a Ribbon Control
20        Select Case control.id
              
              Case Else
30                supertip = "*getSupertip*"
40        End Select
End Sub

'Combobox
Sub GetTextComboBox(control As IRibbonControl, ByRef strText)
10        If gblbolDebug = True Then LogInfo "Sub GetTextComboBox(control As IRibbonControl, ByRef strText)"         'Debug: Write active procedure to log
          ' Callbackname im XML File "GetTextComboBox"
          
          ' Callback getText (Combobox)
                                 
20        Select Case control.id
              
              Case "cmbFontName"
                  ' In Tab:   tabProgressDots
                  ' In Group: grpFont
30                strText = gblstrFontname    'getTheValue(control.Tag, "DefaultValue")
                                
40            Case "cmbFontSize"
50                strText = gbldblFontSize
                                
60            Case "cmbIndicatorSize"
70                strText = gbldblIndicatorSize   'getTheValue(control.Tag, "DefaultValue")
                  
80            Case "cmbMarginLeft"
90                strText = gblintMarginLeft  'getTheValue(control.Tag, "DefaultValue")
                  
100           Case "cmbMarginRight"
110               strText = gblintMarginRight 'getTheValue(control.Tag, "DefaultValue")
                  
120           Case "cmbMarginTop"
130               strText = gblintMarginTop   'getTheValue(control.Tag, "DefaultValue")

140           Case "cmbSectionSpacingSpace"
150               strText = gblintSectionSpacing  'getTheValue(control.Tag, "DefaultValue")
                                
160           Case "cmbIndicatorSpaceAdjustment"
170               strText = gblintIndicatorSpacingAdjustment
                                
180           Case "cmbIndicatorSpaceSqueeze"
190               strText = gblintIndicatorSqueeze
                                
200           Case "cmbSectionIndicatorGap"
210               strText = gblintSectionIndicatorGap
                  
220           Case Else
230               strText = getTheValue(control.Tag, "DefaultValue")
                          
240       End Select
          
End Sub

'EditBox
Sub GetTextEditBox(control As IRibbonControl, ByRef strText)
10        If gblbolDebug = True Then LogInfo "Sub GetTextEditBox(control As IRibbonControl, ByRef strText)"         'Debug: Write active procedure to log
          ' Callbackname in XML File "GetTextEditBox"
          
          ' Callback for an EditBox Control
          ' Indicates which value is to set to the control
20        Select Case control.id
              
              Case "ebxSomething"
                  ' In Tab:   tabProgressDots
                  ' In Group: grpFont
30                strText = getTheValue(control.Tag, "DefaultValue")
                              
                  
40            Case Else
50                strText = getTheValue(control.Tag, "DefaultValue")
                  
60        End Select
          
End Sub

Public Function getTheValue(strTag As String, strValue As String) As String
          ' *************************************************************
          ' Erstellt von     : Avenius
          ' Parameter        : Input String, SuchValue String
          ' Erstellungsdatum : 05.01.2008
          ' Bemerkungen      :
          ' Änderungen       :
          '
          ' Beispiel
          ' getTheValue("DefaultValue:=Test;Enabled:=0;Visible:=1", "DefaultValue")
          ' Return           : "Test"
          ' *************************************************************
          
10        On Error Resume Next
            
20        If gblbolDebug = True Then LogInfo "Public Function getTheValue(strTag As String, strValue As String) As String"         'Debug: Write active procedure to log
            
          Dim workTb()     As String
          Dim Ele()        As String
          Dim myVariabs()  As String
          Dim i            As Integer
30        workTb = Split(strTag, ";")
            
40        ReDim myVariabs(LBound(workTb) To UBound(workTb), 0 To 1)
50        For i = LBound(workTb) To UBound(workTb)
60            Ele = Split(workTb(i), ":=")
70            myVariabs(i, 0) = Ele(0)
80            If UBound(Ele) = 1 Then
90                myVariabs(i, 1) = Ele(1)
100           End If
110       Next
            
120       For i = LBound(myVariabs) To UBound(myVariabs)
130           If strValue = myVariabs(i, 0) Then
140               getTheValue = myVariabs(i, 1)
150           End If
160       Next
            
End Function

Sub GetTitle(control As IRibbonControl, ByRef title)
10        If gblbolDebug = True Then LogInfo "Sub GetTitle(control As IRibbonControl, ByRef title)"         'Debug: Write active procedure to log
          ' Callbackname in XML File "getTitle"
          ' To set the property "title" to a Ribbon MenuSeparator Control
20        Select Case control.id
              
              Case Else
30                title = "*getTitle*"
40        End Select
End Sub

Public Sub GetVisible(control As IRibbonControl, ByRef visible)
10        If gblbolDebug = True Then LogInfo "Public Sub GetVisible(control As IRibbonControl, ByRef visible)"         'Debug: Write active procedure to log
          ' Callbackname in XML File "getVisible"
          
          ' To set the property "visible" to a Ribbon Control
          ' For further information see: http://www.accessribbon.de/en/index.php?Downloads:12
          ' Setzen der Visible Eigenschaft eines Ribbon Controls
          ' Weitere Informationen: http://www.accessribbon.de/index.php?Downloads:12
20        Select Case control.id
              
              Case "ddcGeneralProgressbarEvenSpace"
30                visible = True
              
40            Case "ddcIgnoreIndicator"
50                visible = True
              
60            Case "ddcIgnoreSections"
70                visible = True
              
80            Case "btnGeneralProgressbarCreate"
                  ' Button
                  ' In Tab:   tabProgressDots
                  ' In Group: grpGeneral
90                visible = True
                  
100           Case "btnGeneralProgressbarDelete"
                  ' Button
                  ' In Tab:   tabProgressDots
                  ' In Group: grpGeneral
110               visible = True
                  
120           Case "btnGeneralProgressbarHelp"
                  ' Button
                  ' In Tab:   tabProgressDots
                  ' In Group: grpGeneral
130               visible = True
                  
140           Case "grpGeneral"
                  ' Group:    grpGeneral
                  ' In Tab:   tabProgressDots
150               visible = True
                  
160           Case "btnGeneralProgressbarHideTrue"
                  ' Button
                  ' In Tab:   tabProgressDots
                  ' In Group: grpSlide
170               visible = True
                  
180           Case "btnGeneralProgressbarHideFalse"
                  ' Button
                  ' In Tab:   tabProgressDots
                  ' In Group: grpSlide
190               visible = True
                  
200           Case "grpSlide"
                  ' Group:    grpSlide
                  ' In Tab:   tabProgressDots
210               visible = True
                  
220           Case "cmbFontName"
                  ' ComboBox
                  ' In Tab:   tabProgressDots
                  ' In Group: grpFont
230               visible = True
                  
                 
240           Case "grpFont"
                  ' Group:    grpFont
                  ' In Tab:   tabProgressDots
250               visible = True
                  
260           Case "ddcIndicatorType"
                  ' Dropdown
                  ' In Tab:   tabProgressDots
                  ' In Group: grpIndicator
270               visible = True
                    
280           Case "grpIndicator"
                  ' Group:    grpIndicator
                  ' In Tab:   tabProgressDots
290               visible = True
                  
300           Case "grpMargin"
                  ' Group:    grpMargin
                  ' In Tab:   tabProgressDots
310               visible = True
                  
320           Case "grpSpacing"
                  ' Group:    grpSpacing
                  ' In Tab:   tabProgressDots
330               visible = True
          
340           Case "cbxFontColorUseBG"
350               visible = True
                  
360           Case "cbxFontColorUseBG_Current"
370               visible = True
       
380           Case "cbxFontColorUseBG_UnSeen"
390               visible = True
        
        
400           Case "cbxIndicatorFill_Seen"
410               visible = True
        
420           Case "cbxIndicatorFill_Current"
430               visible = True
        
440           Case "cbxIndicatorFill_UnSeen"
450               visible = True
                                         
                  
460           Case "cbxSectionplacement"
470               visible = True
                  
480           Case "cbxGeneralProgressbarDirect"
490               visible = True
                  
500           Case "cbxIgnoreIgnored"
                  ' Checkbox
                  ' In Tab:   tabProgressDots
                  ' In Group: grpIgnore
510               visible = True
                  
520           Case "cbxIgnoreHidden"
                  ' Checkbox
                  ' In Tab:   tabProgressDots
                  ' In Group: grpIgnore
530               visible = True
                                   
540           Case "grpIgnore"
                  ' Group:    grpIgnore
                  ' In Tab:   tabProgressDots
550               visible = True
                  
560           Case "tabProgressDots"
                  ' Tab:   tabProgressDots
570               visible = True
                  
580           Case Else
590               visible = True
600       End Select
End Sub

' Helper Function
' Hilfsfunktionen
Public Function getXMLForDynamicMenu() As String
10        If gblbolDebug = True Then LogInfo "Public Function getXMLForDynamicMenu() As String"         'Debug: Write active procedure to log
          ' Creates a XML String for DynamicMenu CallBack - getContent
          
          ' Erstellt den Inhalt fuer das DynamicMenu im Callback getContent
          
          Dim lngDummy    As Long
          Dim strDummy    As String
          Dim strContent  As String
          
          Dim Items(4) As ItemsVal
20        Items(0).id = "btnDy1"
30        Items(0).label = "Item 1"
40        Items(0).imageMso = "_1"
50        Items(1).id = "btnDy2"
60        Items(1).label = "Item 2"
70        Items(1).imageMso = "_2"
80        Items(2).id = "btnDy3"
90        Items(2).label = "Item 3"
100       Items(2).imageMso = "_3"
110       Items(3).id = "btnDy4"
120       Items(3).label = "Item 4"
130       Items(3).imageMso = "_4"
140       Items(4).id = "btnDy5"
150       Items(4).label = "Item 5"
160       Items(4).imageMso = "_5"
          
170       strDummy = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbCrLf
          
180       For lngDummy = LBound(Items) To UBound(Items)
190           strContent = strContent & _
                  "<button id=""" & Items(lngDummy).id & """" & _
                  " label=""" & Items(lngDummy).label & """" & _
                  " imageMso=""" & Items(lngDummy).imageMso & """" & _
                  " onAction=""OnActionButton""/>" & vbCrLf
200       Next
       
210       strDummy = strDummy & strContent & "</menu>"
220       getXMLForDynamicMenu = strDummy
End Function

Public Sub OnActionButton(control As IRibbonControl)
          If gblbIndicatorUpdateRunning Then Exit Sub
10        If gblbolDebug = True Then LogInfo "Public Sub OnActionButton(control As IRibbonControl)"         'Debug: Write active procedure to log
          'Callback in XML File "onAction"
          ' Callback for event button click
          
          Dim intSubFrmPositionLeft  As Integer
          Dim strRefresh As String
          Dim strRGBArray() As String
          
20        On Error GoTo ErrorHandler
          
30        strRefresh = ""   'if empty, no refresh of the PI needed
          
40        Select Case control.id
                                                       
              '---General--------------------------------
              Case "btnGeneralProgressbarCreate"
                  PI_ClearGenerationScope
50                strRefresh = "btnGeneralProgressbarCreate"

                  'gbliDrawStart = 1
                  'gbliDrawEnd = 999

60                gblCreatePressed = True
                  'RefreshProgressBar
            
70            Case "btnGeneralProgressbarDelete"
                  PI_ClearGenerationScope
80                strRefresh = ""
90                DelProgressBar
                  'Beep
                                
100           Case "btnGeneralProgressbarHelp"
110               InitializePIVersion
120               frmHelp.lblVersion.Caption = gblProgramVersion
130               frmHelp.cbxDebugMode = gblbolDebug
140               frmHelp.lblSectionNameJoker.Caption = gblstrSectionNameHidden
150               frmHelp.lblSectionNameHidden.Caption = gblstrPISectionHidden
160               frmHelp.Show
        
                  '---Fonts--------------------------------
170           Case "btnFontColorSeen"
180               strRefresh = "Color Font Seen"
190               intSubFrmPositionLeft = 250
            
200           Case "btnFontColorUnSeen"
210               strRefresh = "Color Font Unseen"
220               intSubFrmPositionLeft = 300
                        
230           Case "btnFontColorSeenBG"
240               strRefresh = "Color Font Seen Background"
250               If gblintUseTextBGCol = 0 Then
260                   Beep
270               End If
            
280           Case "btnFontColorUnSeenBG"
290               strRefresh = "Color Font Unseen Background"
300               If gblintUseTextBGCol_UnSeen = 0 Then
310                   Beep
320               End If

330           Case "btnFontColorCurrentBG"
340               strRefresh = "Color Font Active Background"
350               If gblintUseTextBGCol_Current = 0 Then
360                   Beep
370               End If
                        
380           Case "btnFontColorCurrent"
390               strRefresh = "Color Font Active"
400               If gblintUseTextBGCol_Current = 0 Then
410                   Beep
420               End If
                        
                  '---Indicator--------------------------------
            
430           Case "btnIndicatorColorSeen"
440               strRefresh = "Color Indicator Seen"
450               intSubFrmPositionLeft = 600
460               If gblintUseShapeBG_Seen = 0 Then
470                   Beep
480               End If
                       
490           Case "btn_OutlinePlus"
500               strRefresh = "btn_OutlinePlus"
510               LneThickness True
520               optionsFromRibbon control.id, CStr(gblShapeLineWeight)
        
530           Case "btn_OutlineMinus"
540               strRefresh = "btn_OutlineMinus"
550               LneThickness False
560               optionsFromRibbon control.id, CStr(gblShapeLineWeight)

570           Case "btn_Style"
580               strRefresh = "btn_Style"
590               LneStyle True
600               optionsFromRibbon control.id, CStr(gblShapeLineStyle)
            
610           Case "btn_Dash"
620               strRefresh = "btn_Dash"
630               LneDashStyle True
640               optionsFromRibbon control.id, CStr(gblShapeLineDashStyle)
           
650           Case "btnIndicatorColorUnSeen"
660               strRefresh = "Color Indicator Unseen"
670               intSubFrmPositionLeft = 650
680               If gblintUseShapeBG_UnSeen = 0 Then
690                   Beep
700               End If

710           Case "btnIndicatorColorCurrent"
720               strRefresh = "Color Indicator Active"
730               intSubFrmPositionLeft = 700
740               If gblintUseShapeBG_Current = 0 Then
750                   Beep
760               End If

                  '---Ignore--------------------------------
770           Case "btnGeneralProgressbarHideTrue"
                  'this button just for testing new features directly
780               RemoveIgnoreSlide (0)
790               AddIgnoreSlide
800               strRefresh = "btnGeneralProgressbarHideTrue"
            
810           Case "btnGeneralProgressbarRepeat"
820               RemoveIgnoreSlide (0)
830               AddRepeatSlide
840               strRefresh = "btnGeneralProgressbarRepeat"
            
850           Case "btnGeneralProgressbarHideFalse"
860               RemoveIgnoreSlide (1)
870               strRefresh = "btnGeneralProgressbarHideFalse"


880           Case "btnGeneralProgressbarPreview"
                  PI_ClearGenerationScope
890               strRefresh = "btnGeneralProgressbarPreview"
900               gblCreatePressed = True
910               gblShowPreview = 1
                                
                  '--- Settings Export to file / Import from file ------------------
920           Case "btnSettingsExport"
930               strRefresh = ""
940               gblCreatePressed = False
950               SettingsExport
                                
960           Case "btnSettingsImport"
970               strRefresh = ""
980               gblCreatePressed = False
990               SettingsImport
          
1000          Case Else
1010              MsgBox "Button """ & control.id & """ clicked" & vbCrLf & _
                      "Es wurde auf Button """ & control.id & """ in Ribbon geklickt", _
                      vbInformation
1020      End Select
          
          '---Color ... all --------------------------------
1030      If strRefresh <> "" Then
1040          If Left(strRefresh, 6) = "Color " Then
1050              frmThemeColors.lblColorSelected.Tag = "New color for " & strRefresh
                  frmThemeColors.lblColorSelected.Caption = PI_Zh("65B0989C8272")
1060              frmThemeColors.Caption = PI_ColorTitle(strRefresh)
                  'at this time still frmThemeColors property still centered
1070              frmThemeColors.Left = intSubFrmPositionLeft
1080              frmThemeColors.Top = 113
1090              CalcColor
            
                  'set active color to color chooser
1100              Select Case strRefresh
                      Case "Color Font Seen"
1110                      strRGBArray = Split(gblstrFontColorSeen, ";")
                
1120                  Case "Color Font Unseen"
1130                      strRGBArray = Split(gblstrFontColorUnSeen, ";")
                
1140                  Case "Color Font Seen Background"
1150                      strRGBArray = Split(gblstrFontBGColorSeen, ";")
                
1160                  Case "Color Font Unseen Background"
1170                      strRGBArray = Split(gblstrFontBGColorUnSeen, ";")
                
1180                  Case "Color Font Active Background"
1190                      strRGBArray = Split(gblstrFontBGColorCurrent, ";")
                
1200                  Case "Color Font Active"
1210                      strRGBArray = Split(gblstrFontColorCurrent, ";")
        
1220                  Case "Color Indicator Seen"
1230                      strRGBArray = Split(gblstrIndicatorColorSeen, ";")
                
1240                  Case "Color Indicator Unseen"
1250                      strRGBArray = Split(gblstrIndicatorColorUnSeen, ";")

1260                  Case "Color Indicator Active"
1270                      strRGBArray = Split(gblstrIndicatorColorCurrent, ";")

1280              End Select
            
1290              frmThemeColors.lblColorOld.BackColor = RGB(Val(strRGBArray(0)), Val(strRGBArray(1)), Val(strRGBArray(2)))
1300              frmThemeColors.lblColorSelected.BackColor = RGB(Val(strRGBArray(0)), Val(strRGBArray(1)), Val(strRGBArray(2)))
            
1310              frmThemeColors.dotSeenRt = Val(strRGBArray(0))
1320              frmThemeColors.dotSeenGt = Val(strRGBArray(1))
1330              frmThemeColors.dotSeenBt = Val(strRGBArray(2))
            
1340              frmThemeColors.Show
            
1350          End If

1360          RefreshProgressBar
1370      Else
              'no need to do anything here
1380      End If
          
exit_proc:
1390      Exit Sub
          
ErrorHandler:
1400      If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
1410      Select Case Err.Number

              Case -2147188160      'no active presentation
1420              MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "OnActionButton: Nothing to do!"
1430              Resume exit_proc
1440              Resume

1450          Case 9
1460              Auto_Open   'reinitialize!
            
1470          Case Else
1480              MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "OnActionButton"
1490              Resume Next
1500              Resume
        
1510      End Select
End Sub

'Command Button
Sub OnActionButtonHelp(control As IRibbonControl, ByRef CancelDefault)
10        If gblbolDebug = True Then LogInfo "Sub OnActionButtonHelp(control As IRibbonControl, ByRef CancelDefault)"         'Debug: Write active procedure to log
          ' Callbackname in XML File Command "onAction"
          ' Callback for command event button click
          ' Callback fuer Command Button Click
          
20        InitializePIVersion
30        frmHelp.Show
          
          'MsgBox "Button ""Help"" clicked" & vbCrLf & _
           "Es wurde auf Button ""Hilfe"" geklickt", _
           vbInformation
40        CancelDefault = True
End Sub

Sub OnActionCheckBox(control As IRibbonControl, pressed As Boolean)
          If gblbIndicatorUpdateRunning Then Exit Sub
          Dim i As Integer
          ' Callbackname in XML File "OnActionCheckBox"
          ' Callback for event checkbox click
          
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Sub OnActionCheckBox(control As IRibbonControl, pressed As Boolean)"         'Debug: Write active procedure to log
          
          'Always save active value to setting directly! No further action here required
30        If pressed = True Then
40            i = 1
50        Else
60            i = 0
70        End If
80        optionsFromRibbon control.id, CStr(i)
          
90        Select Case control.id
          
              Case "cbxFontColorUseBG"
                  'value assigned in optionsFromRibbon!
100           Case "cbxFontColorUseBG_Current"
                  'value assigned in optionsFromRibbon!
110           Case "cbxFontColorUseBG_UnSeen"
                  'value assigned in optionsFromRibbon!
              
120           Case "cbxIndicatorFill_Seen"
                  'value assigned in optionsFromRibbon!
130           Case "cbxIndicatorFill_Current"
                  'value assigned in optionsFromRibbon!
140           Case "cbxIndicatorFill_UnSeen"
                  'value assigned in optionsFromRibbon!
              
150           Case "cbxSectionplacement"
                  'value assigned in optionsFromRibbon!
          
160           Case "cbxGeneralProgressbarDirect"
                  'value assigned in optionsFromRibbon!
                  
170           Case "cbxIgnoreIgnored"
                  'gblintSlideIgnoreSkipped = pressed            'value assigned in optionsFromRibbon!
                           
180           Case "cbxIgnoreHidden"
                  'gblintSlideIgnoreHidden = pressed            'value assigned in optionsFromRibbon!
                                           
190           Case Else
200               MsgBox "The Value of the Checkbox """ & control.id & """ is: " & pressed & vbCrLf & _
                      "Der Wert der Checkbox """ & control.id & """ ist: " & pressed, _
                      vbInformation
210       End Select
          
220       RefreshProgressBar
          
exit_proc:
230       Exit Sub
          
ErrorHandler:
240       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
250       Select Case Err.Number
              Case -2147188160      'no active presentation
260               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "OnActionCheckBox: Nothing to do!"
270               Resume exit_proc
280               Resume
290           Case Else
300               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "OnActionCheckBox"
310               Resume Next
320               Resume
              
330       End Select
          
End Sub

'DropDown
Sub OnActionDropDown(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
          If gblbIndicatorUpdateRunning Then Exit Sub
          
          ' Callbackname in XML File "OnActionDropDown"
          ' Callback onAction (DropDown)
              
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Sub OnActionDropDown(control As IRibbonControl, selectedId As String, selectedIndex As Integer)"         'Debug: Write active procedure to log
          

30        Select Case control.id
              
              Case "ddcIgnoreIndicator" ' None|All|Inactive Sections
                  ' In Group: grpLimitVisibility
40                Select Case selectedId
                      Case "ddcIgnoreIndicatorItem0"
50                        gblstrSlideIgnoreIndicator = "All"
60                    Case "ddcIgnoreIndicatorItem1"
70                        gblstrSlideIgnoreIndicator = "Active Section"
80                    Case "ddcIgnoreIndicatorItem2"
90                        gblstrSlideIgnoreIndicator = "None"
100               End Select
110               optionsFromRibbon "ddcIgnoreIndicator", gblstrSlideIgnoreIndicator
              
              
120           Case "ddcIgnoreSections" 'None|Text|Full
                  ' In Group: grpLimitVisibility
130               Select Case selectedId
                      Case "ddcIgnoreSectionsItem0"
140                       gblstrSlideIgnoreSections = "Full"
150                   Case "ddcIgnoreSectionsItem1"
160                       gblstrSlideIgnoreSections = "Space Only"
170                   Case "ddcIgnoreSectionsItem2"
180                       gblstrSlideIgnoreSections = "None"
190               End Select
200               optionsFromRibbon "ddcIgnoreSections", gblstrSlideIgnoreSections
              
              
210           Case "ddcGeneralProgressbarEvenSpace" ' Shapes or Sections: Which should be used for width calculation
                  ' In Group: grpGeneral
                  
                  'If CInt(Right(selectedId, Len(selectedId) - 34)) = 0 Then    ' = Shapes
220               If selectedId = "ddcGeneralProgressbarEvenSpaceItem0" Then
230                   gblEvenSpaceSecOrInd = "Shapes"
240               Else
250                   gblEvenSpaceSecOrInd = "Sections"
260               End If
270               optionsFromRibbon "ddcGeneralProgressbarEvenSpace", gblEvenSpaceSecOrInd
              
              
280           Case "ddcIndicatorType"
                  ' In Tab:   tabProgressDots
                  ' In Group: grpIndicator
                  'gblintIndicatorTypeID = CInt(Right(selectedId, Len(selectedId) - 20))      'will be set in optionsFromRibbon

290               If retShapeID(CInt(Right(selectedId, Len(selectedId) - 20))) = "9991" Then     'custom image
300                   If gblbolDebug = True Then LogInfo "Opening frmCustomImages"         'Debug: Write active procedure to log
310                   frmCustomImages.txtImageSeen = gblCustomImage_Seen
320                   frmCustomImages.txtImageActive = gblCustomImage_Active
330                   frmCustomImages.txtImageUnSeen = gblCustomImage_UnSeen
340                   frmCustomImages.chkUseCustomShapes = False
350                   frmCustomImages.Show
360               Else
370                   gblshowCustomImage = 0
380                   optionsFromRibbon "ddcIndicatorType", Right(selectedId, Len(selectedId) - 20)
390               End If
                  
                  
400           Case "ddcGeneralProgressbarStyle"   ' Style: What to see: Full (Indicator and Sections), Indicator only, Sections only
                  ' In Group: grpGeneral
410               Select Case selectedId
                      Case "ddcGeneralProgressbarStyleItem0"
420                       gblstrIndicatorStyle = "Full"
430                   Case "ddcGeneralProgressbarStyleItem1"
440                       gblstrIndicatorStyle = "Indicator only"
450                   Case "ddcGeneralProgressbarStyleItem2"
460                       gblstrIndicatorStyle = "Sections only"
470               End Select
480               optionsFromRibbon "ddcGeneralProgressbarStyle", gblstrIndicatorStyle
                  
                  
490           Case Else
500               MsgBox "The selected ItemID of DropDown-Control """ & control.id & """ is : """ & selectedId & """" & vbCrLf & _
                      "Die selektierte ItemID des DropDown-Control """ & control.id & """ ist : """ & selectedId & """", _
                      vbInformation
510       End Select
          
520       RefreshProgressBar
          
          
exit_proc:
530       Exit Sub
          
ErrorHandler:
540       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
550       Select Case Err.Number
              Case -2147188160      'no active presentation
560               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "OnActionDropDown: Nothing to do!"
570               Resume exit_proc
580               Resume
590           Case Else
600               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "OnActionDropDown"
610               Resume Next
620               Resume
              
630       End Select
End Sub

'Gallery
Sub OnActionGallery(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
          If gblbIndicatorUpdateRunning Then Exit Sub
10        If gblbolDebug = True Then LogInfo "Sub OnActionGallery(control As IRibbonControl, selectedId As String, selectedIndex As Integer)"         'Debug: Write active procedure to log
          ' Callbackname in XML File "OnActionGallery"
          
          ' Callback onAction (Gallery)
          
20        Select Case control.id

              Case Else
30                Select Case selectedId
                      Case Else
40                        MsgBox "The selected ItemID of Gallery-Control """ & control.id & """ is : """ & selectedId & """" & vbCrLf & _
                              "Die selektierte ItemID des Gallery-Control """ & control.id & """ ist : """ & selectedId & """", _
                              vbInformation
50                End Select
60        End Select
          
End Sub

Sub OnActionTglButton(control As IRibbonControl, pressed As Boolean)
          If gblbIndicatorUpdateRunning Then Exit Sub
10        If gblbolDebug = True Then LogInfo "Sub OnActionTglButton(control As IRibbonControl, pressed As Boolean)"         'Debug: Write active procedure to log
          ' Callbackname in XML File "onAction"
          ' Callback for a Toggle Buttons click event
          
20        Select Case control.id
                                     
              Case Else
30                MsgBox "The Value of the Toggle Button """ & control.id & """ is: " & pressed & vbCrLf & _
                      "Der Wert der Toggle Button """ & control.id & """ ist: " & pressed, _
                      vbInformation
40        End Select
End Sub

Sub OnChangeComboBox(control As IRibbonControl, strText As String)
          If gblbIndicatorUpdateRunning Then Exit Sub
          ' Callbackname im XML File "OnChangeCombobox"
          ' Callback onChange (Combobox)
         
10        On Error GoTo ErrorHandler
          
          Dim theOptions As ProgressDotsOptions   'to be able to write back the options to the ribbon in case "...more..." in Fontname was chosen
              
20        If gblbolDebug = True Then LogInfo "Sub OnChangeComboBox(control As IRibbonControl, strText As String)"         'Debug: Write active procedure to log
         
30        Select Case control.id
              
              Case "cmbFontName"
                  If strText = PI_Zh("66F4591A5B574F53002E002E002E") Then strText = gblstrMessageMore
                  ' In Tab:   tabProgressDots
                  ' In Group: grpFont
                  '            MsgBox "The Value of the Combobox """ & control.id & """ is: " & strText & vbCrLf & _
                               "Der Wert der Combobox """ & control.id & """ ist: " & strText, _
                               vbInformation
                      
40                If strText = gblstrMessageMore Then
50                    gblFontnameRibbonActivated = gblstrMessageMore
60                    ChooseFromAllFonts (strText)
70                    strText = gblFontnameRibbonActivated
80                End If
90                If strText = gblstrMessageMore Then  'if still ...more... then set back to default!
100                   strText = "Calibri"
110                   gobjRibbon.InvalidateControl "cmbFontName"   'might not bee needed
120                   optionsToRibbon theOptions    'set theoptions to global variables (from ribbon)
130               End If
140               optionsFromRibbon "cmbFontName", strText

150           Case "cmbFontSize"
160               optionsFromRibbon "cmbFontSize", strText

170           Case "cmbIndicatorSize"
180               optionsFromRibbon "cmbIndicatorSize", strText
                  
190           Case "cmbMarginLeft"
200               optionsFromRibbon "cmbMarginLeft", strText
                  
210           Case "cmbMarginRight"
220               optionsFromRibbon "cmbMarginRight", strText
                  
230           Case "cmbMarginTop"
240               optionsFromRibbon "cmbMarginTop", strText

250           Case "cmbSectionSpacingSpace"
260               optionsFromRibbon "cmbSectionSpacingSpace", strText
                  
270           Case "cmbIndicatorSpaceAdjustment"
280               optionsFromRibbon "cmbIndicatorSpaceAdjustment", strText
                  
290           Case "cmbIndicatorSpaceSqueeze"
300               optionsFromRibbon "cmbIndicatorSpaceSqueeze", strText
                  
310           Case "cmbSectionIndicatorGap"
320               optionsFromRibbon "cmbSectionIndicatorGap", strText
                  
                  
330           Case "cmbIndicatorTransparencySeen"
340               optionsFromRibbon "cmbIndicatorTransparencySeen", strText
                  
350           Case "cmbIndicatorTransparencyActive"
360               optionsFromRibbon "cmbIndicatorTransparencyActive", strText
               
370           Case "cmbIndicatorTransparencyUnseen"
380               optionsFromRibbon "cmbIndicatorTransparencyUnseen", strText
                  
                  
390           Case "cmbFontTransparencySeen"
400               optionsFromRibbon "cmbFontTransparencySeen", strText
                  
410           Case "cmbFontTransparencyActive"
420               optionsFromRibbon "cmbFontTransparencyActive", strText
               
430           Case "cmbFontTransparencyUnseen"
440               optionsFromRibbon "cmbFontTransparencyUnseen", strText
                  

450           Case Else
460               MsgBox "The selected Item of Combobox-Control """ & control.id & """ is : """ & strText & """" & vbCrLf & _
                      "Das selektierte Item des Combobox-Control """ & control.id & """ ist : """ & strText & """", _
                      vbInformation
470       End Select
480       RefreshProgressBar
          
exit_proc:
490       Exit Sub
          
ErrorHandler:
500       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
510       Select Case Err.Number
              Case -2147188160      'no active presentation
520               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "OnChangeComboBox: Nothing to do!"
530               Resume exit_proc
540               Resume
550           Case 91 'can not invalidate control
560               Resume Next
                  
570           Case Else
580               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "OnChangeComboBox"
590               Resume Next
600               Resume
              
610       End Select
End Sub

Sub OnChangeEditBox(control As IRibbonControl, strText As String)
          If gblbIndicatorUpdateRunning Then Exit Sub
          Dim intCorrected As Integer
10        If gblbolDebug = True Then LogInfo "Sub OnChangeEditBox(control As IRibbonControl, strText As String)"         'Debug: Write active procedure to log
          ' Callbackname in XML File "OnChangeEditBox"
          
          ' Callback Editbox: Rückgabewert der Editbox
          ' Callback Editbox: Return value of the Editbox
20        Select Case control.id
              
              Case "ebxSomename"
                  ' In Tab:   tabProgressDots
                  ' In Group: grpFont
30                MsgBox "do something"
                                                
40            Case "ebx_DrawStart"
50                If IsNumeric(strText) Then
60                    If Val(strText) > 0 And Val(strText) < 1000 Then
                          'all ok
70                    Else
80                        intCorrected = 1
90                    End If
100               Else
110                   intCorrected = 1
120               End If
130               If intCorrected = 1 Then strText = "1"
140               gbliDrawStart = Val(strText)


150           Case "ebx_DrawEnd"
160               If IsNumeric(strText) Then
170                   If Val(strText) > 0 And Val(strText) < 1000 Then
                          'all ok
180                   Else
190                       intCorrected = 1
200                   End If
210               Else
220                   intCorrected = 1
230               End If
240               If intCorrected = 1 Then strText = "999"
250               gbliDrawEnd = Val(strText)

                         
260           Case Else
270               MsgBox "The Value of the EditBox """ & control.id & """ is: " & strText & vbCrLf & _
                      "Der Wert der EditBox """ & control.id & """ ist: " & strText, _
                      vbInformation
280       End Select
          
290       optionsFromRibbon control.id, strText
300       RefreshProgressBar
End Sub

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
10        If gblbolDebug = True Then LogInfo "Public Sub OnRibbonLoad(ribbon As IRibbonUI)"         'Debug: Write active procedure to log
          'Callbackname in XML File "onLoad"

20        Set gobjRibbon = ribbon

          'gobjRibbon.InvalidateControl "ebxFontColorSeen" 'requery field in form
          'gobjRibbon.Invalidate
              
30        Auto_Open

End Sub

Public Sub LneThickness(bDirection As Boolean)
          If gblbIndicatorUpdateRunning Then Exit Sub
          Dim intAdjustment As Integer
10        intAdjustment = 1
20        If bDirection = False Then
30            intAdjustment = intAdjustment * -1
40        End If
50        gblShapeLineWeight = gblShapeLineWeight + intAdjustment

60        If gblShapeLineWeight < 0 Then gblShapeLineWeight = gbldblIndicatorSize
70        If gblShapeLineWeight > gbldblIndicatorSize Then gblShapeLineWeight = 0
          'If gblShapeLineWeight = 0 Then gblShapeLineWeight = 1
End Sub

Public Sub LneStyle(bDirection As Boolean)
          If gblbIndicatorUpdateRunning Then Exit Sub
          Dim intAdjustment As Integer
10        intAdjustment = 1
20        If bDirection = False Then
30            intAdjustment = intAdjustment * -1
40        End If
50        gblShapeLineStyle = gblShapeLineStyle + intAdjustment

60        If gblShapeLineStyle < 1 Then gblShapeLineStyle = 5
70        If gblShapeLineStyle > 5 Then gblShapeLineStyle = 1
          '-2 would also be possible:  msoShapeLineStyleMixed
End Sub

Public Sub LneDashStyle(bDirection As Boolean)
          If gblbIndicatorUpdateRunning Then Exit Sub
          Dim intAdjustment As Integer
10        intAdjustment = 1
20        If bDirection = False Then
30            intAdjustment = intAdjustment * -1
40        End If
50        gblShapeLineDashStyle = gblShapeLineDashStyle + intAdjustment

60        If gblShapeLineDashStyle < 1 Then gblShapeLineDashStyle = 5
70        If gblShapeLineDashStyle > 5 Then gblShapeLineDashStyle = 1
          '-2 would also be possible:  msoShapeLineDashStyleMixed
End Sub
