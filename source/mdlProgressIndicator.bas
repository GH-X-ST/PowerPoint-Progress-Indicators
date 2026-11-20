Attribute VB_Name = "mdlProgressIndicator"
' ------------------------------------------------------
' Progress Indicator
' Version see variable gblProgramVersion!
'
' (C) Copyright 2017-2025 Olaf Nöhring ( https://datenbank-projekt.de ), 2007 Edward Lopr
'
' Adjusted for Office 2019 (32 and 64bit, should work for Office 2010, 2013 and 2016 as well) by Olaf Nöhring 2017-2021
' Complete MAKEOVER:
'       lots and lots of new features
'       Ribbon
'       implementation that works on Windows AND MAC OS!
'
'
' Drop me a line to say thanks and/or suggest features
'
'
' This macro adds a ribbon to PowerPoint that can be used to create a
' "progress indicator" for your presentation.  A series of shapes is drawn across
' the top border, corresponding to slides; and these shapes change color
' as you advance through the presentation.
' You can exclude specific slides, hidden slides from getting progress indicators.
' If your slides have sections (see https://support.office.com/en-us/article/Organize-your-slides-into-sections-DE4BF162-E9CC-4F58-B64A-7AB09443B9F8 )
' the slides can be grouped by those to indicate the overall structure of your talk. You can decide to ignore sections and display just the indicator
' See the HELP button in the ribbon as well!
' E.g., the
' generated progress bar could look something like this:
'
'         Intro     Topic 1        Topic 2      Conclusions
'        # # # #    # # * *    * * * * * * * * *    * *
'
' Everytime you change something in the ribbon, the progress indicator can be updated accordingly.
' Settings from the ribbon will be saved as a tag to your presentation (invisible) and be read next time you open your presentation.
' You can remove the progress indicators as well using the corresponding button.
'
' Use the code - don't mind about it's quality as I tried to reuse as much of old code as possible.
'
'
'
'
' Developer Information / Reminder
' In Code already present and saved to settings, but unused in ribbon yet:
' Combobox with ID             cmbIndicatorSpaceAdjustment
'
'
'
'                                License:
' =========================================================================
' Permission is hereby granted, free of charge, to any person obtaining
' a copy of this software and any associated documentation files (the
' "Software"), to deal in the Software without restriction, including
' without limitation the rights to use, copy, modify, merge, publish,
' distribute, sublicense, and/or sell copies of the Software, and to
' permit persons to whom the Software is furnished to do so, subject to
' the following conditions:
'
' The above copyright notice and this permission notice shall be included
' in all copies or substantial portions of the Software.
'
' The software is provided "as is", without warranty of any kind, express
' or implied, including but not limited to the warranties of
' merchantability, fitness for a particular purpose and noninfringement.
' In no event shall the authors or copyright holders be liable for any
' claim, damages or other liability, whether in an action of contract, tort
' or otherwise, arising from, out of or in connection with the software or
' the use or other dealings in the software.
' =========================================================================

Option Explicit

Type ColorType
    red As Integer
    green As Integer
    blue As Integer
End Type

Type ProgressDotsOptions

    fontName As String              'gblstrFontname
    textHeight As Double            'gbldblFontSize
    seenTextColor As ColorType      'gblstrFontColorSeen
    unseenTextColor As ColorType    'gblstrFontColorUnSeen

    dotShape As String              'gblstrIndicatorType
    dotShapeID As Integer           'gblintIndicatorTypeID
    dotSize As Double               'gbldblIndicatorSize
    seenDotColor As ColorType       'gblstrIndicatorColorSeen
    unseenDotColor As ColorType     'gblstrIndicatorColorUnSeen

    LeftMargin As Double            'gblintMarginLeft
    RightMargin As Double           'gblintMarginRight
    yMargin As Double               'gblintMarginTop

    SectionGap As Double            'gblintSectionSpacing
    IndicatorSpaceAdjustment As Double      'gblintIndicatorSpacingAdjustment
    IndicatorSqueeze As Double      'gblintIndicatorSqueeze
    
    ignoreSkipped As Integer        'gblintSlideIgnoreSkipped
    ignoreHidden As Integer         'gblintSlideIgnoreHidden
    ignoreSections As String       'gblstrSlideIgnoreSections   'before 1.92 was integer

    DirectUpdate As Integer         'gblintGeneralProgressbarDirect

    useTextColorBG  As Integer         '0= do not use bg color, 1= use bg color
    useTextColorBG_Current  As Integer
    useTextColorBG_UnSeen   As Integer
    useShapeColorBG_Seen     As Integer
    useShapeColorBG_Current As Integer
    useShapeColorBG_UnSeen  As Integer
    
    seenTextColorBG    As ColorType    'gblstrFontBGColorSeen
    unseenTextColorBG  As ColorType    'gblstrFontBGColorUnseen
    
    SectionPlacement   As Integer       'gblintSectionPlacement
    SectionIndicatorGap As Integer      'gblintSectionIndicatorGap
    
    CurrentDotColor   As ColorType       'gblstrIndicatorColorCurrent
    CurrentTextColorBG  As ColorType     'gblstrFontBGColorCurrent
    CurrentTextColor  As ColorType       'gblstrFontColorCurrent
    
    ShapeLineWeight As Integer
    ShapeLineStyle As Integer
    ShapeLineDashStyle As Integer
    
    ignoreIndicator As String      'gblstrSlideIgnoreIndicator
    
    transparencyIndicatorSeen As Integer        'gblTransparencyIndicatorSeen
    transparencyIndicatorActive As Integer      'gblTransparencyIndicatorActive
    transparencyIndicatorUnseen As Integer      'gblTransparencyIndicatorUnseen
    
    transparencyFontSeen As Integer             'gblTransparencyFontSeen
    transparencyFontActive As Integer           'gblTransparencyFontActive
    transparencyFontUnseen As Integer           'gblTransparencyFontUnseen
    
    EvenSpaceSecOrInd As String     'gblEvenSpaceSecOrInd
    PIStyle As String     'gblstrIndicatorStyle
    SectionBand As Integer
    SectionBandPadding As Double
    SectionBandContrast As Integer
    SectionBandGradient As Integer
    SectionBandTransparencyCurrent As Double
    SectionBandTransparencyLeft As Double
    SectionBandTransparencyRight As Double
    SectionBandLeftNear As ColorType
    SectionBandLeftFar As ColorType
    SectionBandRightNear As ColorType
    SectionBandRightFar As ColorType
    SectionBandBackdrop As Integer
    SectionBandBackdropColor As ColorType
    SectionBandBackdropTransparency As Double
    SectionBandBackdropScale As Double
    FlatBar As Integer
    FlatBarHeight As Double
    FlatBarGap As Double
    FlatBarRoundness As Double
    FlatBarGlow As Integer
    FlatBarGlowRadius As Double
    FlatBarGlowTransparency As Double
    LayeredNavigation As Integer
    NavigationWidth As Double
    SectionBlockHeight As Double
    RailWidth As Double
    RailThickness As Double
    LayerShadowBlur As Double
    LayerShadowTransparency As Double
    LayerDepthStep As Double
    LayerShadowDistance As Double
    AccentMode As Integer
    AccentWidth As Double
    AccentColor As ColorType
    SectionNumbering As Integer
    CurrentTitleBold As Integer
    FutureBarOutline As Integer
    SectionForegroundBars As Integer
    DockEdge As Integer
    SidebarWidth As Double
    DisplayBindingMode As Integer
    DisplayBindingId As String
    UseThemeColors As Integer
    SectionRoundness As Double
    ReuseNavigation As Integer
    HardShadowOffset As Double
    CounterMode As Integer
    DensityMode As Integer
    MinMarkerLength As Double
    TitleOverflow As Integer
    TitleAlignment As Integer
    TitleCase As Integer
    FlatRuleMode As Integer
    FlatRuleWidth As Double
    FlatRuleColor As ColorType
    MarkerKind As Integer
    PanelShadow As Integer
    PresetSize As Integer
    CrimsonType As Integer
    InactiveSectionBarTransparency As Integer
    BandHeightMode As Integer
End Type

Private Const PI_PROPORTIONAL_BAND As Integer = 1
Private Const PI_EXACT_BAND As Integer = 2
Public Const PI_SMALL_BODY_HEIGHT As Double = 13
Public Const PI_CRIMSON_SMALL_BODY_HEIGHT As Double = 16
Private Const PI_CRIMSON_MARKER_SPAN As Double = 0.8
Public Const PI_CRIMSON_NONCURRENT_PAGE_TRANSPARENCY As Integer = 85
Private Const PI_INACTIVE_SEGMENT_HEIGHT As Double = 0.48

' The extra array slot stores the end sentinel after 100 sections.
Type ProgressDotsSectionInfo
    lngSecStartSlide As Long                          'old: where indicators start ... now always on slide 1 (by definition, can be changed by user by skipping/hiding page)
    lngSecSections As Long                              'count of all sections if all slides wouldget an indicator. see: strArySecTotalInvisible
    lngSecSectionsRecalculated As Long                      'recalculated total number of sections (total-sections which are invisible because of hidden/skipped slides). see: lngSecSections
    lngSecIndicatorsRecalculated As Long                    'recalculated total number of visible slide-indicators (how many are on a page (seen+unseen)
    strSecTotalVisibleComplete As String            'list of ALL slides VISIBLE (in ALL sections) e.g. ";1;2;3;4;6;...;12;14;...;26;28;...;80". Needed for quick hyperlink creation on the different indicators
    strSecTotalInvisibleComplete As String          'list of ALL slides invisible (in ALL sections) e.g. ";5;13;27"
    
    strArySecTitles(101) As String                   'title of section X (or empty)
    lngArySecStart(101) As Long                       'start slide of section X
    
    intArySecIgnoreHidden(101) As Integer            'number of hidden pages in section X. see: strArySecTotalInvisible
    intArySecIgnoreSkipped(101) As Integer           'number of skipped pages in section X. see: strArySecTotalInvisible
    
    intArySecTotal(101) As Integer                   'total number of pages in section X
    strArySecTotalInvisible(101) As String           'total number of pages in section X where no indicator - for whatever reason - will be seen. see: intArySecIgnoreSkipped, intArySecIgnoreHidden
        
    intArySecVisible(101) As Integer                 'section X is visible (1) of invisible (0, because all slides in section are hidden/skipped).

    intAryFirstVisiblePageInSection(101) As Integer   '<>0=first visible page in section X;
    
    '2018-04
    intArySecNumberSection(101) As Integer           'which section belongs slide Y to? [1..lngSecSections]
End Type

Private mPISilent As Boolean
Private mPICancelAfter As Long
Private mPICancelFile As String
Private mPIScratchRoot As String
Private mPITraceFile As String
Private mPITraceSection As Long
Private mPISurfaceDirectory As String
Private mPIRunSerial As Long
Private mPIDrawn As Long
Private mPIResultStatus As String
Private mPIFirstNumber As Long
Private mPIFirstSource As String
Private mPIFirstDescription As String
Private mPIFirstStage As String
Private mPISlide As Long
Private mPISection As Long
Private mPIShapes As Long
Private mPIParameters As String
Private mPIStage As String
Private mPIRunActive As Boolean
Private mPILinksBySlide() As String
Private mPIHiddenBySlide() As Boolean
Private mPISkippedBySlide() As Boolean
Private mPIRepeatBySlide() As Boolean
Private mPISlideCount As Long
Private mPIMetadataReads As Long
Private Type PISectionSlot
    axisStart As Double
    axisLength As Double
    pageCount As Long
    surfacePath As String
End Type

Private Type PIMarkerSlot
    axisStart As Double
    crossStart As Double
    axisLength As Double
    crossLength As Double
    slotLength As Double
End Type

Private mPISectionSlots(101) As PISectionSlot
Private mPIMarkerSlots() As PIMarkerSlot
Private mPILayoutKey As String
Private mPIStoreLayout As Boolean
Private mPIReadLayout As Boolean
Private mPILastDrawResult As Shape
Private mPIRenderCanvasWidth As Double
Private mPIRenderCanvasHeight As Double
Private mPIComputedPages As Long
Private mPILayoutHits As Long
Private mPISurfaceHits As Long
Private mPICaptures As Long
Private mPILastSummary As String
Private mPIRunStarted As Double
Private mPIRunPresentation As Presentation
Private mPINativePresentation As Presentation
Private mPIRenderWindow As DocumentWindow
Private mPIOriginalView As PpViewType
Private mPIOriginalSlideId As Long
Private mPIOriginalSelection() As Long
Private mPIOriginalSelectionCount As Long
Private mPIWindowPrepared As Boolean


' ----------------------------------------------------------------
' Procedure Name: RefreshRibbon
' Purpose: to refresh the ribbon with the values from the presentation: invalidate ... simple as can be
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2020-02-18
' ----------------------------------------------------------------
Sub RefreshRibbon()
10        On Error Resume Next    'in case ribbon does not exist (yet/anymore/some other reason this fails)
20        Auto_Open
          'gobjRibbon.Invalidate    ' v192 why was this here before 1.92? I can not remember - and it works (even better) without this line

End Sub

' ----------------------------------------------------------------
' Procedure Name: AddIgnoreSlide
' Purpose: adds the ignore (skipped) comment to a slide
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub AddIgnoreSlide()
    If gblbIndicatorUpdateRunning Then Exit Sub
          'idea and (some) code from https://github.com/utopiaplanetia/ppt-progress-bar
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Sub AddIgnoreSlide"         'Debug: Write active procedure to log
          
          Dim curslide As slide
30        Set curslide = GetActiveSlide(ActiveWindow)
40        If curslide Is Nothing Then GoTo error_section  'if current slide can not be evaluated
          
          'Note to myself if the new method still comes up with error on MAC, maybe try
          ' http://stackoverflow.com/questions/30050130/ddg#32199430
50        With curslide
60            .Comments.Add 0, 0, "Progress Indicator", "pi", gblstrMessageIgnore
70        End With
          'Beep
          'MsgBox "Do not forget to recreate the progress bar to see the effect of ignored slides!", , "Slide ignored now! Don't forget to refresh!"
              
exit_proc:
80        Exit Sub
          
error_section:
90        MsgBox PI_Zh("65E06CD583B753D65F53524D9875780130028BF7900962E94E005F205E7B706F7247FF0C70B951FB7F168F91533A57DF540E91CD8BD53002"), , "AddIgnoreSlide: GetActiveSlide failed"
100       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
110       Resume exit_proc
120       Resume
          
ErrorHandler:
130       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
140       Select Case Err.Number
              Case -2147188160      'no active presentation
150               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "AddIgnoreSlide: Nothing to do!"
160               Resume exit_proc
170               Resume
              
180           Case Else
190               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "AddIgnoreSlide"
                  
200               Resume Next
210               Resume
              
220       End Select
End Sub
' ----------------------------------------------------------------
' Procedure Name: AddRepeatSlide
' Purpose: adds the repeat comment to a slide
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2021-03-05
' ----------------------------------------------------------------
Sub AddRepeatSlide()
    If gblbIndicatorUpdateRunning Then Exit Sub

10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Sub AddRepeatSlide"         'Debug: Write active procedure to log
          
          Dim curslide As slide
30        Set curslide = GetActiveSlide(ActiveWindow)
40        If curslide Is Nothing Then GoTo error_section  'if current slide can not be evaluated
          
          'Note to myself if the new method still comes up with error on MAC, maybe try
          ' http://stackoverflow.com/questions/30050130/ddg#32199430
50        With curslide
60            .Comments.Add 0, 0, "Progress Indicator", "pi", gblstrMessageRepeat
70        End With
          'Beep
          'MsgBox "Do not forget to recreate the progress bar to see the effect of ignored slides!", , "Slide ignored now! Don't forget to refresh!"
              
exit_proc:
80        Exit Sub
          
error_section:
90        MsgBox PI_Zh("65E06CD583B753D65F53524D9875780130028BF7900962E94E005F205E7B706F7247FF0C70B951FB7F168F91533A57DF540E91CD8BD53002"), , "AddRepeatSlide: GetActiveSlide failed"
100       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
110       Resume exit_proc
120       Resume
          
ErrorHandler:
130       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
140       Select Case Err.Number
              Case -2147188160      'no active presentation
150               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "AddRepeatSlide: Nothing to do!"
160               Resume exit_proc
170               Resume
              
180           Case Else
190               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "AddRepeatSlide"
                  
200               Resume Next
210               Resume
              
220       End Select
End Sub


' ----------------------------------------------------------------
' Procedure Name: Auto_Open
' Purpose: run automatically when the macro is run. variables are initialized here
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub Auto_Open()
          If gblbIndicatorUpdateRunning Or mPISilent Then Exit Sub
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Sub Auto_Open"         'Debug: Write active procedure to log
          
          Dim s As String
          Dim theOptions As ProgressDotsOptions
          Dim thePresentation As Presentation
          Dim savedDrawStart As Integer, savedDrawEnd As Integer
          Dim savedCreatePressed As Boolean, savedPreview As Integer

          ''debug for breaking
          'MsgBox "now auto_open from onload_ribbon 1"
          'MsgBox "now auto_open from onload_ribbon 2"
30        If gblbolDebug = True Then LogInfo "Auto_Open running at " & Now()
40        If gblIndicatorSettingsVersion = "" Then  ' do only once, so that values from the ribbon that are NOT saved in the PI-settings are not reset when the user changes them
50            If gblbolDebug = True Then LogInfo "Initializing the first time"
60            Initialize    'write default values to ribbon (fill ribbon with content)
70            getDefaultOptions
80        Else
90            If gblbolDebug = True Then LogInfo "Initializing again (not only the first time)"
100       End If
          'set theOptions to defaults that have been initialized
          
          ' Load the options - if needed get defaults
110       On Error Resume Next  'needed if there is no active presentation to read set active ...
120       Set thePresentation = ActivePresentation
130       s = thePresentation.Tags.Item("__progress-indicator-config__")    'read theOptions from existing presentation
140       On Error GoTo ErrorHandler        'back to usual error handling
          'MsgBox "s_before(" & Len(s) & "): " & s
150       If s = "" Then
              If Not thePresentation Is Nothing Then
                  savedDrawStart = gbliDrawStart
                  savedDrawEnd = gbliDrawEnd
                  savedCreatePressed = gblCreatePressed
                  savedPreview = gblShowPreview
                  Initialize
                  gbliDrawStart = savedDrawStart
                  gbliDrawEnd = savedDrawEnd
                  gblCreatePressed = savedCreatePressed
                  gblShowPreview = savedPreview
              End If
              theOptions = getDefaultOptions()
              If Not thePresentation Is Nothing Then PI_SetFullWidth PI_RenderAxisLength(theOptions, thePresentation)
              s = SerializeProgressOptions(getDefaultOptions())
          End If
          'MsgBox "s_after (" & Len(s) & "): " & s
160       theOptions = DeserializeProgressOptions(s)    ' s is either from current presentation or default values

170       optionsToRibbon theOptions    'set theoptions to global variables (from ribbon)

          'trap ppt-events so we can react on presentation open, change etc.
180       TrapEvents

exit_proc:
190       Exit Sub
          
ErrorHandler:
200       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
          'Just in case there is an error
210       MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "Auto_Open"
220       Resume Next
230       Resume
End Sub

' ----------------------------------------------------------------
' Procedure Name: CharCount
' Purpose: Counts how many characters occur in the given string
' Procedure Kind: Function
' Procedure Access: Public
' Parameter strText (String): string to search
' Parameter strSearch (String): character to count
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function CharCount(strText As String, strSearch As String)
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct CharCount: " & strText & " / " & strSearch          'Debug: Write active procedure to log
          
30        CharCount = Len(strText) - Len(Replace(strText, strSearch, ""))

exit_proc:
40        Exit Function
          
ErrorHandler:
          If mPISilent Then PI_RaiseNativeFailure "CharCount", Erl, Err.Number, Err.Source, Err.Description
50        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
60        Select Case Err.Number
              Case -2147188160      'no active presentation
70                MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "CharCount: Nothing to do!"
80                Resume exit_proc
90                Resume
100           Case Else
110               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "CharCount"
120               Resume Next
130               Resume
              
140       End Select
End Function

' ----------------------------------------------------------------
' Procedure Name: CheckActiveSlideVisible
' Purpose: returns 1 if the slide in question should get an indicator (depends on if it's hidden/skipped and matching settings)
' Procedure Kind: Function
' Procedure Access: Public
' Parameter strHidden (String): string of slides hidden, slidennumbers seperated by ;
' Parameter intSearch (Integer): which slide is in question
' Return Type: Integer
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function CheckActiveSlideVisible(strHidden As String, intSearch As Integer) As Integer
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct CheckActiveSlideVisible: " & strHidden & " / " & intSearch        'Debug: Write active procedure to log
30        CheckActiveSlideVisible = InStr(1, strHidden, ";" & intSearch & ";")
          
exit_proc:
40        Exit Function
          
ErrorHandler:
          If mPISilent Then PI_RaiseNativeFailure "CheckActiveSlideVisible", Erl, Err.Number, Err.Source, Err.Description
50        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
60        Select Case Err.Number
              Case -2147188160      'no active presentation
70                MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "MakeProgressBar: Nothing to do!"
80                Resume exit_proc
90                Resume
100           Case Else
110               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "MakeProgressBar"
120               Resume Next
130               Resume
              
140       End Select
End Function

' ----------------------------------------------------------------
' Procedure Name: CheckSlideHidden
' Purpose: Test if a slide is hidden (returns 1 if hidden, 0 if not)
' Procedure Kind: Function
' Procedure Access: Public
' Parameter intSlideNumber (Integer): slidenumber
' Return Type: Boolean
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function CheckSlideHidden(intSlideNumber As Integer) As Boolean
10        On Error GoTo ErrorHandler
          If mPIRunActive Then
              If intSlideNumber >= 1 And intSlideNumber <= mPISlideCount Then
                  CheckSlideHidden = mPIHiddenBySlide(intSlideNumber)
                  Exit Function
              End If
          End If
20        If gblbolDebug = True Then LogInfo "Fct CheckSlideHidden: " & intSlideNumber         'Debug: Write active procedure to log
              
30        With PI_RenderPresentation()
40            CheckSlideHidden = .Slides(intSlideNumber).SlideShowTransition.Hidden
50        End With
          
exit_proc:
60        Exit Function
          
ErrorHandler:
          If mPISilent Then PI_RaiseNativeFailure "CheckSlideHidden", Erl, Err.Number, Err.Source, Err.Description
70        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
80        Select Case Err.Number
              Case -2147188160      'no active presentation
90                MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "CheckSlideHidden: Nothing to do!"
100               Resume exit_proc
110               Resume
120           Case Else
130               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "CheckSlideHidden"
140               Resume Next
150               Resume
              
160       End Select
End Function

' ----------------------------------------------------------------
' Procedure Name: CheckSlideSkipped
' Purpose: Test if a slide is skipped (returns 1 if skipped, 0 if not)
' Procedure Kind: Function
' Procedure Access: Public
' Parameter intSlideNumber (Integer): slidenumber
' Return Type: Boolean
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function CheckSlideSkipped(intSlideNumber As Integer) As Boolean
10        On Error GoTo ErrorHandler
          If mPIRunActive Then
              If intSlideNumber >= 1 And intSlideNumber <= mPISlideCount Then
                  CheckSlideSkipped = mPISkippedBySlide(intSlideNumber)
                  Exit Function
              End If
          End If
20        If gblbolDebug = True Then LogInfo "Fct CheckSlideSkipped: " & intSlideNumber         'Debug: Write active procedure to log
          
          Dim i As Integer
30        CheckSlideSkipped = False   'default
40        With PI_RenderPresentation().Slides(intSlideNumber)
50            If .Comments.Count > 0 Then
60                For i = 1 To .Comments.Count
70                    If .Comments.Item(i).Text = gblstrMessageIgnore Then
80                        CheckSlideSkipped = True
                          'skipp rest
90                        i = .Comments.Count
100                   End If
110               Next i
120           End If
130       End With
         
exit_proc:
140       Exit Function
          
ErrorHandler:
          If mPISilent Then PI_RaiseNativeFailure "CheckSlideSkipped", Erl, Err.Number, Err.Source, Err.Description
150       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
160       Select Case Err.Number
              Case -2147188160      'no active presentation
170               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "CheckSlideSkipped: Nothing to do!"
180               Resume exit_proc
190               Resume
200           Case Else
210               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "CheckSlideSkipped"
220               Resume Next
230               Resume
              
240       End Select
End Function


' ----------------------------------------------------------------
' Procedure Name: CheckSlideRepeat
' Purpose: Test if PI should be repeated on slide (returns 1 if skipped, 0 if not)
' Procedure Kind: Function
' Procedure Access: Public
' Parameter intSlideNumber (Integer): slidenumber
' Return Type: Boolean
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2021-03-08
' ----------------------------------------------------------------
Function CheckSlideRepeat(intSlideNumber As Integer) As Boolean
10        On Error GoTo ErrorHandler
          If mPIRunActive Then
              If intSlideNumber >= 1 And intSlideNumber <= mPISlideCount Then
                  CheckSlideRepeat = mPIRepeatBySlide(intSlideNumber)
                  Exit Function
              End If
          End If
20        If gblbolDebug = True Then LogInfo "Fct CheckSlideRepeat: " & intSlideNumber         'Debug: Write active procedure to log
          
          Dim i As Integer
          
          'If ActivePresentation.Slides.Count >= intSlideNumber Then

30        CheckSlideRepeat = False   'default
40        With PI_RenderPresentation().Slides(intSlideNumber)
50            If .Comments.Count > 0 Then
60                For i = 1 To .Comments.Count
70                    If .Comments.Item(i).Text = gblstrMessageRepeat Then
80                        CheckSlideRepeat = True
                          'skipp rest
90                        i = .Comments.Count
100                   End If
110               Next i
120           End If
130       End With
         
          'End If
         
exit_proc:
140       Exit Function
          
ErrorHandler:
          If mPISilent Then PI_RaiseNativeFailure "CheckSlideRepeat", Erl, Err.Number, Err.Source, Err.Description
150       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
160       Select Case Err.Number
              Case -2147188160      'no active presentation
170               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "CheckSlideRepeat: Nothing to do!"
180               Resume exit_proc
190               Resume
200           Case Else
210               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "CheckSlideRepeat"
220               Resume Next
230               Resume
              
240       End Select
End Function

' ----------------------------------------------------------------
' Procedure Name: ChooseFromAllFonts
' Purpose: opens the font-chooser-form where all installed fonts are read from word and displayed in a dropdown list
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter gblstrFontname (String): which font to show as selected in the dropdown list. also: which font to return later to the ribbon
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub ChooseFromAllFonts(gblstrFontname As String)
10        On Error GoTo ErrorHandler
          
20        If gblbolDebug = True Then LogInfo "Sub ChooseFromAllFonts:" & gblstrFontname          'Debug: Write active procedure to log
          
30        gblFontnameRibbonActivated = gblstrMessageMore
40        frmFontChooser.ListFontsChooser
50        frmFontChooser.fontName = gblstrFontname
60        frmFontChooser.Show
          
exit_proc:
70        Exit Sub
          
ErrorHandler:
80        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
90        Select Case Err.Number
              Case -2147188160      'no active presentation
100               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "ChooseFromAllFonts: Nothing to do!"
110               Resume exit_proc
120               Resume
130           Case Else
140               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "ChooseFromAllFonts"
150               Resume Next
160               Resume
              
170       End Select
End Sub

' ----------------------------------------------------------------
' Procedure Name: DelProgressBar
' Purpose: completely remove indicatorbar
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub DelProgressBar()
    If gblbIndicatorUpdateRunning Then Exit Sub
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Sub DelProgressBar"         'Debug: Write active procedure to log
          
          Dim slideNum As Long
30        If gblbolDebug = True Then LogInfo "Deleting old progress bar..."
          '35        On Error GoTo error_NoSlides      'to prevent error when there are no slides at all / no presentation
40        For slideNum = 1 To ActivePresentation.Slides.Count
50            If slideNum >= gbliDrawStart And slideNum <= gbliDrawEnd Then
60                If PI_IsGenerationTarget(ActivePresentation.Slides(slideNum)) Then DelProgressBarOnSlide slideNum, "__progress-indicator__"
70            End If
80        Next slideNum
          '65        On Error GoTo errorHandler    'back to regular error handler
          
exit_proc:

90        Exit Sub
          
          'error_NoSlides:
          '73        Beep
          '75        Resume exit_proc:
          '78        Resume
          
ErrorHandler:
100       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
110       Select Case Err.Number
          
              Case -2147188160      'no active presentation
120               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "DelProgressBar: Nothing to do!"
130               Resume exit_proc
140               Resume

150           Case Else
160               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "DelProgressBar"
170               Resume Next
180               Resume
              
190       End Select

End Sub

' ----------------------------------------------------------------
' Procedure Name: DelProgressBarOnSlide
' Purpose: remove indicatorbar on slide X
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter slideNum (Long): slidenumber
' Parameter shapeTag (String): to remove only indicator shapes
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub DelProgressBarOnSlide(slideNum As Long, shapeTag As String)
    Dim presentation As Presentation, targetSlide As Slide, index As Long
    PI_RequireNativeTarget
    Set presentation = PI_RenderPresentation()
    Set targetSlide = presentation.Slides(slideNum)
    For index = targetSlide.Shapes.Count To 1 Step -1
        If UCase$(targetSlide.Shapes(index).Tags(shapeTag)) = "YES" Then targetSlide.Shapes(index).Delete
    Next index
End Sub

' ----------------------------------------------------------------
' Procedure Name: GetSlideIndex
' Purpose: Return a nice link title for each of the indicators
' Procedure Kind: Function
' Procedure Access: Public
' Parameter slideNum (Long): slidenumber we want the index for
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-24
' ----------------------------------------------------------------
Function GetSlideIndex(slideNum As Long)
10        On Error GoTo ErrorHandler
          If mPIRunActive Then
              If slideNum >= 1 And slideNum <= mPISlideCount Then
                  GetSlideIndex = mPILinksBySlide(slideNum)
                  Exit Function
              End If
          End If
20        If gblbolDebug = True Then LogInfo "Fct GetSlideIndex: " & slideNum
          
          Dim strTitle As String
          Dim strID As String
          Dim strIndx As String

          'get ID,Index and Title from slidenumber
30        With PI_RenderPresentation().Slides(slideNum)
40            strID = CStr(.SlideID)
50            strIndx = CStr(.SlideIndex)
60            If .Shapes.HasTitle Then strTitle = .Shapes.title.TextFrame.TextRange _
                  Else: strTitle = "Slide " & strIndx
70        End With
80        If strTitle = "" Then strTitle = "-"    'may be removed, see http://www.pptfaq.com/FAQ00162_Hyperlink_-SubAddress_-_How_to_interpret_it.htm
90        GetSlideIndex = strID & "," & strIndx & "," & strTitle          'strID (index) needed for hyperlink to work!
          'GetSlideIndex = strIndx & "," & strTitle
          
exit_proc:
100       Exit Function
          
ErrorHandler:
          If mPISilent Then PI_RaiseNativeFailure "GetSlideIndex", Erl, Err.Number, Err.Source, Err.Description
110       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
120       Select Case Err.Number
              Case -2147188160      'no active presentation
130               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "GetSlideIndex: Nothing to do!"
140               Resume exit_proc
150               Resume
160           Case Else
170               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "GetSlideIndex"
180               Resume Next
190               Resume
              
200       End Select
End Function


' ----------------------------------------------------------------
' Procedure Name: DrawProgressBarOnSlide
' Purpose: create indicatorbar on a certain slide
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter slideNum (Long): slidenumber
' Parameter sectinfo (ProgressDotsSectionInfo): information about (existing) sections
' Parameter theOptions (ProgressDotsOptions): options to use (font, color etc.)
' Parameter shapeTag (String): which shape to use
' Parameter slideNumVisible (Integer): number of slides visible
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2022-02-04
' ----------------------------------------------------------------
Sub DrawProgressBarOnSlide(slideNum As Long, _
    sectinfo As ProgressDotsSectionInfo, _
    theOptions As ProgressDotsOptions, _
    shapeTag As String, _
    slideNumVisible As Integer)
10        On Error GoTo ErrorHandler

20        If gblbolDebug = True Then LogInfo "Sub DrawProgressBarOnSlide(slideNum As Long, sectinfo As ProgressDotsSectionInfo, theOptions As ProgressDotsOptions, _" & shapeTag & " / " & slideNumVisible       'Debug: Write active procedure to log
          
          Dim renderPresentation As Presentation
          Dim slideNumLast As Long
          
          Dim varArray() As String
          Dim strLinkTitle  As String
          
          Dim strDebugtext  As String
          
          Dim intShapetypeNr As Integer
          Dim labelWidth  As Double     '    x, sectionMargin    , dotMargin
          Dim labelLeft As Double     '    x, sectionMargin    , dotMargin

          Dim numDotsInSection As Long
          Dim dotNum As Long
          Dim sectNum As Long
          Dim sectNum_old As Long   'Introduced in 1.923 to make sure section width is calculated only once
          Dim numDots As Long
          Dim numSections As Long

          Dim dot As Shape        ', sectTitle
                    
          Dim indicator2 As Shape
          '    Dim slideWidth, gapSize As Double
          Dim firstShape As Long
          
          Dim lngSeenDotColor As Long
          Dim lngUnseenDotColor As Long
          Dim lngSeenTextColor As Long
          Dim lngUnseenTextColor As Long
          Dim lngTextBGColorSeen As Long
          Dim lngTextBGColorUnseen As Long
          
          Dim numSectionsTotal As Integer
          Dim numDotsCount As Integer           'current indicator of all indicators to paint
          Dim numDotsCountInSection As Integer  'Current indicator of all indicators in section to paint (resets with every new section)
          Dim strSectionNameToShow As String
          Dim fullSectionTitle As String
          
          Dim intIndicatorCntTotal_new As Integer            'DT (dot count total) how many indicators to draw on one slide
          Dim intIndicatorCntTotalInSect_new As Integer      'DS (dots in section) how many indicators in active calculated section
          Dim intSectCntTotal_new As Integer                 'ST (section total) how many sections to show
          Dim dblSectLeft_new As Double                      'SL (section left) left position of the current section
          Dim dblSectWidth_new As Double                     'SW (section width) how wide is the currently calculated / drawn section
          Dim dblPageWidth_new As Double                     'PW (page with), presentation total width
          Dim dblUnitWidth_new As Double                     'UW (unit width) width of each indicator (evenly spread on page)
          
          '    Dim dblSpaceNeeded_new As Double                 'SN (space needed) space between two sections
          '    Dim intUnitCntTotal_new As Integer               'UT (unit count total) how many sections to show
          '    Dim dblIndicatorWidth_new As Double              'UW (unit width) how wide is the indicator
          
          Dim dblSpaceIndicatorAdjust As Double
          Dim dblIndicatorSqueeze  As Double

          Dim currentSection As Long
          Dim sectionIs As Integer  '0= Seen, 1= Active, 2= UNseen
          Dim visibleSectionOrdinal As Long, sectionBarColor As Long
          
          Dim intSectionIndicatorGap As Integer
             
          Dim dbllabelWidth_test, dbllabelLeft_test As Double
          Dim dblSpace_test As Double
          Dim newShape2 As Shape
          Dim bandShape As Shape
          Dim bandMode As Boolean
          Dim titleTop As Double, dotsTop As Double, bandHeight As Double
          Dim splitBand As Boolean, bodyHeight As Double, bodyTop As Double, rowTop As Double
          Dim rowHeight As Double, counterTop As Double
          Dim titleHeight As Double, bandPadding As Double, titlePadding As Double
          Dim addedHeight As Double, markerClearance As Double, presetGrowth As Double, titleBodyHeight As Double
          Dim renderDotSize As Double, renderDotTop As Double, dotPadding As Double
          Dim dotAlong As Double, dotAcross As Double, ovalSlot As Double
          Dim renderBarWidth As Double, renderBarHeight As Double
          Dim sizedMarkers As Boolean, crimsonMarkers As Boolean, slotCount As Long, slotIndex As Long
          Dim barSlotWidth As Double, barGap As Double, barAreaLeft As Double, barAreaWidth As Double
          Dim markerRowHeight As Double
          Dim flatBarMode As Boolean
          Dim collapsedSection As Boolean
          Dim counter As Shape
          Dim layeredMode As Boolean
          Dim navigationLeft As Double, navigationWidth As Double
          Dim railLeft As Double, railWidth As Double
          Dim bandPosition As Double
          Dim backdropErrorNumber As Long, backdropErrorSource As String
          Dim backdropErrorDescription As String
          
          Dim dblLeft_new  As Double
          Dim intDebugValue As Integer
          
          Dim dblLeftCalculatedNew  As Double            'This is the base left for the new dot
          Dim dbllabelWidthCalculated As Double          'label width corrected by padding
          Dim dbllabelWidthCalculatedBy2 As Double       'label width corrected by padding
          Dim dblLeft_newC2 As Double
          Dim dblnumDotsCountCorrector As Double         'to make -n to n calculation possible
          
          Dim lngCurDotColF As Long
          Dim lngCurDotColB As Long
          Dim lngCurTxtColF As Long
          Dim lngCurTxtColB As Long
          
          Dim bFillShape As Boolean
          
          Dim strCustomImage_Current As String
          Dim strCustomImage_Seen As String
          Dim strCustomImage_Active As String
          Dim strCustomImage_UnSeen As String
          
30        strCustomImage_Current = ""
          Set mPILastDrawResult = Nothing
          Set renderPresentation = PI_RenderPresentation()
          PI_RequireNativeTarget
40        strCustomImage_Seen = gblCustomImage_Seen
50        strCustomImage_Active = gblCustomImage_Active
60        strCustomImage_UnSeen = gblCustomImage_UnSeen
          
70        intDebugValue = 0               'debug only if old and new indicators are used, this add to the y value of the new (place it lower than the old)
          
80        intSectionIndicatorGap = theOptions.SectionIndicatorGap
          bandMode = PI_UsesSectionBand(theOptions)
          layeredMode = (bandMode And theOptions.LayeredNavigation = 1)
          flatBarMode = (theOptions.FlatBar = 1)
          crimsonMarkers = flatBarMode And PI_UsesCrimsonMarkers(theOptions)
          sizedMarkers = flatBarMode And (crimsonMarkers Or theOptions.PresetSize > 0 Or theOptions.MarkerKind = 3)
          markerRowHeight = PI_MarkerHeight(theOptions)
          navigationLeft = theOptions.LeftMargin
          mPIRenderCanvasWidth = PI_RenderAxisLength(theOptions, renderPresentation)
          mPIRenderCanvasHeight = renderPresentation.PageSetup.SlideHeight
          If theOptions.DockEdge >= 2 Then mPIRenderCanvasHeight = renderPresentation.PageSetup.SlideWidth
          navigationWidth = mPIRenderCanvasWidth - theOptions.LeftMargin - theOptions.RightMargin
          firstShape = renderPresentation.Slides(slideNum).Shapes.Count + 1
          If bandMode Or flatBarMode Or layeredMode Then
              PI_NavigationGeometry theOptions, mPIRenderCanvasWidth, _
                  navigationLeft, navigationWidth, railLeft, railWidth
          End If
          If bandMode Then
              PI_BandGeometry theOptions, titleTop, dotsTop, bandHeight, presetGrowth
              splitBand = (theOptions.PresetSize > 0 And theOptions.DockEdge < 2)
              bodyTop = theOptions.yMargin
              bodyHeight = bandHeight
              If splitBand Then
                  bodyHeight = PI_CompactBodyHeight(theOptions) + presetGrowth
                  markerClearance = presetGrowth / 3
                  rowHeight = markerRowHeight + PI_CounterRowHeight(theOptions)
                  rowTop = theOptions.yMargin
                  If theOptions.SectionPlacement = 0 Then
                      rowTop = theOptions.yMargin + bandHeight - rowHeight
                  Else
                      bodyTop = rowTop + rowHeight
                  End If
              End If
              titleHeight = PI_BandTitleHeight(theOptions.textHeight)
              If theOptions.PresetSize = 2 And theOptions.DockEdge < 2 Then
                  titleBodyHeight = bodyHeight
                  If theOptions.CrimsonType = 1 And theOptions.BandHeightMode = 0 And _
                          bodyHeight = PI_CRIMSON_SMALL_BODY_HEIGHT Then
                      ' Keep the text anchored to its outer edge as the gap to the markers grows.
                      titleBodyHeight = PI_SMALL_BODY_HEIGHT
                      If theOptions.SectionPlacement <> 0 Then titleTop = titleTop + bodyHeight - titleBodyHeight
                  End If
                  titleHeight = PI_BarClamp(theOptions.textHeight, 0, titleBodyHeight)
                  addedHeight = PI_BarClamp(titleBodyHeight - 10, 0, 3)
                  markerClearance = addedHeight / 3
                  addedHeight = PI_BarClamp(addedHeight, 0, titleBodyHeight - titleHeight)
                  titleTop = titleTop + (titleBodyHeight - addedHeight - titleHeight) / 2
                  If theOptions.SectionPlacement = 0 Then
                      titleTop = titleTop + 2 * addedHeight / 3
                      dotsTop = dotsTop - markerClearance
                  Else
                      titleTop = titleTop + addedHeight / 3
                      dotsTop = dotsTop + markerClearance
                  End If
              End If
              bandPadding = theOptions.SectionBandPadding
          Else
              titleTop = theOptions.yMargin + (theOptions.SectionPlacement * markerRowHeight + 1) + (theOptions.SectionPlacement * intSectionIndicatorGap)
              dotsTop = theOptions.yMargin + ((1 - theOptions.SectionPlacement) * (theOptions.textHeight + 3)) + ((1 - theOptions.SectionPlacement) * intSectionIndicatorGap)
              titleHeight = theOptions.textHeight + 10
              If theOptions.PIStyle = "Sections only" Then titleTop = theOptions.yMargin
              If theOptions.SectionPlacement = 0 And PI_CounterRowHeight(theOptions) > 0 Then
                  dotsTop = titleTop + titleHeight + PI_CounterRowHeight(theOptions) + intSectionIndicatorGap
              End If
          End If
          
          'get colors once, could be done ONCE before this sub!
90        lngSeenDotColor = ColorToRGB(theOptions.seenDotColor)
100       lngUnseenDotColor = ColorToRGB(theOptions.unseenDotColor)

110       lngSeenTextColor = ColorToRGB(theOptions.seenTextColor)
120       lngUnseenTextColor = ColorToRGB(theOptions.unseenTextColor)
          
130       lngTextBGColorSeen = ColorToRGB(theOptions.seenTextColorBG)
140       lngTextBGColorUnseen = ColorToRGB(theOptions.unseenTextColorBG)

150       lngCurDotColF = ColorToRGB(theOptions.CurrentDotColor)
160       lngCurDotColB = lngCurDotColF
170       lngCurTxtColF = ColorToRGB(theOptions.CurrentTextColor)
180       lngCurTxtColB = ColorToRGB(theOptions.CurrentTextColorBG)


                  
          ' Record the index of the first shape we'll be creating.
          ' We'll use this later to select the whole range of shapes & group them.
190       firstShape = renderPresentation.Slides(slideNum).Shapes.Count + 1
          
          'recalculated number of indicator shapes that will be shown in total
200       numDots = sectinfo.lngSecIndicatorsRecalculated
210       intIndicatorCntTotal_new = sectinfo.lngSecIndicatorsRecalculated            'DT n2
220       numSections = sectinfo.lngSecSectionsRecalculated             ' = recalculated number of sections that will be shown in total
230       numSectionsTotal = sectinfo.lngSecSections                    ' = total number of sections that would get an indicator
240       intSectCntTotal_new = sectinfo.lngSecSectionsRecalculated                   ' ST  n2
          If bandMode Or flatBarMode Or layeredMode Then
              barAreaWidth = navigationWidth
              If barAreaWidth - (numSections - 1) * theOptions.SectionGap <= 0 Then
                  Err.Raise vbObjectError + 3021, "DrawProgressBarOnSlide", PI_Zh("6A2A67616CA167098DB3591F76846C345E737A7A95F4FF0C8BF751CF5C0F5DE653F38FB98DDD62167AE0828295F48DDD3002")
              End If
          End If
          
250       numDotsInSection = -99      'invalidate counter
          
260       sectNum_old = 0 'Default
          'set sectNum to FIRST visible (!) section
270       sectNum = 1

280       While sectinfo.intArySecVisible(sectNum) = 0 And sectNum < numSectionsTotal
290           If gblbolDebug = True Then LogInfo "no slide to show in section: " & sectNum & " " & sectinfo.strArySecTitles(sectNum)
300           sectNum = sectNum + 1
310       Wend
          
          'sectNum_old = sectNum 'Default, but after checking if there are any slides in this section
          
          'get shape id from the id in list, could be done ONCE before this sub!
320       intShapetypeNr = retShapeID(theOptions.dotShapeID)

          'get Array of visible pages for hyperlinks, could be done ONCE before this sub!
330       varArray = Split(sectinfo.strSecTotalVisibleComplete, ";")
          
340       numDotsCount = 0    'number of dots already drawn in current section
350       dotNum = 0          'total number of dots to draw
          'numDotsCountInSection = 0
          
          ' Loop through the slides, and draw an indicator (formerly dot) for each one.
360       For dotNum = 1 To numDots       'because we want to start on the first page where there is an indicator and place X (numDots) indicators

              'how many dots have we dont already (can not use FOR value as we need to keep track for other sections)
370           numDotsCount = numDotsCount + 1
              'numDotsCountInSection = numDotsCountInSection + 1
              
              'get title for hyperlink to slide
380           strLinkTitle = GetSlideIndex(CInt(varArray(dotNum))) & "," & dotNum         '& "," & dotNum         my be remove, see http://www.pptfaq.com/FAQ00162_Hyperlink_-SubAddress_-_How_to_interpret_it.htm and https://www.experts-exchange.com/questions/27793436/PowerPoint-2010-Change-Hyperlink-programmatically.html

              'as we have "always" sections we go ahead ... but skip the writing part IF cbxIgnoreSections is checked!
              
390           If numDotsCount = 1 And sectinfo.intArySecVisible(sectNum) = 1 Then 'only draw section title IF this is the first dot of the section AND if section has visible indicators
                  ' sectinfo.intArySecTotal : total number of slides in section sectNum
                  ' CharCount:  skipped/repeated slides are listed in the array strArySecTotalInvisible seperated by ;
                  If mPIReadLayout Then
                      numDotsInSection = mPISectionSlots(sectNum).pageCount
                  Else
400                   numDotsInSection = sectinfo.intArySecTotal(sectNum) - CharCount(sectinfo.strArySecTotalInvisible(sectNum), ";")
                  End If
410               intIndicatorCntTotalInSect_new = numDotsInSection
                  'make it possible to show an empty sectionname
420               If sectinfo.strArySecTitles(sectNum) = gblstrSectionNameHidden Then
430                   strSectionNameToShow = ""
440               Else
450                   strSectionNameToShow = sectinfo.strArySecTitles(sectNum)
460               End If
                  visibleSectionOrdinal = visibleSectionOrdinal + 1
                  If theOptions.SectionNumbering = 1 And Len(strSectionNameToShow) > 0 Then
                      strSectionNameToShow = Format$(visibleSectionOrdinal, "00") & " " & strSectionNameToShow
                  End If
                  fullSectionTitle = strSectionNameToShow
                  If theOptions.TitleCase = 1 Then strSectionNameToShow = UCase$(strSectionNameToShow)
                  
                  
                  '--START NEW
                  '-- START NEW CALCULATION -----------------------------------------------------
                  If mPIReadLayout Then
                      dblSectLeft_new = mPISectionSlots(sectNum).axisStart
                      dblSectWidth_new = mPISectionSlots(sectNum).axisLength
                  Else
470               If gblbolDebug = True Then LogInfo "old dblSectWidth_new: " & dblSectWidth_new & " / old dblSectLeft_new: " & dblSectLeft_new
480               dblSpace_test = theOptions.SectionGap * 2
                  '                   of course: width of the slide          - left margin           - right margin           + correct for last section width
490               dblSpaceIndicatorAdjust = theOptions.IndicatorSpaceAdjustment
500               dblIndicatorSqueeze = theOptions.IndicatorSqueeze
510               dblPageWidth_new = navigationWidth + (dblSpace_test / 2)
520               intIndicatorCntTotal_new = sectinfo.lngSecIndicatorsRecalculated    '   dt = sectinfo.lngSecIndicatorsRecalculated
530               intSectCntTotal_new = sectinfo.lngSecSectionsRecalculated           '   st = sectinfo.lngSecSectionsRecalculated

540               Select Case theOptions.EvenSpaceSecOrInd
                      Case "Shapes"
550                       dblUnitWidth_new = dblPageWidth_new / intIndicatorCntTotal_new      'uw = pw / ut
                          If bandMode Or flatBarMode Or layeredMode Then dblUnitWidth_new = (dblPageWidth_new - numSections * theOptions.SectionGap) / intIndicatorCntTotal_new
                          '   next variables must always be recalculated in loop
560                       intIndicatorCntTotalInSect_new = sectinfo.intArySecTotal((sectNum)) - CharCount(sectinfo.strArySecTotalInvisible((sectNum)), ";")       '   ds  must be calculated in FOR loop
                          'The next IF is exactly the dame as in the CASE "SECTIONS" area ... but if it if places after the SELECT CASE the result is different than if it doubles inside.. VERY STRANGE!
                          'Difference: here, dblSectWidth_new is calculated after this IF again

570                       If sectNum_old = 0 Then
580                           dblSectLeft_new = navigationLeft
                              If Not bandMode And Not flatBarMode And Not layeredMode Then dblSectLeft_new = dblSectLeft_new + (1 / 100000)
590                       Else
600                           dblSectLeft_new = dblSectLeft_new + dblSectWidth_new + (dblSpace_test / 2)  '   sl = sl + sw
610                       End If
620                       If sectNum_old <> sectNum Then
630                           dblSectWidth_new = dblUnitWidth_new * intIndicatorCntTotalInSect_new - (dblSpace_test / 2)         '   sw = uw * ds
                              If bandMode Or flatBarMode Or layeredMode Then dblSectWidth_new = dblUnitWidth_new * intIndicatorCntTotalInSect_new
640                       End If

650                   Case "Sections"
660                       dblSectWidth_new = (dblPageWidth_new - (numSections * dblSpace_test / 2)) / numSections
                          '   next variables must always be recalculated in loop
                          'The next IF is exactly the dame as in the CASE "SHAPES" area ... but if it if places after the SELECT CASE the result is different than if it doubles inside.. VERY STRANGE!
670                       If sectNum_old = 0 Then
680                           dblSectLeft_new = navigationLeft
                              If Not bandMode And Not flatBarMode And Not layeredMode Then dblSectLeft_new = dblSectLeft_new + (1 / 100000)
690                       Else
700                           dblSectLeft_new = dblSectLeft_new + dblSectWidth_new + (dblSpace_test / 2)  '   sl = sl + sw
710                       End If

720               End Select

                      If mPIStoreLayout Then
                          With mPISectionSlots(sectNum)
                              .axisStart = dblSectLeft_new
                              .axisLength = dblSectWidth_new
                              .pageCount = numDotsInSection
                          End With
                      End If
                  End If
730               dbllabelLeft_test = dblSectLeft_new
740               dbllabelWidth_test = dblSectWidth_new
750               If gblbolDebug = True Then LogInfo "NEW dblSectWidth_new: " & dblSectWidth_new & " / NEW dblSectLeft_new: " & dblSectLeft_new

                  '-- END NEW CALCULATION -----------------------------------------------------
                  
                  
                  
                  
                  '####### START SECTIONS ########
    If Len(mPITraceFile) > 0 Then PI_Trace "chapter.draw", sectNum
                  Set bandShape = Nothing
                  If bandMode Then
                      If Not layeredMode And theOptions.SectionRoundness > 0 Then
                          Set bandShape = renderPresentation.Slides(slideNum).Shapes.AddShape( _
                              msoShapeRoundedRectangle, dbllabelLeft_test, theOptions.yMargin, dbllabelWidth_test, bandHeight)
                          bandShape.Adjustments.Item(1) = PI_BarClamp(theOptions.SectionRoundness, 0, 50) / 100
                      Else
                          Set bandShape = renderPresentation.Slides(slideNum).Shapes.AddShape( _
                              msoShapeRectangle, dbllabelLeft_test, theOptions.yMargin, dbllabelWidth_test, bandHeight)
                      End If
                      If splitBand And Not layeredMode Then
                          bandShape.Top = bodyTop
                          bandShape.Height = bodyHeight
                      End If
                      bandShape.Tags.Add shapeTag, "yes"
                      bandShape.Tags.Add "PI_ROLE", "section-band"
                      bandShape.Tags.Add "PI_SECTION", CStr(sectNum)
                      bandShape.ActionSettings(ppMouseClick).Action = ppActionHyperlink
                      bandShape.ActionSettings(ppMouseClick).Hyperlink.SubAddress = strLinkTitle
                  End If
760               Set newShape2 = renderPresentation.Slides(slideNum). _
                      Shapes.AddTextbox(msoTextOrientationHorizontal, _
                      dbllabelLeft_test, _
                      titleTop, _
                      dbllabelWidth_test, _
                      titleHeight)
770               newShape2.Tags.Add shapeTag, "yes"
                  newShape2.Tags.Add "PI_ROLE", "section-title"
                  newShape2.Tags.Add "PI_SECTION", CStr(sectNum)
                  newShape2.AlternativeText = fullSectionTitle
                      
                  ' Style the section title.
780               With newShape2.TextFrame
790                   .TextRange.Text = strSectionNameToShow
800                   .TextRange.Font.Size = theOptions.textHeight
810                   .TextRange.Font.Name = theOptions.fontName
820                   .HorizontalAnchor = msoAnchorNone
                      '.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
830                   .TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
840                   If gblbolDebug = True Then LogInfo "slide: " & slideNum & ", Section: " & strSectionNameToShow & ", align: sollte ppAlignCenter"
                      '.VerticalAnchor = msoVerticalAnchorMixed
850                   .MarginBottom = 0
860                   .MarginTop = 0
870                   .MarginLeft = 0
880                   .MarginRight = 0
                        
                      ' --- START color active/current SECTION ---
                      'slidenum is >= the first slide in this (sectnum) section and < than the first slidenumber of the next section

                      ' save current values for next REPEAT slide or set the last values, if slide is set to repeat
                      ' this is needed to format the sections when a slide is repeated
890                   If gblLastPI = 0 Then
900                       gblLastPIslideNum = slideNum
910                   Else
920                       gblLastPIslideNum = gblLastPIslideNum
930                   End If

                      ' select correct color
940                   If gblLastPIslideNum >= sectinfo.lngArySecStart(sectNum) And gblLastPIslideNum < sectinfo.lngArySecStart(sectNum + 1) Then
950                       sectionIs = 1 '0 = Seen, 1= Active, 2= Unseen
                          currentSection = sectNum
960                       .TextRange.Font.Color = lngCurTxtColF   'ColorToRGB(theOptions.XXXX)
970                       If theOptions.useTextColorBG_Current = 1 Then
980                           newShape2.Fill.BackColor.RGB = lngCurTxtColB
990                           newShape2.Line.BackColor.RGB = lngCurTxtColB
1000                          newShape2.Fill.Transparency = theOptions.transparencyFontActive / 100
1010                          newShape2.Line.Transparency = theOptions.transparencyFontActive / 100
1020                          newShape2.TextFrame2.TextRange.Font.Fill.Transparency = theOptions.transparencyFontActive / 100
1030                      End If
1040                  Else
1050                      If dotNum <= slideNumVisible Then       'dotnum - sectinfo.lngSecStartSlide + 1
1060                          sectionIs = 0 '0 = Seen, 1= Active, 2= Unseen
1070                          .TextRange.Font.Color = lngSeenTextColor   ' ColorToRGB(theOptions.seenTextColor)
1080                          If theOptions.useTextColorBG = 1 Then
1090                              newShape2.Fill.BackColor.RGB = lngTextBGColorSeen
1100                              newShape2.Line.BackColor.RGB = lngTextBGColorSeen
1110                              newShape2.Fill.Transparency = theOptions.transparencyFontSeen / 100
1120                              newShape2.Line.Transparency = theOptions.transparencyFontSeen / 100
1130                              newShape2.TextFrame2.TextRange.Font.Fill.Transparency = theOptions.transparencyFontSeen / 100
1140                          End If
1150                      Else
1160                          sectionIs = 2 '0 = Seen, 1= Active, 2= Unseen
1170                          .TextRange.Font.Color = lngUnseenTextColor   'ColorToRGB(theOptions.unseenTextColor)
1180                          If theOptions.useTextColorBG_UnSeen = 1 Then
1190                              newShape2.Fill.BackColor.RGB = lngTextBGColorUnseen
1200                              newShape2.Line.BackColor.RGB = lngTextBGColorUnseen
1210                              newShape2.Fill.Transparency = theOptions.transparencyFontUnseen / 100
1220                              newShape2.Line.Transparency = theOptions.transparencyFontUnseen / 100
1230                              newShape2.TextFrame2.TextRange.Font.Fill.Transparency = theOptions.transparencyFontUnseen / 100
1240                          End If
1250                      End If
1260                  End If
                      ' --- END color active/current SECTION ---
1270                  If gblbolDebug = True Then LogInfo ",IS: align: " & .TextRange.Paragraphs.ParagraphFormat.Alignment
1280              End With
                  If theOptions.CrimsonType = 1 And StrComp(theOptions.fontName, "Inter Medium", vbTextCompare) = 0 Then
                      With newShape2.TextFrame.TextRange.Font
                          .Bold = msoFalse
                          If sectionIs = 1 And theOptions.CurrentTitleBold = 1 Then
                              .Name = "Inter SemiBold"
                          Else
                              .Name = theOptions.fontName
                          End If
                      End With
                  ElseIf theOptions.CurrentTitleBold = 1 Then
                      newShape2.TextFrame.TextRange.Font.Bold = (sectionIs = 1)
                  End If
                  If bandMode Then
                      newShape2.Fill.Visible = msoFalse
                      newShape2.Line.Visible = msoFalse
                      newShape2.TextFrame.AutoSize = ppAutoSizeNone
                      newShape2.TextFrame.VerticalAnchor = msoAnchorMiddle
                      titlePadding = bandPadding
                      If titlePadding > dbllabelWidth_test / 4 Then titlePadding = dbllabelWidth_test / 4
                      newShape2.TextFrame.MarginLeft = titlePadding
                      newShape2.TextFrame.MarginRight = titlePadding
                      newShape2.TextFrame2.WordWrap = msoFalse
                      ' AddTextbox may have grown to fit the unformatted title.
                      newShape2.LockAspectRatio = msoFalse
                      newShape2.Left = dbllabelLeft_test
                      newShape2.Top = titleTop
                      newShape2.Width = dbllabelWidth_test
                      newShape2.Height = titleHeight
                      newShape2.TextFrame2.AutoSize = msoAutoSizeTextToFitShape
                      bandPosition = PI_BandGradientPosition(sectinfo, sectNum, gblLastPIslideNum)
                      PI_PaintSectionBackground bandShape, theOptions, sectionIs, bandPosition
                      If Not layeredMode Then
                          PI_AddHardShadow bandShape, theOptions, shapeTag, presetGrowth
                          If theOptions.PanelShadow = 1 Then PI_ApplyLayerShadow bandShape, theOptions, theOptions.LayerShadowDistance
                      End If
                      If layeredMode Then
                          PI_StyleLayerBand renderPresentation.Slides(slideNum), bandShape, _
                              theOptions, sectinfo, sectionIs, sectNum, bandHeight, _
                              railLeft, railWidth, shapeTag, strLinkTitle
                      End If
                      If splitBand Then
                          bandShape.Tags.Add "PI_BODY_PT", Trim$(Str$(bodyHeight))
                          bandShape.Tags.Add "PI_ROW_PT", Trim$(Str$(rowHeight))
                          If layeredMode Then PI_AddCompactRegion renderPresentation.Slides(slideNum), theOptions, _
                              dbllabelLeft_test, bodyTop, dbllabelWidth_test, bodyHeight, sectNum, sectionIs, _
                              bandPosition, "section-body", shapeTag, strLinkTitle, False
                          PI_AddCompactRegion renderPresentation.Slides(slideNum), theOptions, _
                              dbllabelLeft_test, rowTop, dbllabelWidth_test, rowHeight, sectNum, sectionIs, _
                              bandPosition, "section-marker-row", shapeTag, strLinkTitle, Not layeredMode
                      End If
                      PI_BandTextContrast newShape2, theOptions, sectionIs, bandPosition, slideNum
                      PI_AddBandBackdrop renderPresentation.Slides(slideNum), bandShape, newShape2, _
                          theOptions, strSectionNameToShow, sectionIs, sectNum, shapeTag, strLinkTitle
                      If sectionIs = 1 Then PI_AddSectionAccent renderPresentation.Slides(slideNum), _
                          theOptions, dbllabelLeft_test, dbllabelWidth_test, bodyHeight, sectNum, shapeTag, strLinkTitle
                  End If
                  sectionBarColor = newShape2.TextFrame.TextRange.Font.Color.RGB
                  If bandMode And theOptions.FlatRuleMode > 0 Then
                      PI_AddFlatRules renderPresentation.Slides(slideNum), theOptions, _
                          dbllabelLeft_test, dbllabelWidth_test, bodyHeight, sectNum, _
                          visibleSectionOrdinal, sectionIs, sectionBarColor, shapeTag
                  End If
                  If theOptions.CounterMode > 0 And sectionIs = 1 And _
                      theOptions.PIStyle <> "Indicator only" And theOptions.ignoreSections = "Full" Then
                      counterTop = titleTop + titleHeight
                      If splitBand Then
                          counterTop = rowTop - markerClearance
                          If theOptions.SectionPlacement <> 0 Then counterTop = dotsTop + markerRowHeight
                      End If
                      Set counter = PI_AddCounter(renderPresentation.Slides(slideNum), sectinfo, theOptions, _
                          slideNumVisible, sectNum, dbllabelLeft_test, dbllabelWidth_test, _
                          counterTop, sectionBarColor, shapeTag)
                  End If
                  collapsedSection = False
                  If flatBarMode And theOptions.PIStyle = "Full" And sectionIs <> 1 Then
                      If theOptions.DensityMode = 2 Then collapsedSection = True
                      If theOptions.DensityMode = 1 Then
                          If crimsonMarkers Then
                              renderBarWidth = PI_CrimsonMarkerWidth(theOptions, dbllabelWidth_test, numDotsInSection, barGap)
                          Else
                              barSlotWidth = dbllabelWidth_test / numDotsInSection
                              barSlotWidth = barSlotWidth * (1 - 0.05 * theOptions.PresetSize)
                              barGap = PI_BarClamp(theOptions.FlatBarGap * (1 - 0.25 * theOptions.PresetSize), 0, barSlotWidth / 2)
                              renderBarWidth = barSlotWidth - barGap
                          End If
                          collapsedSection = (renderBarWidth < PI_BarClamp(theOptions.MinMarkerLength, 2, 20))
                      End If
                  End If

                  'Add hyperlink to section title as well
1290              With newShape2.ActionSettings(ppMouseClick)
1300                  .Action = ppActionHyperlink
                      '.Hyperlink.Address = slideNum
1310                  .Hyperlink.SubAddress = strLinkTitle 'we need the active slide that gets indicator X
                  
#If Mac Then
                      ' Code here will be run only on Macs but ignored on PCs
                      'screentip crashes macro on MAC!
#Else
                      ' Code here will be run on PCs but ignored on Macs
1320                  .Hyperlink.screentip = strLinkTitle 'we need the active slide that gets indicator X
#End If
1330              End With

1340              If gblbolDebug = True Then LogInfo ", .Width: " & newShape2.Width
1350              If gblbolDebug = True Then LogInfo ", .Left: " & newShape2.Left

1360              If theOptions.ignoreSections <> "Full" Or theOptions.PIStyle = "Indicator only" Then
1370                  newShape2.visible = False
                      'newShape2.Delete
1380              End If
          
          
          
          
1390              gblPIShapeCountTotalSection = gblPIShapeCountTotalSection + 1 'one more section-shape has been drawn
                  'END NEW-----------------------------------------------------
                      
1400              If gblbolDebug = True Then
1410                  strDebugtext = strDebugtext & "# " & sectNum & "; t: " & strSectionNameToShow & "; Left: " & labelLeft & "; Width: " & labelWidth & vbCrLf
1420              End If
                      
1430          End If
              'End If
              '####### END SECTIONS ########





              '####### START INDICATORS/SHAPES ########
              '---START INDICATOR V2------------------------------------
              'calucaltions needed for Indicator V3 !

              If sizedMarkers Then
                  barAreaLeft = dbllabelLeft_test
                  barAreaWidth = dbllabelWidth_test
                  If theOptions.ignoreIndicator = "Active Section" And sectionIs = 1 And Not bandMode Then
                      barAreaLeft = theOptions.LeftMargin
                      barAreaWidth = mPIRenderCanvasWidth - theOptions.LeftMargin - theOptions.RightMargin
                  End If
                  slotCount = numDotsInSection
                  slotIndex = numDotsCount
                  If collapsedSection Then
                      slotCount = 1
                      slotIndex = 1
                  End If
                  PI_PresetMarkerGeometry theOptions, barAreaLeft, barAreaWidth, slotIndex, slotCount, _
                      CLng(dotNum) - slideNumVisible, markerRowHeight, dblLeft_newC2, renderBarWidth, renderBarHeight, barSlotWidth
                  renderDotTop = dotsTop + (markerRowHeight - renderBarHeight) / 2
              ElseIf mPIReadLayout Then
                  With mPIMarkerSlots(dotNum)
                      dblLeft_newC2 = .axisStart
                      renderDotTop = .crossStart
                      renderBarWidth = .axisLength
                      renderBarHeight = .crossLength
                      barSlotWidth = .slotLength
                  End With
              Else
1440          If theOptions.ignoreIndicator = "Active Section" And sectionIs = 1 And Not bandMode Then
1450              dblLeft_new = 0
1460              dblLeft_new = dblPageWidth_new / (numDotsInSection + 1)                                   'e.g.: labelwidth=300 with 2 indicators -> indicator should be placed at 100 and 200 (labelwidth / (2+1))
1470              dblLeft_new = dblLeft_new * numDotsCount
1480              dblLeft_new = dblLeft_new + 0 + -(theOptions.dotSize / 2) '0 to ignore dbllabelLeft_test
1490          Else

1500              dblLeft_new = 0
1510              dblLeft_new = dbllabelWidth_test / (numDotsInSection + 1)                                   'e.g.: labelwidth=300 with 2 indicators -> indicator should be placed at 100 and 200 (labelwidth / (2+1))
1520              dblLeft_new = dblLeft_new * numDotsCount
1530              dblLeft_new = dblLeft_new + dbllabelLeft_test + -(theOptions.dotSize / 2)
1540          End If


              '---END INDICATOR V2------------------------------------

              'start indicators V3 ########################################################
              'intDebugValue = theOptions.dotSize * 2             'place new (test)-indicators below the original ones

1550          dblSpaceIndicatorAdjust = theOptions.IndicatorSpaceAdjustment
1560          dblIndicatorSqueeze = theOptions.IndicatorSqueeze

1570          dbllabelWidthCalculatedBy2 = theOptions.IndicatorSpaceAdjustment * 2
1580          dbllabelWidthCalculated = dbllabelWidth_test - theOptions.IndicatorSpaceAdjustment + (dbllabelWidthCalculatedBy2 / 2)
1590          dblLeftCalculatedNew = 0

1600          dblnumDotsCountCorrector = (Int(numDotsInSection / 2) - numDotsInSection + 1)    'to calculcate the current indicator's position relativ to the middle of the section. in-section-centered indicator will have corrector of 0

1610          dblLeft_newC2 = dblLeft_new
1620          dblLeft_newC2 = dblLeft_newC2 - ((dblnumDotsCountCorrector + (numDotsCount - 1)) * dblIndicatorSqueeze)   'this corrects for the Squeeze setting!


              renderDotSize = theOptions.dotSize
              renderDotTop = dotsTop
              If flatBarMode Then
                  barAreaLeft = dbllabelLeft_test
                  barAreaWidth = dbllabelWidth_test
                  If theOptions.ignoreIndicator = "Active Section" And sectionIs = 1 And Not bandMode Then
                      barAreaLeft = theOptions.LeftMargin
                      barAreaWidth = mPIRenderCanvasWidth - theOptions.LeftMargin - theOptions.RightMargin
                  End If
                  barSlotWidth = barAreaWidth / numDotsInSection
                  If collapsedSection Then barSlotWidth = barAreaWidth
                  If barSlotWidth <= 0 Then Err.Raise vbObjectError + 3021, "DrawProgressBarOnSlide", PI_Zh("6A2A67616CA167098DB3591F76846C345E737A7A95F4FF0C8BF751CF5C0F5DE653F38FB98DDD62167AE0828295F48DDD3002")
                  barGap = PI_BarClamp(theOptions.FlatBarGap, 0, 20)
                  If barGap > barSlotWidth / 2 Then barGap = barSlotWidth / 2
                  renderBarWidth = barSlotWidth - barGap
                  renderBarHeight = markerRowHeight
                  Select Case theOptions.MarkerKind
                      Case 0
                          If renderBarHeight > renderBarWidth / 3 Then renderBarHeight = renderBarWidth / 3
                      Case 2
                          renderBarWidth = PI_BarClamp(renderBarWidth, 0, markerRowHeight * 0.32)
                          If (numDotsCount - 1) Mod 5 <> 0 Then renderBarHeight = markerRowHeight * 0.63
                      Case 3, 5
                          If renderBarHeight > renderBarWidth Then renderBarHeight = renderBarWidth
                          renderBarWidth = renderBarHeight
                      Case 4
                          renderBarWidth = barSlotWidth
                  End Select
                  renderDotTop = dotsTop + (markerRowHeight - renderBarHeight) / 2
                  dblLeft_newC2 = barAreaLeft + barSlotWidth * (numDotsCount - 1) + (barSlotWidth - renderBarWidth) / 2
              ElseIf bandMode And dbllabelWidth_test > 0 Then
                  If renderDotSize > dbllabelWidth_test Then renderDotSize = dbllabelWidth_test
                  renderDotTop = dotsTop + (theOptions.dotSize - renderDotSize) / 2
                  dotPadding = (dbllabelWidth_test - renderDotSize) / 2
                  If dotPadding > bandPadding Then dotPadding = bandPadding
                  If dotPadding < 0 Then dotPadding = 0
                  dblLeft_newC2 = dbllabelLeft_test + dotPadding + (dbllabelWidth_test - 2 * dotPadding) * numDotsCount / (numDotsInSection + 1) - renderDotSize / 2
                  dblLeft_newC2 = dblLeft_newC2 - ((dblnumDotsCountCorrector + numDotsCount - 1) * dblIndicatorSqueeze)
                  If dblLeft_newC2 < dbllabelLeft_test + dotPadding Then dblLeft_newC2 = dbllabelLeft_test + dotPadding
                  If dblLeft_newC2 > dbllabelLeft_test + dbllabelWidth_test - dotPadding - renderDotSize Then dblLeft_newC2 = dbllabelLeft_test + dbllabelWidth_test - dotPadding - renderDotSize
              End If

                  If mPIStoreLayout Then
                      With mPIMarkerSlots(dotNum)
                          .axisStart = dblLeft_newC2
                          .crossStart = renderDotTop
                          .axisLength = renderBarWidth
                          .crossLength = renderBarHeight
                          .slotLength = barSlotWidth
                      End With
                  End If
              End If

1630          If gblbolDebug = True Then LogInfo "Ignore Indicator (ddcIgnoreIndicator) is set to: " & theOptions.ignoreIndicator
1640          If theOptions.ignoreIndicator <> "None" And theOptions.PIStyle <> "Sections only" Then

                  If flatBarMode Then
                      Select Case theOptions.MarkerKind
                          Case 0: intShapetypeNr = msoShapeRoundedRectangle
                          Case 3: intShapetypeNr = msoShapeOval
                          Case Else: intShapetypeNr = msoShapeRectangle
                      End Select
                      Set indicator2 = renderPresentation.Slides(slideNum).Shapes.AddShape( _
                          intShapetypeNr, dblLeft_newC2, renderDotTop, renderBarWidth, renderBarHeight)
                      If theOptions.MarkerKind = 0 Then indicator2.Adjustments.Item(1) = PI_BarClamp(theOptions.FlatBarRoundness, 0, 50) / 100
                      If crimsonMarkers Or (theOptions.MarkerKind > 0 And Not sizedMarkers) Then
                          indicator2.Tags.Add "PI_GLYPH_SIZE", Trim$(Str$(renderBarWidth)) & "|" & _
                              Trim$(Str$(renderBarHeight)) & "|" & Trim$(Str$(barSlotWidth))
                      End If
                      indicator2.Fill.Solid
1650              ElseIf gblshowCustomImage = 0 Then
                      dotAlong = renderDotSize
                      dotAcross = renderDotSize
                      If intShapetypeNr = msoShapeOval Then
                          barAreaLeft = dbllabelLeft_test
                          barAreaWidth = dbllabelWidth_test
                          If theOptions.ignoreIndicator = "Active Section" And sectionIs = 1 And Not bandMode Then
                              barAreaLeft = theOptions.LeftMargin
                              barAreaWidth = mPIRenderCanvasWidth - theOptions.LeftMargin - theOptions.RightMargin
                          End If
                          ovalSlot = barAreaWidth / numDotsInSection
                          dotAcross = theOptions.dotSize
                          renderDotTop = dotsTop
                          If dotAcross > ovalSlot Then
                              dotAlong = ovalSlot
                              dblLeft_newC2 = barAreaLeft + (numDotsCount - 1) * ovalSlot
                          End If
                      End If
1660                  Set indicator2 = renderPresentation.Slides(slideNum).Shapes.AddShape( _
                          intShapetypeNr, _
                          dblLeft_newC2, _
                          renderDotTop, _
                          dotAlong, _
                          dotAcross)
                      If intShapetypeNr = msoShapeOval Then indicator2.Tags.Add "PI_MARKER", "oval"
1670              Else
                      ' which image to show?
1680                  Select Case dotNum
                          Case Is < slideNumVisible
1690                          strCustomImage_Current = strCustomImage_Seen
1700                      Case Is > slideNumVisible
1710                          strCustomImage_Current = strCustomImage_UnSeen
1720                      Case Is = slideNumVisible
1730                          strCustomImage_Current = strCustomImage_Active
1740                      Case Else
1750                          MsgBox PI_Zh("8BBE7F6E8FDB5EA656FE6807989C827259318D253002")
1760                  End Select

1770                  Set indicator2 = renderPresentation.Slides(slideNum).Shapes.AddPicture _
                          (strCustomImage_Current, msoCTrue, msoCTrue, _
                          dblLeft_newC2, _
                          renderDotTop, _
                          renderDotSize, _
                          renderDotSize)
1780              End If
                        
1790              With indicator2.ActionSettings(ppMouseClick)
1800                  .Action = ppActionHyperlink
                      '.Hyperlink.Address = slideNum
1810                  .Hyperlink.SubAddress = strLinkTitle 'we need the active slide that gets indicator X
                  
#If Mac Then
                      ' Code here will be run only on Macs but ignored on PCs
                      'screentip crashes macro on MAC!
#Else
                      ' Code here will be run on PCs but ignored on Macs
1820                  .Hyperlink.screentip = strLinkTitle 'we need the active slide that gets indicator X
#End If

1830              End With
                            
1840              indicator2.Tags.Add shapeTag, "yes"
                  indicator2.Tags.Add "PI_ROLE", "indicator"
                  indicator2.Tags.Add "PI_SECTION", CStr(sectNum)
                  indicator2.Tags.Add "PI_PAGE", CStr(dotNum)
                  indicator2.Tags.Add "PI_SECTION_COLOR", CStr(sectionBarColor)
                  If collapsedSection Then
                      indicator2.Tags.Add "PI_ROLE", "indicator-summary"
                      indicator2.Tags.Add "PI_COUNT", CStr(numDotsInSection)
                      indicator2.Tags.Add "PI_RANGE_START", CStr(CLng(varArray(dotNum)))
                      indicator2.Tags.Add "PI_RANGE_END", CStr(CLng(varArray(dotNum + numDotsInSection - 1)))
                      indicator2.AlternativeText = sectinfo.strArySecTitles(sectNum) & ": " & _
                          CStr(CLng(varArray(dotNum))) & "-" & CStr(CLng(varArray(dotNum + numDotsInSection - 1))) & _
                          " (" & CStr(numDotsInSection) & ")"
                  End If
                  If flatBarMode Then indicator2.Tags.Add "PI_MARKER", Choose(theOptions.MarkerKind + 1, "flat-bar", "segment", "tick", "node", "track", "block")
                    
1850              bFillShape = False     'set some default, does not matter which, but check for opposite later
              
                  'color of dot/indicator as: seen - active/current - unseen
1860              Select Case dotNum
                      Case Is < slideNumVisible
1870                      indicator2.Fill.ForeColor.RGB = lngSeenDotColor
1880                      indicator2.Line.ForeColor.RGB = lngSeenDotColor
1890                      indicator2.Fill.Transparency = theOptions.transparencyIndicatorSeen / 100
1900                      indicator2.Line.Transparency = theOptions.transparencyIndicatorSeen / 100
1910                      If theOptions.useShapeColorBG_Seen = 1 Then bFillShape = True
                      
1920                  Case Is > slideNumVisible
1930                      indicator2.Fill.ForeColor.RGB = lngUnseenDotColor
1940                      indicator2.Line.ForeColor.RGB = lngUnseenDotColor
1950                      indicator2.Fill.Transparency = theOptions.transparencyIndicatorUnseen / 100
1960                      indicator2.Line.Transparency = theOptions.transparencyIndicatorUnseen / 100
1970                      If theOptions.useShapeColorBG_UnSeen = 1 Then bFillShape = True
                      
1980                  Case Is = slideNumVisible
1990                      indicator2.Fill.ForeColor.RGB = lngCurDotColF
2000                      indicator2.Line.ForeColor.RGB = lngCurDotColB
2010                      indicator2.Fill.Transparency = theOptions.transparencyIndicatorActive / 100
2020                      indicator2.Line.Transparency = theOptions.transparencyIndicatorActive / 100
2030                      If theOptions.useShapeColorBG_Current = 1 Then bFillShape = True
                      
2040                  Case Else
2050                      MsgBox PI_Zh("8BBE7F6E8FDB5EA656FE6807989C827259318D253002")
2060              End Select

                  'style line
2070              indicator2.Line.Weight = theOptions.ShapeLineWeight
2080              indicator2.Line.Style = theOptions.ShapeLineStyle
2090              indicator2.Line.DashStyle = theOptions.ShapeLineDashStyle

                  'make outline of shapes completele invible, hint from Temmink K. 2022-02-04
2100              If theOptions.ShapeLineWeight = 0 Then
2110                  indicator2.Line.visible = False
2120              Else
2130                  indicator2.Line.visible = True
2140              End If

                  'Fill indicator if needed (to create indicators with outline only)
2150              indicator2.Fill.visible = bFillShape
                  If bandMode Then PI_BandDotContrast indicator2, theOptions, sectionIs, bandPosition, slideNum
                  If flatBarMode Then
                      indicator2.Fill.Visible = msoTrue
                      indicator2.Line.Visible = msoFalse
                      If theOptions.SectionForegroundBars = 1 Then
                          indicator2.Fill.ForeColor.RGB = sectionBarColor
                          indicator2.Line.ForeColor.RGB = sectionBarColor
                      End If
                      If theOptions.FutureBarOutline = 1 And dotNum > slideNumVisible Then PI_OutlineFutureBar indicator2
                      ' Keep markers neutral while grouping and docking can store hidden effects.
                      PI_BarApplyGlow indicator2, theOptions, False
                  End If
       
2160              gblPIShapeCountTotalShape = gblPIShapeCountTotalShape + 1 'one new indicator shape has been created
2170              intDebugValue = 0
                  'END indicators V3 ########################################################
                  
                  ' v192 if theOptions.ignoreIndicator = "Inactive Sections" we go ahead and make the currently drawn indicator invisible .. if it's not the active section
2180              If theOptions.ignoreIndicator = "Active Section" And sectionIs <> 1 Then
                      'indicator2.visible = False
2190                  indicator2.Delete
2200              End If
                  
2210              If theOptions.PIStyle = "Sections only" Then
2220                  indicator2.visible = False
2230              End If
                  
                  
2240          End If  'theOptions.ignoreIndicator

              '####### END INDICATORS/SHAPES ########
              If collapsedSection Or theOptions.PIStyle = "Sections only" Then
                  dotNum = dotNum + numDotsInSection - 1
                  numDotsCount = numDotsInSection
              End If

              ' move to the next (visible!) section if we have all drawn indicators in the current one
2250          If numDotsCount = numDotsInSection Then
2260              numDotsCount = 0    'reset the counted dots for the new section

                  'numDotsCountInSection = 0  'reset the counted dots for the new section
2270              sectNum_old = sectNum
2280              sectNum = sectNum + 1   'move to the next section
                  
2290              While sectinfo.intArySecVisible(sectNum) = 0 And sectNum < numSectionsTotal      'if this section is invisible, continue moving forward to the next until a visible is found
2300                  If gblbolDebug = True Then LogInfo "no slide to show in section: " & sectNum & " " & sectinfo.strArySecTitles(sectNum)
2310                  sectNum = sectNum + 1
2320              Wend
2330          Else
                  'nothing to do here anymore
2340          End If

2350      Next dotNum
    If Len(mPITraceFile) > 0 Then PI_Trace "markers.complete"
          If layeredMode Then
              PI_ClipLayeredBackgrounds renderPresentation.Slides(slideNum), firstShape, shapeTag, _
                  theOptions.yMargin + bandHeight + PI_BarClamp(theOptions.RailThickness, 0.5, 30)
          ElseIf bandMode And theOptions.PanelShadow = 1 Then
              PI_ClipLayeredBackgrounds renderPresentation.Slides(slideNum), firstShape, shapeTag, theOptions.yMargin + bandHeight
          End If
          If layeredMode Or (bandMode And (theOptions.AccentMode > 0 Or theOptions.PanelShadow = 1)) Then
    If Len(mPITraceFile) > 0 Then PI_Trace "zorder.begin"
              PI_OrderLayeredShapes renderPresentation.Slides(slideNum), firstShape, shapeTag, numSections
          End If
                    

                    
          'START Group all the new shapes.
          Dim numShapes, i As Long
          Dim oShapeArray() As Long
          Dim oGroupShape As Shape
2360      numShapes = renderPresentation.Slides(slideNum).Shapes.Count - firstShape + 1
2370      If numShapes > 1 Then
2380          ReDim oShapeArray(0 To numShapes - 1)
2390          For i = 0 To numShapes - 1
2400              oShapeArray(i) = firstShape + i
2410          Next i
    If Len(mPITraceFile) > 0 Then PI_Trace "navigation.group"
2420          Set oGroupShape = renderPresentation.Slides(slideNum).Shapes.Range(oShapeArray).Group
2430          oGroupShape.Tags.Add shapeTag, "yes"
              If layeredMode Then PI_DisableLayerGroupShadow oGroupShape
              If Len(mPITraceFile) > 0 Then PI_Trace "navigation.dock"
              PI_ApplyDock oGroupShape, renderPresentation.Slides(slideNum), theOptions
              If Len(mPITraceFile) > 0 Then PI_Trace "navigation.text"
              PI_FinalizeNavigationText oGroupShape, theOptions
              If Len(mPITraceFile) > 0 Then PI_Trace "navigation.glow"
              PI_FinalizeNavigationGlow oGroupShape, theOptions, CLng(slideNumVisible), currentSection
              Set mPILastDrawResult = oGroupShape
          ElseIf numShapes = 1 Then
              Set oGroupShape = renderPresentation.Slides(slideNum).Shapes(firstShape)
              If Len(mPITraceFile) > 0 Then PI_Trace "navigation.dock"
              PI_ApplyDock oGroupShape, renderPresentation.Slides(slideNum), theOptions
              If Len(mPITraceFile) > 0 Then PI_Trace "navigation.text"
              PI_FinalizeNavigationText oGroupShape, theOptions
              If Len(mPITraceFile) > 0 Then PI_Trace "navigation.glow"
              PI_FinalizeNavigationGlow oGroupShape, theOptions, CLng(slideNumVisible), currentSection
              Set mPILastDrawResult = oGroupShape
2440      End If
          'END Group all the new shapes.
          


2450      If gblbolDebug = True Then
2460          LogInfo "DrawProgressBarOnSlide, Debug, Slide: " & slideNum & "; Value: " & strDebugtext
2470      End If
             
exit_proc:
2480      Exit Sub
          
ErrorHandler:
          If mPIRunActive And gblshowCustomImage <> 0 And Not flatBarMode Then
              If Err.Number = -2147024809 And Erl >= 1870 And Erl <= 2150 Then Resume Next
          End If
          If Err.Number = vbObjectError + 3020 Or Err.Number = vbObjectError + 3021 Or _
              Err.Number = vbObjectError + 3022 Or Err.Number = vbObjectError + 3023 Or _
              Err.Number = vbObjectError + 3024 Or mPIRunActive Or flatBarMode Or layeredMode Then
              backdropErrorNumber = Err.Number
              backdropErrorSource = Err.Source
              backdropErrorDescription = Err.Description
              PI_FirstError "DrawProgressBarOnSlide", Erl, Err.Number, Err.Source, Err.Description
              On Error Resume Next
              If firstShape < 1 Then firstShape = renderPresentation.Slides(slideNum).Shapes.Count + 1
              For i = renderPresentation.Slides(slideNum).Shapes.Count To firstShape Step -1
                  If UCase$(renderPresentation.Slides(slideNum).Shapes(i).Tags(shapeTag)) = "YES" Then
                      renderPresentation.Slides(slideNum).Shapes(i).Delete
                  End If
              Next i
              On Error GoTo 0
              Err.Raise backdropErrorNumber, backdropErrorSource, backdropErrorDescription
          End If
2490      If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
2500      Select Case Err.Number
              Case -2147024809
                  'shape can not use setting
2510              Resume Next

2520          Case -2147188160      'no active presentation
2530              MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "DrawProgressBarOnSlide: Nothing to do!"
2540              Resume exit_proc
2550              Resume
              
2560          Case 424
                  'example: negative shape size
2570              MsgBox PI_Zh("8F935165503C65E065483002") & vbCrLf & PI_Zh("63D24EF65DF291CD7F6E8BE598794EE56062590D8FD0884C3002") & vbCrLf & PI_Zh("8BF791CD65B08F93516567096548503C3002"), vbOKOnly Or vbCritical Or vbSystemModal, PI_Zh("8F93516565E06548")
2580              Resume Next
              
              
2590          Case Else
2600              MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "DrawProgressBarOnSlide"
2610              Resume Next
2620              Resume
              
2630      End Select
                              
End Sub

' ----------------------------------------------------------------
' Procedure Name: MakeProgressBar
' Purpose: create the indicator bar
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter theOptions (ProgressDotsOptions): options to use (color, font, shape etc.)
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Private Sub PI_BindNativeTarget(ByRef fields() As String)
    Dim index As Long, pair() As String, active As Presentation
#If Mac Then
    Const pathCompare As Long = vbBinaryCompare
#Else
    Const pathCompare As Long = vbTextCompare
#End If
    Set active = Application.ActivePresentation
    For index = 7 To UBound(fields)
        pair = Split(fields(index), "=", 2)
        If pair(0) = "targetDeck" Then
            If Len(active.Path) = 0 Or StrComp(active.FullName, pair(1), pathCompare) <> 0 Then
                Err.Raise vbObjectError + 3088, "PI_BindNativeTarget", _
                    "Expected active presentation: " & pair(1) & ". Actual: " & active.FullName
            End If
        End If
    Next index
    Set mPINativePresentation = active
    PI_RequireNativeTarget
End Sub

Private Function PI_NativeTargetActive() As Boolean
    If Not mPISilent Then
        PI_NativeTargetActive = True
    ElseIf Application.Windows.Count > 0 Then
        PI_NativeTargetActive = (Application.ActivePresentation Is mPINativePresentation)
    End If
End Function

Private Sub PI_RequireNativeTarget()
    If Not PI_NativeTargetActive() Then
        Err.Raise vbObjectError + 3088, "PI_RequireNativeTarget", _
            "The active presentation changed. Generation stopped."
    End If
End Sub

Private Function PI_RenderPresentation() As Presentation
    If Not (mPIRunPresentation Is Nothing) Then
        Set PI_RenderPresentation = mPIRunPresentation
    ElseIf Not (mPINativePresentation Is Nothing) Then
        Set PI_RenderPresentation = mPINativePresentation
    Else
        Set PI_RenderPresentation = ActivePresentation
    End If
End Function

Private Sub PI_PrepareRenderWindow()
    Dim selected As SlideRange, current As Slide, index As Long
    PI_RequireNativeTarget
    Set mPIRenderWindow = ActiveWindow
    If mPISilent Then
        If Not (mPIRenderWindow.Presentation Is mPINativePresentation) Then
            Err.Raise vbObjectError + 3088, "PI_PrepareRenderWindow", "The active window belongs to a different presentation."
        End If
    End If
    mPIOriginalView = mPIRenderWindow.ViewType
    mPIOriginalSlideId = 0
    mPIOriginalSelectionCount = 0
    Erase mPIOriginalSelection
    On Error Resume Next
    Set current = mPIRenderWindow.View.Slide
    Set selected = mPIRenderWindow.Selection.SlideRange
    On Error GoTo 0
    If Not current Is Nothing Then mPIOriginalSlideId = current.SlideID
    If Not selected Is Nothing Then
        mPIOriginalSelectionCount = selected.Count
        ReDim mPIOriginalSelection(1 To mPIOriginalSelectionCount)
        For index = 1 To mPIOriginalSelectionCount
            mPIOriginalSelection(index) = selected(index).SlideID
        Next index
        If mPIOriginalSlideId = 0 Then mPIOriginalSlideId = selected(1).SlideID
    End If
    mPIWindowPrepared = True
    mPIRenderWindow.ViewType = ppViewSlide
    If mPIOriginalSlideId > 0 Then
        mPIRenderWindow.View.GotoSlide mPIRunPresentation.Slides.FindBySlideID(mPIOriginalSlideId).SlideIndex
    Else
        mPIRenderWindow.View.GotoSlide 1
    End If
End Sub

Private Sub PI_RestoreRenderWindow()
    Dim index As Long, selectedIndexes() As Long
    On Error Resume Next
    If mPIWindowPrepared And PI_NativeTargetActive() Then
        mPIRenderWindow.Activate
        mPIRenderWindow.ViewType = ppViewSlide
        If mPIOriginalSlideId > 0 Then
            mPIRenderWindow.View.GotoSlide mPIRunPresentation.Slides.FindBySlideID(mPIOriginalSlideId).SlideIndex
        End If
        mPIRenderWindow.ViewType = mPIOriginalView
        Select Case mPIOriginalView
            Case ppViewNormal, ppViewSlideSorter, ppViewOutline, ppViewThumbnails
                If mPIOriginalSelectionCount > 0 Then
                    ReDim selectedIndexes(1 To mPIOriginalSelectionCount)
                    For index = 1 To mPIOriginalSelectionCount
                        selectedIndexes(index) = mPIRunPresentation.Slides.FindBySlideID(mPIOriginalSelection(index)).SlideIndex
                    Next index
                    mPIRunPresentation.Slides.Range(selectedIndexes).Select
                End If
        End Select
    End If
    mPIWindowPrepared = False
    mPIOriginalSelectionCount = 0
    Erase mPIOriginalSelection
    Set mPIRenderWindow = Nothing
    Set mPIRunPresentation = Nothing
    On Error GoTo 0
End Sub

Private Sub PI_YieldRender()
    If Len(mPITraceFile) > 0 Then PI_Trace "events.pump.begin"
    DoEvents
    PI_RequireNativeTarget
    If Len(mPITraceFile) > 0 Then PI_Trace "events.pump.end"
    If mPIRenderWindow Is Nothing Then Exit Sub
    If Not (ActiveWindow Is mPIRenderWindow) Then mPIRenderWindow.Activate
    If mPIRenderWindow.ViewType <> ppViewSlide Then mPIRenderWindow.ViewType = ppViewSlide
    If mPIRunPresentation.Slides.Count <> mPISlideCount Then
        Err.Raise vbObjectError + 3024, "PI_YieldRender", PI_Zh("751F62108FC77A0B4E2D9875976265705DF253D8531630028BF791CD65B0751F62105BFC822A3002")
    End If
End Sub

Private Sub PI_BeginRenderRun(ByVal presentation As Presentation)
    Dim index As Long, commentIndex As Long, targetSlide As Slide, title As String, message As String
    mPIRunActive = False
    Set mPILastDrawResult = Nothing
    mPIComputedPages = 0
    mPILayoutHits = 0
    mPISurfaceHits = 0
    mPICaptures = 0
    mPIMetadataReads = 0
    mPIRunStarted = Timer
    mPISlideCount = presentation.Slides.Count
    ReDim mPILinksBySlide(1 To mPISlideCount)
    ReDim mPIHiddenBySlide(1 To mPISlideCount)
    ReDim mPISkippedBySlide(1 To mPISlideCount)
    ReDim mPIRepeatBySlide(1 To mPISlideCount)
    mPIRunActive = True
    For index = 1 To mPISlideCount
        Set targetSlide = presentation.Slides(index)
        title = "Slide " & CStr(targetSlide.SlideIndex)
        If targetSlide.Shapes.HasTitle Then title = targetSlide.Shapes.Title.TextFrame.TextRange.Text
        If Len(title) = 0 Then title = "-"
        mPILinksBySlide(index) = CStr(targetSlide.SlideID) & "," & CStr(targetSlide.SlideIndex) & "," & title
        mPIHiddenBySlide(index) = targetSlide.SlideShowTransition.Hidden
        For commentIndex = 1 To targetSlide.Comments.Count
            message = targetSlide.Comments.Item(commentIndex).Text
            If message = gblstrMessageIgnore Then mPISkippedBySlide(index) = True
            If message = gblstrMessageRepeat Then mPIRepeatBySlide(index) = True
        Next commentIndex
        mPIMetadataReads = mPIMetadataReads + 1
    Next index
End Sub

Private Function PI_RenderElapsed(ByVal started As Double) As Double
    PI_RenderElapsed = Timer - started
    If PI_RenderElapsed < 0 Then PI_RenderElapsed = PI_RenderElapsed + 86400
End Function

Private Sub PI_EndRenderRun(ByVal cancelled As Boolean, ByVal failureDescription As String)
    Dim status As String
    If Not mPIRunActive Then Exit Sub
    status = PI_Zh("5B8C6210")
    If cancelled Then status = PI_Zh("5DF2505C6B62")
    If Len(failureDescription) > 0 Then status = PI_Zh("53D1751F95198BEF")
    mPILastSummary = PI_Zh("72B66001FF1A") & status & vbCrLf & _
        PI_Zh("5B9E9645801765F6FF1A") & Format$(PI_RenderElapsed(mPIRunStarted), "0.00") & " s" & vbCrLf & _
        PI_Zh("8BA17B975E035C4098756570FF1A") & CStr(mPIComputedPages) & vbCrLf & _
        PI_Zh("590D75285E035C4098756570FF1A") & CStr(mPILayoutHits) & vbCrLf & _
        PI_Zh("590D752896345F7156FE6B216570FF1A") & CStr(mPISurfaceHits) & vbCrLf & _
        PI_Zh("0050004E00470020635583B76B216570FF1A") & CStr(mPICaptures) & vbCrLf & _
        PI_Zh("98848BFB9875976251436570636EFF1A") & CStr(mPIMetadataReads)
    If Len(failureDescription) > 0 Then mPILastSummary = mPILastSummary & vbCrLf & failureDescription
    mPIRunActive = False
    PI_ClearLayoutCache
    If Len(mPISurfaceDirectory) > 0 Then
        RmDir mPISurfaceDirectory
        mPISurfaceDirectory = ""
    End If
    Set mPILastDrawResult = Nothing
    Erase mPILinksBySlide
    Erase mPIHiddenBySlide
    Erase mPISkippedBySlide
    Erase mPIRepeatBySlide
    If gblbolDebug Then
        On Error Resume Next
        LogInfo Replace(mPILastSummary, vbCrLf, "; ")
        On Error GoTo 0
    End If
End Sub

Public Sub PI_ShowRenderStats(control As IRibbonControl)
    If Len(mPILastSummary) = 0 Then
        MsgBox PI_Zh("672C6B2100200050006F0077006500720050006F0069006E007400204F1A8BDD5C1A672A751F62105BFC822A3002"), vbInformation, PI_Zh("751F62107EDF8BA1")
    Else
        MsgBox mPILastSummary, vbInformation, PI_Zh("4E0A6B21751F62107EDF8BA1")
    End If
End Sub

Public Function PI_CounterRowHeight(opts As ProgressDotsOptions) As Double
    If opts.CounterMode < 1 Or opts.CounterMode > 5 Then Exit Function
    If opts.PIStyle = "Indicator only" Or opts.ignoreSections <> "Full" Then Exit Function
    PI_CounterRowHeight = PI_BarClamp(opts.textHeight * 0.65, 7, 10) * 1.25 + 2
End Function

Private Function PI_CounterText(sections As ProgressDotsSectionInfo, opts As ProgressDotsOptions, _
                                ByVal visibleIndex As Long, ByVal sectionIndex As Long) As String
    Dim index As Long, before As Long, sectionTotal As Long, sectionPage As Long, total As Long
    total = sections.lngSecIndicatorsRecalculated
    If total < 1 Or sectionIndex < 1 Or sectionIndex > sections.lngSecSections Then Exit Function
    For index = 1 To sectionIndex - 1
        before = before + sections.intArySecTotal(index) - CharCount(sections.strArySecTotalInvisible(index), ";")
    Next index
    sectionTotal = sections.intArySecTotal(sectionIndex) - CharCount(sections.strArySecTotalInvisible(sectionIndex), ";")
    If sectionTotal < 1 Then Exit Function
    sectionPage = CLng(PI_BarClamp(visibleIndex - before, 1, sectionTotal))
    Select Case opts.CounterMode
        Case 1: PI_CounterText = CStr(visibleIndex) & " / " & CStr(total)
        Case 2: PI_CounterText = Format$(visibleIndex / total, "0%")
        Case 3: PI_CounterText = CStr(sectionPage) & " / " & CStr(sectionTotal)
        Case 4: PI_CounterText = Format$(sectionPage / sectionTotal, "0%")
        Case 5: PI_CounterText = CStr(visibleIndex) & " / " & CStr(total) & " | " & CStr(sectionPage) & " / " & CStr(sectionTotal)
    End Select
End Function

Private Function PI_AddCounter(ByVal targetSlide As Slide, sections As ProgressDotsSectionInfo, _
                                opts As ProgressDotsOptions, ByVal visibleIndex As Long, ByVal sectionIndex As Long, _
                                ByVal alongStart As Double, ByVal alongLength As Double, ByVal top As Double, _
                                ByVal color As Long, ByVal shapeTag As String) As Shape
    Dim counter As Shape, padding As Double
    Set counter = targetSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, alongStart, top, alongLength, PI_CounterRowHeight(opts))
    counter.Tags.Add shapeTag, "yes"
    counter.Tags.Add "PI_ROLE", "section-counter"
    counter.Tags.Add "PI_SECTION", CStr(sectionIndex)
    counter.Tags.Add "PI_COUNTER_START", CStr(alongStart)
    counter.Tags.Add "PI_COUNTER_LENGTH", CStr(alongLength)
    counter.Fill.Visible = msoFalse
    counter.Line.Visible = msoFalse
    padding = PI_BarClamp(opts.SectionBandPadding, 0, alongLength / 4)
    With counter.TextFrame2
        .TextRange.Text = PI_CounterText(sections, opts, visibleIndex, sectionIndex)
        .TextRange.Font.Name = opts.fontName
        .TextRange.Font.Size = PI_BarClamp(opts.textHeight * 0.65, 7, 10)
        .TextRange.Font.Fill.ForeColor.RGB = color
        .TextRange.Font.Bold = msoFalse
        .MarginTop = 0
        .MarginBottom = 0
        .MarginLeft = padding
        .MarginRight = padding
        .WordWrap = msoFalse
        .VerticalAnchor = msoAnchorMiddle
        .AutoSize = msoAutoSizeTextToFitShape
    End With
    Set PI_AddCounter = counter
End Function

Public Sub PI_DockCounter(ByVal counter As Shape, opts As ProgressDotsOptions, _
                          ByVal pageWidth As Double, ByVal bandHeight As Double)
    Dim alongStart As Double, alongLength As Double, padding As Double, alongPadding As Double
    Dim rowHeight As Double, markerWidth As Double, gap As Double, textWidth As Double
    alongStart = CDbl(counter.Tags("PI_COUNTER_START"))
    alongLength = CDbl(counter.Tags("PI_COUNTER_LENGTH"))
    padding = PI_BarClamp(opts.SectionBandPadding, 0, bandHeight / 4)
    alongPadding = PI_BarClamp(padding, 0, alongLength / 4)
    rowHeight = PI_BarClamp(PI_CounterRowHeight(opts), 0, alongLength / 3)
    markerWidth = PI_BarClamp(PI_MarkerHeight(opts), 0, bandHeight / 3)
    gap = PI_BarClamp(opts.SectionIndicatorGap, 0, bandHeight / 3)
    If LCase$(opts.PIStyle) = "sections only" Then
        markerWidth = 0
        gap = 0
    End If
    textWidth = bandHeight - 2 * padding - markerWidth - gap
    If textWidth < 12 Then textWidth = 12
    counter.Rotation = 0
    counter.LockAspectRatio = msoFalse
    counter.Width = textWidth
    counter.Height = rowHeight
    counter.Top = alongStart + alongLength - rowHeight - alongPadding
    If opts.DockEdge = 2 Then
        counter.Left = opts.yMargin + padding
    Else
        counter.Left = pageWidth - opts.yMargin - bandHeight + padding + markerWidth + gap
    End If
    With counter.TextFrame2
        .Orientation = msoTextOrientationHorizontal
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeTextToFitShape
    End With
End Sub

Private Sub PI_FinalizeNavigationText(ByVal navigation As Shape, opts As ProgressDotsOptions)
    Dim index As Long, part As Shape
    If navigation.Type = msoGroup Then
        For index = 1 To navigation.GroupItems.Count
            Set part = navigation.GroupItems.Item(index)
            PI_FinalizeTextShape part, opts
        Next index
    Else
        PI_FinalizeTextShape navigation, opts
    End If
End Sub

Private Sub PI_FinalizeTextShape(ByVal title As Shape, opts As ProgressDotsOptions)
    Dim role As String, original As String, lower As Long, upper As Long, middle As Long
    Dim compactLine As Boolean
    role = LCase$(title.Tags("PI_ROLE"))
    If role <> "section-title" And role <> "section-counter" Then Exit Sub
    Select Case opts.TitleAlignment
        Case 1: title.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
        Case 2: title.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignRight
        Case Else: title.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    End Select
    If role = "section-counter" Then Exit Sub
    original = title.AlternativeText
    If opts.TitleCase = 1 Then original = UCase$(original)
    compactLine = (opts.PresetSize = 2 And opts.DockEdge < 2 And PI_UsesSectionBand(opts))
    If compactLine Then
        original = Replace(original, vbCrLf, " ")
        original = Replace(Replace(Replace(original, vbCr, " "), vbLf, " "), Chr$(11), " ")
    End If
    title.TextFrame2.TextRange.Text = original
    If Len(original) = 0 Then Exit Sub
    If opts.TitleOverflow = 0 Then
        If title.TextFrame2.AutoSize <> msoAutoSizeTextToFitShape Then Exit Sub
    Else
        title.TextFrame2.WordWrap = (opts.DockEdge >= 2)
    End If
    title.TextFrame2.AutoSize = msoAutoSizeNone
    title.TextFrame2.TextRange.Font.Size = opts.textHeight
    If compactLine Then
        ' Fixed leading in PowerPoint places glyph ascenders above the 7 pt row.
        If title.Height <= opts.textHeight + 0.01 Then
            title.TextFrame2.VerticalAnchor = msoAnchorTop
            title.TextFrame2.MarginTop = opts.textHeight / 7
        End If
        With title.TextFrame2.TextRange.ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineRuleWithin = msoFalse
            .SpaceWithin = title.Height
        End With
    End If
    If opts.TitleOverflow <> 2 Then
        PI_FitFullTitle title, compactLine
        Exit Sub
    End If
    If PI_TitleFits(title, compactLine) Then Exit Sub
    lower = 0
    upper = Len(original)
    Do While lower < upper
        middle = (lower + upper + 1) \ 2
        title.TextFrame2.TextRange.Text = PI_SafeTextPrefix(original, middle) & ChrW$(&H2026)
        If PI_TitleFits(title, compactLine) Then
            lower = middle
        Else
            upper = middle - 1
        End If
    Loop
    title.TextFrame2.TextRange.Text = PI_SafeTextPrefix(original, lower) & ChrW$(&H2026)
    If Not PI_TitleFits(title, compactLine) Then PI_FitFullTitle title, compactLine
End Sub

Private Sub PI_FitFullTitle(ByVal title As Shape, ByVal compactLine As Boolean)
    Dim availableWidth As Double, availableHeight As Double, textScale As Double
    Dim lower As Long, upper As Long, middle As Long, estimate As Long
    With title.TextFrame2
        availableWidth = title.Width - .MarginLeft - .MarginRight
        availableHeight = title.Height - .MarginTop - .MarginBottom
        If PI_TitleFits(title, compactLine) Then Exit Sub
        textScale = 1
        If .TextRange.BoundWidth > availableWidth Then textScale = availableWidth / .TextRange.BoundWidth
        If Not compactLine Then
            If .TextRange.BoundHeight * textScale > availableHeight Then textScale = availableHeight / .TextRange.BoundHeight
        End If
        ' The requested size already overflowed. Search quarter-point sizes below it.
        upper = -Int(-.TextRange.Font.Size * 4) - 1
        estimate = Fix(.TextRange.Font.Size * textScale * 4)
        If estimate < 4 Then estimate = 4
        .TextRange.Font.Size = estimate / 4
        If PI_TitleFits(title, compactLine) Then
            ' A smaller font can use fewer lines than the original measurement.
            lower = estimate
        Else
            lower = 4
            upper = estimate - 1
        End If
        Do While lower < upper
            middle = (lower + upper + 1) \ 2
            .TextRange.Font.Size = middle / 4
            If PI_TitleFits(title, compactLine) Then
                lower = middle
            Else
                upper = middle - 1
            End If
        Loop
        .TextRange.Font.Size = lower / 4
    End With
End Sub

Private Function PI_TitleFits(ByVal title As Shape, ByVal compactLine As Boolean) As Boolean
    Dim availableWidth As Double, availableHeight As Double
    With title.TextFrame2
        availableWidth = title.Width - .MarginLeft - .MarginRight
        availableHeight = title.Height - .MarginTop - .MarginBottom
        PI_TitleFits = (.TextRange.BoundWidth <= availableWidth + 0.1 And _
            (compactLine Or .TextRange.BoundHeight <= availableHeight + 0.1))
    End With
End Function

Private Function PI_SafeTextPrefix(ByVal value As String, ByVal count As Long) As String
    Dim lastCode As Long
    If count < 1 Then Exit Function
    If count < Len(value) Then
        lastCode = AscW(Mid$(value, count, 1)) And &HFFFF&
        If lastCode >= &HD800& And lastCode <= &HDBFF& Then count = count - 1
    End If
    PI_SafeTextPrefix = Left$(value, count)
End Function

Private Sub PI_AddFlatRules(ByVal targetSlide As Slide, opts As ProgressDotsOptions, _
                            ByVal left As Double, ByVal width As Double, ByVal height As Double, _
                            ByVal sectionIndex As Long, ByVal ordinal As Long, ByVal state As Integer, _
                            ByVal sectionColor As Long, ByVal shapeTag As String)
    Dim color As Long, weight As Double, inset As Double, x0 As Double, x1 As Double
    Dim y0 As Double, y1 As Double, arm As Double, index As Long, x As Double, y As Double
    Dim frame As Shape
    If width <= 0 Or height <= 0 Then Exit Sub
    weight = PI_BarClamp(opts.FlatRuleWidth, 0.25, 3)
    If weight > width / 4 Then weight = width / 4
    If weight > height / 4 Then weight = height / 4
    color = ColorToRGB(opts.FlatRuleColor)
    If opts.UseThemeColors = 1 Then color = sectionColor
    inset = PI_BarClamp(weight / 2 + 0.5, 0, width / 4)
    x0 = left + inset
    x1 = left + width - inset
    y0 = opts.yMargin + inset
    y1 = opts.yMargin + height - inset
    arm = PI_BarClamp(width / 6, 0, 8)
    If arm > height / 4 Then arm = height / 4
    If ordinal > 1 And (opts.FlatRuleMode = 2 Or opts.FlatRuleMode = 3) Then
        PI_AddRuleLine targetSlide, x0, y0, x0, y1, weight, color, sectionIndex, shapeTag
    End If
    Select Case opts.FlatRuleMode
        Case 8
            ' A spine and short return recall the end of an index card.
            PI_AddRuleLine targetSlide, x0, y0, x0, y1, weight, color, sectionIndex, shapeTag
            PI_AddRuleLine targetSlide, x0, y0, x0 + arm, y0, weight, color, sectionIndex, shapeTag
        Case 1
            If ordinal > 1 Then PI_AddRuleLine targetSlide, x0, y0, x0, y1, weight, color, sectionIndex, shapeTag
        Case 2
            PI_AddRuleLine targetSlide, x0, y0, x1, y0, weight, color, sectionIndex, shapeTag
            y = PI_BarClamp(y0 + weight + 2, y0, y1)
            PI_AddRuleLine targetSlide, x0, y, x1, y, weight, color, sectionIndex, shapeTag
        Case 3
            PI_AddRuleLine targetSlide, x0, y0, x1, y0, weight, color, sectionIndex, shapeTag
        Case 4
            Set frame = targetSlide.Shapes.AddShape(msoShapeRectangle, x0, y0, x1 - x0, y1 - y0)
            frame.Tags.Add shapeTag, "yes"
            frame.Tags.Add "PI_ROLE", "section-rule"
            frame.Tags.Add "PI_SECTION", CStr(sectionIndex)
            frame.Fill.Visible = msoFalse
            frame.Shadow.Visible = msoFalse
            frame.Line.Visible = msoTrue
            frame.Line.Weight = weight
            frame.Line.ForeColor.RGB = color
            frame.Line.DashStyle = msoLineSolid
            frame.Line.Style = msoLineSingle
            frame.Line.Transparency = 0
        Case 5
            For index = 0 To 3
                If index Mod 2 = 0 Then x = x0 Else x = x1
                If index < 2 Then y = y0 Else y = y1
                If index Mod 2 = 0 Then
                    PI_AddRuleLine targetSlide, x, y, x + arm, y, weight, color, sectionIndex, shapeTag
                Else
                    PI_AddRuleLine targetSlide, x - arm, y, x, y, weight, color, sectionIndex, shapeTag
                End If
                If index < 2 Then
                    PI_AddRuleLine targetSlide, x, y, x, y + arm, weight, color, sectionIndex, shapeTag
                Else
                    PI_AddRuleLine targetSlide, x, y - arm, x, y, weight, color, sectionIndex, shapeTag
                End If
            Next index
        Case 6
            If ordinal > 1 Then
                arm = PI_BarClamp(width / 8, 0, 3)
                For index = 0 To 2
                    x = x0 + index * arm
                    PI_AddRuleLine targetSlide, x, y0, x, y1, weight, color, sectionIndex, shapeTag
                Next index
            End If
        Case 7
            arm = PI_BarClamp(arm / 2, 0, 3)
            For index = 0 To 3
                If index Mod 2 = 0 Then x = x0 + arm Else x = x1 - arm
                If index < 2 Then y = y0 + arm Else y = y1 - arm
                PI_AddRuleLine targetSlide, x - arm, y, x + arm, y, weight, color, sectionIndex, shapeTag
                PI_AddRuleLine targetSlide, x, y - arm, x, y + arm, weight, color, sectionIndex, shapeTag
            Next index
    End Select
End Sub

Private Sub PI_AddRuleLine(ByVal targetSlide As Slide, ByVal x0 As Double, ByVal y0 As Double, _
                           ByVal x1 As Double, ByVal y1 As Double, ByVal weight As Double, _
                           ByVal color As Long, ByVal sectionIndex As Long, ByVal shapeTag As String)
    Dim rule As Shape
    Set rule = targetSlide.Shapes.AddLine(x0, y0, x1, y1)
    rule.Tags.Add shapeTag, "yes"
    rule.Tags.Add "PI_ROLE", "section-rule"
    rule.Tags.Add "PI_SECTION", CStr(sectionIndex)
    rule.Line.Visible = msoTrue
    rule.Line.Weight = weight
    rule.Line.ForeColor.RGB = color
    rule.Line.DashStyle = msoLineSolid
    rule.Line.Style = msoLineSingle
    rule.Line.Transparency = 0
    rule.Line.BeginArrowheadStyle = msoArrowheadNone
    rule.Line.EndArrowheadStyle = msoArrowheadNone
    rule.Shadow.Visible = msoFalse
End Sub

Private Function PI_ReuseKey(ByVal targetSlide As Slide, sections As ProgressDotsSectionInfo, _
                             opts As ProgressDotsOptions) As String
    Dim index As Long, currentSection As Long
    For index = 1 To sections.lngSecSections
        If gblLastPIslideNum >= sections.lngArySecStart(index) Then currentSection = index
    Next index
    PI_ReuseKey = SerializeProgressOptions(opts) & "|" & CStr(currentSection) & "|" & _
        CStr(targetSlide.Design.Index) & ":" & CStr(targetSlide.CustomLayout.Index) & "|" & _
        CStr(PI_BandSlideBackground(targetSlide.SlideIndex))
End Function

Private Function PI_CanReuse(opts As ProgressDotsOptions) As Boolean
    PI_CanReuse = mPIRunActive And opts.ReuseNavigation = 1 And opts.FlatBar = 1 And _
        PI_UsesSectionBand(opts) And opts.PIStyle = "Full" And opts.ignoreSections = "Full" And _
        opts.ignoreIndicator = "All" And opts.DensityMode = 0
End Function

Private Sub PI_ClearLayoutCache()
    Dim index As Long
    For index = 1 To UBound(mPISectionSlots)
        If Len(mPISectionSlots(index).surfacePath) > 0 Then
            Kill mPISectionSlots(index).surfacePath
        End If
    Next index
    Erase mPISectionSlots
    Erase mPIMarkerSlots
    mPILayoutKey = ""
    mPIStoreLayout = False
    mPIReadLayout = False
End Sub

Private Function PI_DrawWithReuse(ByVal targetSlide As Slide, sections As ProgressDotsSectionInfo, _
                                  opts As ProgressDotsOptions, ByVal visibleIndex As Long) As Shape
    Dim key As String
    PI_RequireNativeTarget
    mPIStoreLayout = PI_CanReuse(opts)
    If mPIStoreLayout Then key = PI_ReuseKey(targetSlide, sections, opts)
    mPIReadLayout = (mPIStoreLayout And key = mPILayoutKey)
    If Not mPIReadLayout And Len(mPILayoutKey) > 0 Then PI_ClearLayoutCache
    mPIStoreLayout = (Len(key) > 0)
    If mPIStoreLayout And Not mPIReadLayout Then
        ReDim mPIMarkerSlots(1 To sections.lngSecIndicatorsRecalculated)
        mPILayoutKey = key
    End If
    DrawProgressBarOnSlide targetSlide.SlideIndex, sections, opts, "__progress-indicator__", CInt(visibleIndex)
    If mPILastDrawResult Is Nothing Then Err.Raise vbObjectError + 3024, "PI_DrawWithReuse", PI_Zh("672C98755BFC822A672A80FD5B8C6574751F62103002")
    If mPIReadLayout Then
        mPILayoutHits = mPILayoutHits + 1
    Else
        mPIComputedPages = mPIComputedPages + 1
    End If
    Set PI_DrawWithReuse = mPILastDrawResult
End Function

Private Sub PI_RemoveOtherNavigation(ByVal targetSlide As Slide, ByVal keep As Shape, ByVal shapeTag As String)
    Dim index As Long, candidate As Shape, keepId As Long
    PI_RequireNativeTarget
    keepId = keep.Id
    For index = targetSlide.Shapes.Count To 1 Step -1
        Set candidate = targetSlide.Shapes(index)
        If UCase$(candidate.Tags(shapeTag)) = "YES" Then
            If candidate.Id <> keepId Then candidate.Delete
        End If
    Next index
End Sub

Private Sub PI_AddHardShadow(ByVal band As Shape, opts As ProgressDotsOptions, ByVal shapeTag As String, _
                             ByVal presetGrowth As Double)
    Dim offset As Double, along As Double, across As Double
    Dim shadow As Shape, duplicates As ShapeRange
    Dim failureDescription As String
    On Error GoTo ShadowError
50     offset = PI_BarClamp(opts.HardShadowOffset, 0, 8)
60     If offset = 0 Or band.Fill.Visible <> msoTrue Then Exit Sub
70     If band.Fill.Transparency >= 1 Then Exit Sub
       ' Reserve the offset inside the allocated chapter, including the last one.
       along = PI_BarClamp(offset, 0, band.Width / 3)
       across = PI_BarClamp(offset, 0, (band.Height - presetGrowth) / 3)
       band.LockAspectRatio = msoFalse
       band.Width = band.Width - along
       If opts.PresetSize = 0 Or opts.DockEdge >= 2 Then band.Height = band.Height - across
80     Set duplicates = band.Duplicate
90     Set shadow = duplicates.Item(1)
100     shadow.Left = band.Left + along
110     shadow.Top = band.Top + across
120     shadow.Fill.Solid
130     shadow.Fill.ForeColor.RGB = RGB(0, 0, 0)
140     shadow.Fill.Transparency = 0
150     PI_ClearSurfaceEffects shadow
160     shadow.Tags.Add shapeTag, "yes"
170     shadow.Tags.Add "PI_ROLE", "section-hard-shadow"
180     shadow.ActionSettings(ppMouseClick).Action = ppActionHyperlink
        shadow.ActionSettings(ppMouseClick).Hyperlink.SubAddress = band.ActionSettings(ppMouseClick).Hyperlink.SubAddress
190     shadow.ActionSettings(ppMouseOver).Action = ppActionNone
200     Do While shadow.ZOrderPosition > band.ZOrderPosition
210         shadow.ZOrder msoSendBackward
    Loop
    Exit Sub
ShadowError:
    failureDescription = Err.Description
    PI_FirstError "PI_AddHardShadow", Erl, Err.Number, Err.Source, Err.Description
    On Error Resume Next
    If Not shadow Is Nothing Then shadow.Delete
    On Error GoTo 0
    Err.Raise vbObjectError + 3023, "PI_AddHardShadow", failureDescription
End Sub

Private Sub PI_ValidateDocumentCapacity(ByVal presentation As Presentation, opts As ProgressDotsOptions)
    Dim index As Long, count As Long
    If presentation.Slides.Count > 32767 Then
        Err.Raise vbObjectError + 3024, "PI_ValidateDocumentCapacity", PI_Zh("5F53524D7248672C6700591A652F63010033003200370036003798753002")
    End If
    If opts.ignoreSections = "None" Then Exit Sub
    For index = 1 To presentation.SectionProperties.Count
        If presentation.SectionProperties.FirstSlide(index) > 0 Then count = count + 1
    Next index
    If count > 100 Then
        Err.Raise vbObjectError + 3024, "PI_ValidateDocumentCapacity", PI_Zh("5F53524D7248672C6700591A652F63010031003000304E2A975E7A7A7AE0828230028BF754085E767AE08282540E518D751F62103002")
    End If
End Sub

Private Sub PI_PreflightNavigation(ByVal presentation As Presentation, sections As ProgressDotsSectionInfo, _
                                    opts As ProgressDotsOptions, ByVal firstSlide As Long, ByVal lastSlide As Long)
    Dim index As Long, targetSlide As Slide, resolved As ProgressDotsOptions
    Dim axis As Double, across As Double, navLeft As Double, navWidth As Double
    Dim railLeft As Double, railWidth As Double, titleTop As Double, dotsTop As Double, bandHeight As Double
    For index = firstSlide To lastSlide
        mPISlide = index
        Set targetSlide = presentation.Slides(index)
        mPISection = targetSlide.sectionIndex
        If PI_IsGenerationTarget(targetSlide) And PI_ShouldDisplaySlide(targetSlide, opts) Then
            resolved = PI_OptionsForSlide(opts, targetSlide)
            axis = PI_RenderAxisLength(resolved, presentation)
            across = presentation.PageSetup.SlideHeight
            If resolved.DockEdge >= 2 Then across = presentation.PageSetup.SlideWidth
            PI_NavigationGeometry resolved, axis, navLeft, navWidth, railLeft, railWidth
            If navWidth - (sections.lngSecSectionsRecalculated - 1) * resolved.SectionGap <= 0 Then
                Err.Raise vbObjectError + 3024, "PI_PreflightNavigation", _
                    PI_Zh("5BFC822A53EF7528957F5EA64E0D8DB3FF0C98759762FF1A") & CStr(index)
            End If
            If PI_UsesSectionBand(resolved) Then
                PI_BandGeometry resolved, titleTop, dotsTop, bandHeight
                If resolved.LayeredNavigation = 1 Then bandHeight = bandHeight + PI_BarClamp(resolved.RailThickness, 0.5, 30)
                If resolved.yMargin < 0 Or resolved.yMargin + bandHeight > across Then
                    Err.Raise vbObjectError + 3024, "PI_PreflightNavigation", _
                        PI_Zh("5BFC822A539A5EA68D8551FA98759762FF0C8BF78C0365748FB98DDD621665B957579AD85EA6002F4FA7680F5BBD5EA6300298759762FF1A") & CStr(index)
                End If
            End If
        End If
    Next index
End Sub

Private Sub PI_StartUndoBoundary()
    PI_RequireNativeTarget
    On Error Resume Next
    CallByName Application, "StartNewUndoEntry", VbMethod
    On Error GoTo 0
End Sub

Private Sub PI_FirstError(ByVal procedureName As String, ByVal lineNumber As Long, _
                          ByVal number As Long, ByVal source As String, ByVal description As String)
    If mPIFirstNumber <> 0 Then Exit Sub
    mPIFirstNumber = number
    mPIFirstSource = source
    mPIFirstDescription = description
    If Len(mPITraceFile) > 0 Then PI_Trace "error " & procedureName & "@" & CStr(lineNumber) & " #" & CStr(number)
    mPIFirstStage = procedureName & "@" & CStr(lineNumber) & " / " & mPIStage
    If Not PI_NativeTargetActive() Then Exit Sub
    If mPISlide > 0 And Not mPIRunPresentation Is Nothing Then mPIShapes = mPIRunPresentation.Slides(mPISlide).Shapes.Count
End Sub

Public Sub PI_RaiseNativeFailure(ByVal procedureName As String, ByVal lineNumber As Long, _
                                 ByVal number As Long, ByVal source As String, ByVal description As String)
    PI_FirstError procedureName, lineNumber, number, source, description
    Err.Raise number, source, description
End Sub

Private Sub PI_Trace(ByVal stage As String, Optional ByVal drawSection As Long = 0)
    Dim channel As Integer, shapes As Long, failure As Long, detail As String
    On Error GoTo TraceError
    If drawSection > 0 Then mPITraceSection = drawSection
    shapes = -1
    If mPISlide > 0 And Not mPIRunPresentation Is Nothing Then
        shapes = mPIRunPresentation.Slides(mPISlide).Shapes.Count
    End If
    channel = FreeFile
    Open mPITraceFile For Append As #channel
    Print #channel, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & Format$(Timer, "0.000") & vbTab & _
        stage & vbTab & CStr(mPISlide) & vbTab & CStr(mPISection) & vbTab & CStr(mPITraceSection) & _
        vbTab & CStr(shapes) & vbTab & CStr(mPICaptures) & vbTab & CStr(mPISurfaceHits)
    Close #channel
    Exit Sub
TraceError:
    failure = Err.Number
    detail = mPITraceFile & " / " & stage & ": " & Err.Description
    mPITraceFile = ""
    On Error Resume Next
    If channel <> 0 Then Close #channel
    On Error GoTo 0
    Err.Raise failure, "PI_Trace", detail
End Sub

Private Function PI_ResultField(ByVal value As String) As String
    PI_ResultField = Replace(Replace(Replace(value, vbTab, " "), vbCr, " "), vbLf, " ")
End Function

' The caller activates its own presentation. The style, range and instrumentation
' apply only to this invocation; the presentation's stored options are untouched.
Public Function PI_NativeGenerate(ByVal request As String) As String
    Dim fields() As String, previous As ProgressDotsOptions, opts As ProgressDotsOptions
    Dim savedStart As Integer, savedEnd As Integer, savedPreview As Integer
    Dim savedImages As Integer, savedCancel As Integer, saved As String
    Dim savedLast As Integer, savedPage As Integer, optionIndex As Long, pair() As String
    Dim shouldRestore As Boolean
    If gblbIndicatorUpdateRunning Or mPISilent Then Err.Raise 5, "PI_NativeGenerate", "Generation is already active."
    mPISilent = True
    mPIFirstNumber = 0
    mPIFirstSource = ""
    mPIFirstDescription = ""
    mPIFirstStage = ""
    mPIParameters = ""
    mPIStage = "request"
    mPISlide = 0
    mPISection = 0
    mPIShapes = 0
    mPIDrawn = 0
    mPIComputedPages = 0
    mPILayoutHits = 0
    mPICaptures = 0
    mPISurfaceHits = 0
    mPIResultStatus = "error"
    On Error GoTo NativeError
    fields = Split(request, "|")
    PI_BindNativeTarget fields
    If gblIndicatorSettingsVersion = "" Then Initialize
    previous = getDefaultOptions()
    savedStart = gbliDrawStart
    savedEnd = gbliDrawEnd
    savedPreview = gblShowPreview
    savedImages = gblshowCustomImage
    savedCancel = gblCancelCreate
    savedLast = gblLastPI
    savedPage = gblLastPIslideNum
    shouldRestore = True
    saved = ReadPersistensPISettings(mPINativePresentation)
    opts = previous
    If Len(saved) > 0 Then opts = DeserializeProgressOptions(saved)
    optionsToRibbon opts
    gblDockEdge = CInt(fields(1))
    If fields(0) <> "@document" Then PI_ApplyNamedPreset fields(0), opts
    gblEvenSpaceSecOrInd = fields(2)
    gbliDrawStart = CInt(fields(3))
    gbliDrawEnd = CInt(fields(4))
    gblReuseNavigation = CInt(fields(5))
    mPICancelAfter = CLng(fields(6))
    For optionIndex = 7 To UBound(fields)
        pair = Split(fields(optionIndex), "=", 2)
        Select Case pair(0)
            Case "targetDeck": ' Bound before the presentation settings were read.
            Case "display": gblstrIndicatorStyle = pair(1)
            Case "density": gblDensityMode = CInt(pair(1))
            Case "counter": gblCounterMode = CInt(pair(1))
            Case "height": PI_SetManualBandHeight Val(pair(1))
            Case "backdrop": gblSectionBandBackdrop = CInt(pair(1))
            Case "backdropAlpha": gblSectionBandBackdropTransparency = CDbl(pair(1))
            Case "backdropScale": gblSectionBandBackdropScale = CDbl(pair(1))
            Case "cancelFile": mPICancelFile = pair(1)
            Case "scratchDir": mPIScratchRoot = pair(1)
            Case "traceFile": mPITraceFile = pair(1)
            Case Else: PI_ApplyStyleValues fields(optionIndex)
        End Select
    Next optionIndex
    mPITraceSection = 0
    If Len(mPITraceFile) > 0 Then PI_Trace "run.begin " & Join(Array(fields(0), fields(1), fields(2), fields(3), fields(4), fields(5)), "|")
    gblShowPreview = 0
    opts = getDefaultOptions()
    MakeProgressBar opts
    If mPIResultStatus = "ok" Then PI_RequireNativeTarget
NativeComplete:
    If shouldRestore Then
        shouldRestore = False
        gbliDrawStart = savedStart
        gbliDrawEnd = savedEnd
        gblShowPreview = savedPreview
        gblshowCustomImage = savedImages
        gblCancelCreate = savedCancel
        gblLastPI = savedLast
        gblLastPIslideNum = savedPage
        optionsToRibbon previous
    End If
    If Len(mPITraceFile) > 0 Then PI_Trace "run.end " & mPIResultStatus
    mPITraceFile = ""
    Set mPINativePresentation = Nothing
    mPISilent = False
    mPICancelAfter = 0
    mPICancelFile = ""
    mPIScratchRoot = ""
    PI_NativeGenerate = Join(Array(mPIResultStatus, CStr(mPIFirstNumber), PI_ResultField(mPIFirstSource), _
        PI_ResultField(mPIFirstDescription), PI_ResultField(mPIFirstStage), CStr(mPISlide), _
        CStr(mPISection), CStr(mPIShapes), PI_ResultField(mPIParameters), CStr(mPIComputedPages), _
        CStr(mPILayoutHits), CStr(mPICaptures), CStr(mPIDrawn), CStr(mPISurfaceHits)), vbTab)
    Exit Function
NativeError:
    mPIResultStatus = "error"
    PI_FirstError "PI_NativeGenerate", Erl, Err.Number, Err.Source, Err.Description
    Resume NativeComplete
End Function

Sub MakeProgressBar(theOptions As ProgressDotsOptions)
    Dim sectinfo As ProgressDotsSectionInfo, slideOptions As ProgressDotsOptions
    Dim slideNum As Long, startSlide As Long, endSlide As Long, visibleIndex As Long, drawn As Long
    Dim renderDeck As Presentation, targetSlide As Slide, drawnGroup As Shape
    Dim preview As Boolean, visible As Boolean, repeated As Boolean, target As Boolean
    Dim started As Double, lastNotice As Double, elapsed As Double, estimated As Double
    Dim cancelled As Boolean, failureNumber As Long, failureDescription As String, ownsLock As Boolean
    Dim undoBoundaryStarted As Boolean
    On Error GoTo RenderError
100     If mPIRunActive Then Exit Sub
110     mPIFirstNumber = 0
120     mPIFirstStage = ""
130     mPISlide = 0
140     mPISection = 0
150     mPIStage = "preflight"
155     PI_RequireNativeTarget
160     If Application.Presentations.Count = 0 Then Exit Sub
165     Set renderDeck = PI_RenderPresentation()
170     If renderDeck.Slides.Count = 0 Then Exit Sub
180     PI_ValidateDocumentCapacity renderDeck, theOptions
190     If gbliDrawEnd < gbliDrawStart Then Err.Raise vbObjectError + 3024, "MakeProgressBar", PI_Zh("8D7759CB987578014E0D80FD59274E8E7ED3675F987578013002")

210     If Not PI_ValidateDisplayBinding(theOptions, renderDeck) Then
220         gblShowPreview = 0
        Exit Sub
    End If
250     If Not gblbIndicatorUpdateRunning Then
260         gblbIndicatorUpdateRunning = True
270         ownsLock = True
    End If
290     Set mPIRunPresentation = renderDeck
    If Len(mPITraceFile) > 0 Then PI_Trace "window.prepare.begin"
300     PI_PrepareRenderWindow
    If Len(mPITraceFile) > 0 Then PI_Trace "window.prepare.end"
310     preview = (gblShowPreview <> 0)
320     startSlide = gbliDrawStart
330     If startSlide < 1 Then startSlide = 1
340     endSlide = mPIRunPresentation.Slides.Count
350     If gbliDrawEnd < endSlide Then endSlide = gbliDrawEnd
360     PI_BeginRenderRun mPIRunPresentation
370     started = Timer
380     lastNotice = started
390     gblCancelCreate = 0
400     gblPIShapeCountTotalShape = 0
410     gblPIShapeCountTotalSection = 0
420     gblLastPIslideNum = 0
430     If Not mPISilent Then
440     frmPIisWorking.cmdCancelCreation.Caption = PI_Zh("505C6B62751F6210")
450     frmPIisWorking.lblInfo.Caption = PI_Zh("6B635728751F62108FDB5EA66761FF0C5904740698757801FF1A") & startSlide & " - " & endSlide
460     frmPIisWorking.lblStateXofY.Caption = PI_Zh("6B6357288BFB53D67AE082824E0E987597624FE1606F3002")
470     frmPIisWorking.Show
    End If
490     PI_YieldRender

510     RemoveSection "@progress_dots_remove_all_sections@"
520     If gblstrSlideIgnoreSections <> "None" Then UsePPTSectionsList
530     sectinfo = SectionsBase()
540     SectionsDetails sectinfo
550     If sectinfo.lngSecSectionsRecalculated = 0 Or sectinfo.lngSecIndicatorsRecalculated = 0 Then
560         Err.Raise vbObjectError + 3024, "MakeProgressBar", PI_Zh("6CA1670953EF8BA151658FDB5EA6768498759762FF0C8BF768C067E5969085CF9875548C8DF38FC798758BBE7F6E3002")
    End If
    If Len(mPITraceFile) > 0 Then PI_Trace "preflight.begin"
580     PI_PreflightNavigation mPIRunPresentation, sectinfo, theOptions, startSlide, endSlide
    If Len(mPITraceFile) > 0 Then PI_Trace "preflight.end"
590     PI_StartUndoBoundary
600     undoBoundaryStarted = True

620     For slideNum = 1 To endSlide
625         PI_RequireNativeTarget
630         mPISlide = slideNum
640         mPIStage = "slide policy"
650         If mPISilent Then
660             If mPICancelAfter > 0 And drawn >= mPICancelAfter Then gblCancelCreate = 1
670             If Len(mPICancelFile) > 0 Then
680                 If Len(Dir$(mPICancelFile)) > 0 Then gblCancelCreate = 1
            End If
        End If
710         If gblCancelCreate = 1 Then
720             cancelled = True
            Exit For
        End If
750         Set targetSlide = mPIRunPresentation.Slides(slideNum)
            If Len(mPITraceFile) > 0 Then
                mPISection = targetSlide.sectionIndex
                mPITraceSection = 0
                PI_Trace "page.begin"
            End If
760         visible = (CheckActiveSlideVisible(sectinfo.strSecTotalInvisibleComplete, CInt(slideNum)) = 0)
770         repeated = CheckSlideRepeat(CInt(slideNum))
780         If visible Then
790             visibleIndex = visibleIndex + 1
800             repeated = False
810             gblLastPI = 0
820             gblLastPIslideNum = slideNum
        ElseIf repeated Then
840             gblLastPI = 1
        End If
860         target = (slideNum >= startSlide)
870         If target Then target = PI_IsGenerationTarget(targetSlide)
880         If target Then
890             If PI_ShouldDisplaySlide(targetSlide, theOptions) And (visible Or (repeated And visibleIndex > 0)) Then
900                 slideOptions = PI_OptionsForSlide(theOptions, targetSlide)
910                 mPISection = targetSlide.sectionIndex
920                 mPIStage = "generate / " & CStr(visibleIndex)
930                 mPIParameters = "dock=" & slideOptions.DockEdge & " mode=" & slideOptions.EvenSpaceSecOrInd & _
                    " blur=" & slideOptions.LayerShadowBlur & " shadowAlpha=" & slideOptions.LayerShadowTransparency & _
                    " distance=" & slideOptions.LayerShadowDistance & " hardOffset=" & slideOptions.HardShadowOffset
    If Len(mPITraceFile) > 0 Then PI_Trace "draw.begin"
960                 Set drawnGroup = PI_DrawWithReuse(targetSlide, sectinfo, slideOptions, visibleIndex)
    If Len(mPITraceFile) > 0 Then PI_Trace "replace.begin"
970                 PI_RemoveOtherNavigation targetSlide, drawnGroup, "__progress-indicator__"
980                 drawn = drawn + 1
990                 If preview And drawn >= 3 Then Exit For
            ElseIf Not preview Then
1010                 DelProgressBarOnSlide slideNum, "__progress-indicator__"
            End If
        End If

1050         elapsed = PI_RenderElapsed(started)
1060         If PI_RenderElapsed(lastNotice) >= 0.15 Or slideNum = endSlide Then
1070             estimated = 0
1080             If slideNum > 0 Then estimated = elapsed * endSlide / slideNum
1090             If Not mPISilent Then
1100             WaitProgressbarDraw CInt(100 * slideNum / endSlide)
1110             frmPIisWorking.lblStateXofY.Caption = PI_Zh("5DF268C067E598759762FF1A") & slideNum & " / " & endSlide & vbCrLf & _
                PI_Zh("5DF2751F62105BFC822AFF1A") & drawn & vbCrLf & _
                PI_Zh("801765F6FF1A") & Format$(elapsed, "0.0") & " s" & vbCrLf & _
                PI_Zh("98848BA1603B801765F6FF1A") & Format$(estimated, "0.0") & " s"
            End If
1160             lastNotice = Timer
1170             PI_YieldRender
        End If
        PI_RequireNativeTarget
    If Len(mPITraceFile) > 0 Then PI_Trace "page.end"
    Next slideNum

RenderComplete:
    On Error Resume Next
1230     gblShowPreview = 0
    If Len(mPITraceFile) > 0 Then PI_Trace "cache.release.begin"
1240     PI_EndRenderRun cancelled, failureDescription
    If Len(mPITraceFile) > 0 Then PI_Trace "cache.release.end"
1250     If Not mPISilent Then frmPIisWorking.Hide
    If Len(mPITraceFile) > 0 Then PI_Trace "window.restore.begin"
1260     PI_RestoreRenderWindow
    If Len(mPITraceFile) > 0 Then PI_Trace "window.restore.end"
1270     If undoBoundaryStarted And PI_NativeTargetActive() Then PI_StartUndoBoundary
1280     If ownsLock Then gblbIndicatorUpdateRunning = False
    On Error GoTo 0
1300     mPIDrawn = drawn
1310     If failureNumber <> 0 Then
1320         mPIResultStatus = "error"
    ElseIf cancelled Then
1340         mPIResultStatus = "cancelled"
    Else
1360         mPIResultStatus = "ok"
    End If
1380     If mPISilent Then Exit Sub
1390     If failureNumber <> 0 Then
1400         MsgBox CStr(mPIFirstNumber) & " " & mPIFirstSource & vbCrLf & mPIFirstDescription & vbCrLf & _
                mPIFirstStage & vbCrLf & PI_Zh("98759762FF1A") & mPISlide & "  " & PI_Zh("7AE08282FF1A") & mPISection & _
                "  " & PI_Zh("5BF98C616570FF1A") & mPIShapes & vbCrLf & mPIParameters, vbCritical, PI_Zh("8FDB5EA66761751F6210")
    ElseIf cancelled Then
1420         MsgBox PI_Zh("5DF2505C6B62751F621030025DF25B8C6210987597624FDD755965B05BFC822AFF0C51764F59987597624FDD6301539F72B63002"), vbInformation, PI_Zh("8FDB5EA66761751F6210")
    End If
    Exit Sub
RenderError:
    failureNumber = Err.Number
    failureDescription = Err.Description
    PI_FirstError "MakeProgressBar", Erl, Err.Number, Err.Source, Err.Description
    Resume RenderComplete
End Sub
' ----------------------------------------------------------------
' Procedure Name: WaitProgressbarDraw
' Purpose: Place a nice progress bar on the waiting information window
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter intTarget (Integer): percent value to which to fill the progress bar
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2020-03-07
' ----------------------------------------------------------------
Private Sub WaitProgressbarDraw(intTarget As Integer)
    Dim completed As Long
    completed = CLng(PI_BarClamp(intTarget, 0, 100))
    frmPIisWorking.lblProgressbar.Caption = String$(completed, "|") & String$(100 - completed, ".")
    frmPIisWorking.lblProgressbar.ControlTipText = CStr(completed) & PI_Zh("002500205DF25B8C6210")
End Sub
'Private Sub WaitProgressbarPrepare()
'Dim strProgressBase As String
'strProgressBase = ""
'    For intProgressBarCount = 1 To 10
'        strProgressBase = strProgressBase & String(9, ".") & Right(CStr(intProgressBarCount), 1)
'    Next intProgressBarCount
'    frmPIisWorking.lblProgressbar.Caption = strProgressBase
'End Sub
' ----------------------------------------------------------------
' Procedure Name: procFormatDirect
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter intPosition (Integer):
' Parameter lngColor (Long):
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Public Sub procFormatDirect(intPosition As Integer, lngColor As Long)
10        On Error GoTo ErrorHandler
          
20        If gblbolDebug = True Then LogInfo "Public Sub procFormatDirect: " & intPosition & " / " & lngColor          'Debug: Write active procedure to log
              
          Dim ctl As control
          '    Dim lngIndex As Long
          
          'Dim frm As control
          'Set frm = frmThemeColors
          
30        For Each ctl In frmThemeColors.Controls 'OptionsForm.Controls
40            If (ctl.Name Like "cmdColor*" And Len(ctl.Name) > 8) Then
50                If Right(ctl.Name, Len(Str(intPosition)) - 1) = intPosition Then
60                    If gblbolDebug = True Then LogInfo ":" & ctl.Name
70                    With ctl
80                        .BackColor = lngColor
90                        .ControlTipText = "RGB: " & Color_to_RGB(lngColor) & " (" & lngColor & ")"
100                   End With
110               End If
120           End If
130       Next ctl

exit_proc:
140       Exit Sub

ErrorHandler:
150       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
160       Select Case Err.Number
              Case -2147188160      'no active presentation
170               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "procFormatDirect: Nothing to do!"
180               Resume exit_proc
190               Resume
200           Case 9999
                  'spezifische Fehler behandeln
210           Case Else
220               MsgBox PI_Zh("95198BEF7F1653F7FF1A") & Err.Number & ". " & Err.description, , "procFormatDirect"
230               Resume exit_proc
240               Resume
250       End Select
              
End Sub

' ----------------------------------------------------------------
' Procedure Name: RefreshProgressBar
' Purpose: Refreshes the indicator bar
'       Replaces each target page after its new navigation is ready.
'       Immediately run after relevant changes to the ribbon/settings
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub RefreshProgressBar()
    Dim saved As String, theOptions As ProgressDotsOptions, ownsLock As Boolean
    If gblbIndicatorUpdateRunning Then Exit Sub
    On Error GoTo RefreshError
    Auto_Open
    If gblintGeneralProgressbarDirect <> 1 And Not gblCreatePressed Then Exit Sub
    gblCreatePressed = False
    If Application.Presentations.Count = 0 Then Exit Sub
    If ActivePresentation.Slides.Count = 0 Then Exit Sub
    saved = ReadPersistensPISettings()
    If Len(saved) = 0 Then saved = SerializeProgressOptions(getDefaultOptions())
    theOptions = DeserializeProgressOptions(saved)
    gblbIndicatorUpdateRunning = True
    ownsLock = True
    MakeProgressBar theOptions
RefreshComplete:
    If ownsLock Then gblbIndicatorUpdateRunning = False
    Exit Sub
RefreshError:
    MsgBox CStr(Err.Number) & ": " & Err.Description, vbCritical, "RefreshProgressBar"
    Resume RefreshComplete
End Sub

' ----------------------------------------------------------------
' Procedure Name: RemoveIgnoreSlide
' Purpose: removes the ignore (skipped) comment from a slide. Comment can be easily removed by hand by the user!
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub RemoveIgnoreSlide(intShowMsg As Integer)
    If gblbIndicatorUpdateRunning Then Exit Sub
          'idea and (some) code from https://github.com/utopiaplanetia/ppt-progress-bar
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Sub RemoveIgnoreSlide"         'Debug: Write active procedure to log
          
          Dim i As Integer
          Dim intRemoved As Integer
          Dim curslide As slide
          
30        intRemoved = 0
          
40        Set curslide = GetActiveSlide(ActiveWindow)
50        If curslide Is Nothing Then GoTo error_section  'if current slide can not be evaluated
          
          'Note to myself if the new method still comes up with error on MAC, maybe try
          ' http://stackoverflow.com/questions/30050130/ddg#32199430
         
60        With curslide
70            If .Comments.Count > 0 Then
80                For i = .Comments.Count To 1 Step -1    'backwards, so we can delete without messing up the counter
90                    If .Comments.Item(i).Text = gblstrMessageIgnore Then
100                       .Comments.Item(i).Delete
110                       intRemoved = intRemoved + 1
120                   End If

130                   If .Comments.Item(i).Text = gblstrMessageRepeat Then
140                       .Comments.Item(i).Delete
150                       intRemoved = intRemoved + 2
160                   End If

170               Next i
180           End If
190       End With
200       If intRemoved > 0 Then
              '    Beep
              '    MsgBox "Do not forget to recreate the progress bar to see the effect of not anymore ignored slides!", , "Slide not ignored anymore! Don't forget to refresh!"
210       Else
220           If intShowMsg = 1 Then
                  ' MsgBox "This slide was not skipped!", vbInformation, "Nothing to do..."
230           End If
240       End If

          
exit_proc:
250       Exit Sub
          
error_section:
260       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
270       MsgBox PI_Zh("65E06CD583B753D65F53524D9875780130028BF7900962E94E005F205E7B706F7247FF0C70B951FB7F168F91533A57DF540E91CD8BD53002"), , "RemoveIgnoreSlide: GetActiveSlide failed"
280       Resume exit_proc
290       Resume
          
ErrorHandler:
300       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
310       Select Case Err.Number
              Case -2147188160      'no active presentation
320               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "RemoveIgnoreSlide: Nothing to do!"
330               Resume exit_proc
340               Resume

350           Case -2147467259    'collection is empty = no comments on page ?!
360               Resume exit_proc
370               Resume

380           Case Else
390               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "RemoveIgnoreSlide"
400               MsgBox PI_Zh("82E50020004D0061006300204E0A63017EED51FA73B06B6495198BEFFF0C8BF74FDD7559524D4E00676195198BEF4FE1606F4EE54FBF639267E53002")
410               Resume Next
420               Resume
              
430       End Select
End Sub

' ----------------------------------------------------------------
' Procedure Name: RemoveSection
' Purpose: removes a single (or all) sections from the internal list (in prev. versions there werer sections beside ppt-sections. those have been removed for consistency)
' Procedure Kind: Function
' Procedure Access: Public
' Parameter strSectionName (String): sectionumber OR joker-string (which removes all sections)
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function RemoveSection(strSectionName As String)
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct RemoveSection: " & strSectionName         'Debug: Write active procedure to log
              
          
          Dim curTitle As String
          Dim slideNum As Long
          Dim strJoker As String
          Dim renderDeck As Presentation
          Set renderDeck = PI_RenderPresentation()
30        strJoker = "@progress_dots_remove_all_sections@"    'use this as option to remove all sections at once
          
          ' Find slide with sectionname
40        If gblbolDebug = True Then LogInfo "Scanning for section to remove..."

50        For slideNum = 1 To renderDeck.Slides.Count
60            curTitle = renderDeck.Slides(slideNum).Tags.Item(gblstrMessageTagTitle)
70            If (curTitle = strSectionName Or strSectionName = strJoker) Then
                  ' Delete the sectiontitle from slide tag.
                  '            curSlide.Tags.Delete gblstrMessageTagTitle
                  'Debug.Print renderDeck.Slides(slideNum).Tags.Value(slideNum)
75                PI_RequireNativeTarget
80                renderDeck.Slides(slideNum).Tags.Delete (gblstrMessageTagTitle)
90                RemoveSection = "success"
100           End If
110           If gblbolDebug = True Then LogInfo "Removed section: " & curTitle & " / " & strSectionName & " on slide " & slideNum
120       Next slideNum

exit_proc:
130       Exit Function
          
ErrorHandler:
          If mPISilent Then PI_RaiseNativeFailure "RemoveSection", Erl, Err.Number, Err.Source, Err.Description
140       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
150       Select Case Err.Number
              Case -2147188160      'no active presentation
160               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "RemoveSection: Nothing to do!"
170               Resume exit_proc
180               Resume
190           Case Else
200               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "RemoveSection"
210               RemoveSection = "error"
220               Resume exit_proc
230               Resume
              
240       End Select
          
End Function

' ----------------------------------------------------------------
' Procedure Name: SectionsBase
' Purpose: Get some base information about (existing) sections in ppt. Further details in sectiondetails()
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: ProgressDotsSectionInfo
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function SectionsBase() As ProgressDotsSectionInfo
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct SectionsBase"         'Debug: Write active procedure to log
          
          Dim bolNotThis As Boolean
          Dim slideNum, sectionNum As Integer
          Dim curTitle As String
          '    Dim i As Integer
          
          Dim renderDeck As Presentation
          Set renderDeck = PI_RenderPresentation()
30        bolNotThis = False
          
          ' Initialize the 0th group.
40        SectionsBase.strArySecTitles(0) = ""
50        SectionsBase.lngArySecStart(0) = 1
          
          ' Scan for section titles.
60        sectionNum = 0
70        If gblbolDebug = True Then LogInfo "Scanning for sections..."

80        For slideNum = 1 To renderDeck.Slides.Count
90            curTitle = renderDeck.Slides(slideNum).Tags.Item(gblstrMessageTagTitle)
100           If gblbolDebug = True Then LogInfo slideNum & " " & curTitle
              
110           If (curTitle <> "") Then
120               sectionNum = sectionNum + 1
130               SectionsBase.strArySecTitles(sectionNum) = curTitle
140               SectionsBase.lngArySecStart(sectionNum) = slideNum
150               SectionsBase.intArySecIgnoreSkipped(sectionNum) = 0
160               SectionsBase.intArySecIgnoreHidden(sectionNum) = 0
170               If gblbolDebug = True Then LogInfo "new section " & sectionNum & " : " & slideNum & " / " & curTitle
180           End If
              
190           If gblintSlideIgnoreSkipped = 1 Then    'slides marked by user to be ignored
200               If CheckSlideSkipped(CInt(slideNum)) = True Or CheckSlideRepeat(CInt(slideNum)) = True Then
210                   SectionsBase.intArySecIgnoreSkipped(sectionNum) = SectionsBase.intArySecIgnoreSkipped(sectionNum) + 1
220                   bolNotThis = True
230               End If

240           End If
250           If gblbolDebug = True Then LogInfo "S: " & SectionsBase.intArySecIgnoreSkipped(sectionNum)   'CheckSlideSkipped(CInt(slideNum)),


260           If gblintSlideIgnoreHidden = 1 Then 'ppt hidden slides
270               If CheckSlideHidden(CInt(slideNum)) = True Then
280                   SectionsBase.intArySecIgnoreHidden(sectionNum) = SectionsBase.intArySecIgnoreHidden(sectionNum) + 1
290                   bolNotThis = True
300               End If
310           End If
320           If gblbolDebug = True Then LogInfo "H: " & SectionsBase.intArySecIgnoreHidden(sectionNum) 'CheckSlideHidden(CInt(slideNum))
              
330           If SectionsBase.strArySecTitles(sectionNum) = gblstrPISectionHidden Then
                  'SectionsBase.intArySecIgnoreSkipped(sectionNum) = SectionsBase.intArySecIgnoreSkipped(sectionNum) + 1
                  'SectionsBase.intArySecIgnoreHidden(sectionNum) = SectionsBase.intArySecIgnoreHidden(sectionNum) + 1
340               bolNotThis = True
350           End If
              
360           If bolNotThis = True Then
370               SectionsBase.strArySecTotalInvisible(sectionNum) = SectionsBase.strArySecTotalInvisible(sectionNum) & ";" & slideNum
380               SectionsBase.strSecTotalInvisibleComplete = SectionsBase.strSecTotalInvisibleComplete & ";" & slideNum
390               bolNotThis = False
400           Else    'slide is VISIBLE and GETS indicator
410               SectionsBase.strSecTotalVisibleComplete = SectionsBase.strSecTotalVisibleComplete & ";" & slideNum
420               If SectionsBase.intAryFirstVisiblePageInSection(sectionNum) = 0 Then 'set only the first page
430                   SectionsBase.intAryFirstVisiblePageInSection(sectionNum) = slideNum
440               End If
450               If SectionsBase.lngSecStartSlide = 0 Then SectionsBase.lngSecStartSlide = slideNum  'set the start for the first section(name)
460           End If
              
470           SectionsBase.intArySecTotal(sectionNum) = SectionsBase.intArySecTotal(sectionNum) + 1

480           If SectionsBase.intArySecTotal(sectionNum) - CharCount(SectionsBase.strArySecTotalInvisible(sectionNum), ";") > 0 Then
490               SectionsBase.intArySecVisible(sectionNum) = 1
500           Else
510               SectionsBase.intArySecVisible(sectionNum) = 0
520           End If
              
530           If gblbolDebug = True Then
540               LogInfo "Section " & sectionNum & " is Visible: " & SectionsBase.intArySecVisible(sectionNum)
550               LogInfo "1st visible page in sec " & sectionNum & " : " & SectionsBase.intAryFirstVisiblePageInSection(sectionNum)
560               LogInfo "total_pages_in_sec " & sectionNum & " : " & SectionsBase.intArySecTotal(sectionNum)
570               LogInfo "Invisible Pages: " & SectionsBase.strArySecTotalInvisible(sectionNum)
580           End If
              
590       Next slideNum
       
          ' If only one section (or no section) update the values
600       If sectionNum = 0 And renderDeck.Slides.Count >= 1 Then
610           If gblbolDebug = True Then LogInfo "only one or no section was found!"
620           sectionNum = 1
630           SectionsBase.strArySecTitles(1) = SectionsBase.strArySecTitles(0)       'Title is empty if no user-defined sections
              
640           SectionsBase.lngArySecStart(1) = SectionsBase.intAryFirstVisiblePageInSection(0)  '2 to start on this slide if no user-defined sections
650           SectionsBase.lngSecStartSlide = SectionsBase.intAryFirstVisiblePageInSection(0)     '2 to start on this slide if no user-defined sections
660           SectionsBase.intAryFirstVisiblePageInSection(1) = SectionsBase.intAryFirstVisiblePageInSection(0)
              
670           SectionsBase.strArySecTotalInvisible(1) = SectionsBase.strArySecTotalInvisible(0)
680           SectionsBase.intArySecTotal(1) = SectionsBase.intArySecTotal(0)
              
690           SectionsBase.intArySecIgnoreHidden(1) = SectionsBase.intArySecIgnoreHidden(0)
700           SectionsBase.intArySecIgnoreSkipped(1) = SectionsBase.intArySecIgnoreSkipped(0)
              
710           SectionsBase.intArySecVisible(1) = SectionsBase.intArySecVisible(0)
720       End If
          
          ' Record the end slide as the "next" section start -- this makes calculations easier.
730       SectionsBase.lngArySecStart(sectionNum + 1) = renderDeck.Slides.Count + 1
              
          'add another ; for find contents more easy
740       SectionsBase.strSecTotalInvisibleComplete = SectionsBase.strSecTotalInvisibleComplete & ";"
              
          ' Return the information we found.
750       SectionsBase.lngSecSections = sectionNum
              
760       If gblbolDebug = True Then
770           LogInfo "Sections found: " & SectionsBase.lngSecSections
780           LogInfo "all Visible pages in all sections: " & SectionsBase.strSecTotalVisibleComplete
790           LogInfo "all invisible pages in all sections: " & SectionsBase.strSecTotalInvisibleComplete
800           LogInfo "---"
810       End If

exit_proc:
820       Exit Function
          
ErrorHandler:
          If mPISilent Then PI_RaiseNativeFailure "SectionsBase", Erl, Err.Number, Err.Source, Err.Description
830       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
840       Select Case Err.Number
              Case -2147188160      'no active presentation
850               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "SectionsBase: Nothing to do!"
860               Resume exit_proc
870               Resume
880           Case Else
890               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "SectionsBase"
900               Resume Next
910               Resume
              
920       End Select
         
End Function

' ----------------------------------------------------------------
' Procedure Name: SectionsDetails
' Purpose: Check for some details of a section
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter sectinfo (ProgressDotsSectionInfo): section information aquired in a previous step
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub SectionsDetails(sectinfo As ProgressDotsSectionInfo)
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Sub SectionsDetails (sectinfo As ProgressDotsSectionInfo)"         'Debug: Write active procedure to log
              
          Dim bolIndicatorOnThisSlide As Boolean
          Dim totalSectionsNew, totalSectionsOld As Integer
          Dim totalSlides, totalSlidesHidden As Integer
          Dim totalIndicatorsOld, totalIndicatorsNew  As Integer
          Dim intTestvalue As Integer
          Dim sectNum As Integer

          Dim renderDeck As Presentation
          Set renderDeck = PI_RenderPresentation()
30        bolIndicatorOnThisSlide = False
          
40        totalSlides = renderDeck.Slides.Count
50        totalSlidesHidden = CharCount(sectinfo.strSecTotalInvisibleComplete, ";") - 1 'because one ending ; as well
          
60        totalSectionsNew = 0
70        totalSectionsOld = sectinfo.lngSecSections
          
80        totalIndicatorsNew = 0
90        totalIndicatorsOld = totalSlides
          
          'START Section check--------------
          'check if there are sections which contain only not-marked slides. If so: no header for this section either!
100       totalSectionsNew = 0
110       For sectNum = 1 To totalSectionsOld
          
120           If gblbolDebug = True Then LogInfo "total: " & sectinfo.intArySecTotal(sectNum) & " of those are hidden: " & CharCount(sectinfo.strArySecTotalInvisible(sectNum), ";")
130           intTestvalue = sectinfo.intArySecTotal(sectNum) - CharCount(sectinfo.strArySecTotalInvisible(sectNum), ";")
140           If intTestvalue > 0 Then   'more slides in section that hidden/skipped/invisible
150               sectinfo.intArySecVisible(sectNum) = 1
160               totalSectionsNew = totalSectionsNew + 1
170               totalIndicatorsNew = totalIndicatorsNew + intTestvalue
180               If gblbolDebug = True Then LogInfo "ok " & sectinfo.strArySecTitles(sectNum)
190           Else
200               sectinfo.intArySecVisible(sectNum) = 0
210               If gblbolDebug = True Then LogInfo "not ok " & sectinfo.strArySecTitles(sectNum)
220           End If
230           If gblbolDebug = True Then LogInfo "sectVisi: " & sectinfo.intArySecVisible(sectNum)
              
240       Next sectNum
          
250       sectinfo.lngSecSectionsRecalculated = totalSectionsNew
260       sectinfo.lngSecIndicatorsRecalculated = totalIndicatorsNew
          
270       If gblbolDebug = True Then LogInfo "totalSections Old / New: " & totalSectionsOld & ", " & totalSectionsNew & " if new=0 no indicator at all needed!"
280       If gblbolDebug = True Then LogInfo "totalIndicators Old / New: " & totalIndicatorsOld & "," & totalIndicatorsNew & " if new=0 no indicator at all needed!"
          'END Section check--------------
290       If gblbolDebug = True Then LogInfo "total-totalinvisible: " & totalSlides & " - " & totalSlidesHidden & " = " & totalSlides - totalSlidesHidden

exit_proc:
300       Exit Sub
          
ErrorHandler:
          If mPISilent Then PI_RaiseNativeFailure "SectionsDetails", Erl, Err.Number, Err.Source, Err.Description
310       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
320       Select Case Err.Number
              Case -2147188160      'no active presentation
330               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "MakeProgressBar: Nothing to do!"
340               Resume exit_proc
350               Resume
360           Case Else
370               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "MakeProgressBar"
380               Resume Next
390               Resume
              
400       End Select


End Sub

' ----------------------------------------------------------------
' Procedure Name: SetSection
' Purpose: Create a new section. Calls from UsePPTSectionsList() to create a section which will be used then. This would also allow adding user-defined sections!
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter strTitle (String): section title
' Parameter nrSlide (Long): slidenumber where the section starts
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub SetSection(Optional strTitle As String, Optional nrSlide As Long)
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Sub SetSection: " & strTitle & " / " & nrSlide       'Debug: Write active procedure to log
              
          Dim curslide As slide
          Dim intWorkNow As Integer
          
          'if slidenumber is given, go to this slide - to use old code as much as possible
30        If nrSlide <> -1 And nrSlide > 0 Then                           '-1 means section contains no slides (if so, we should not get here)
40            Application.ActiveWindow.View.GotoSlide nrSlide
50        Else
60            GoTo error_unknown
70        End If
          
80        intWorkNow = 0
90        Set curslide = GetActiveSlide(ActiveWindow)
100       If curslide Is Nothing Then GoTo error_section  'if current slide can not be evaluated
          
          'not needed anymore - this enabled the user to input his own (ppt independent) sections
          '    'IF not needed anymore: This allows to create internal (non-ppt) sections
          '    If strTitle = "" Then
          '        strTitle = InputBox("Sections can be defined  manually by selecting a slide and entering a section name OR use the built-in powerpoint sections.  To use powerpoint sections, see the 'sections' tab of the create box." & vbCrLf & vbCrLf & "Enter a title for the section " & _
          '            "starting on slide " & curslide.SlideNumber & "." & vbCrLf & vbCrLf & "Use SPACE to create a section without a name.", _
          '            "Progress Indicator: Section Title", _
          '            curslide.Tags.Item(gblstrMessageTagTitle))
          '    End If
          '
          '    'old: title is NOW automatically taken from ppt sections - so intWorkNow will always be 1
          '    If strTitle = "" And curslide.Tags.Item(gblstrMessageTagTitle) <> "" Then
          '        If MsgBox("You did not enter a name, removed the name or pressed the abort button." & vbCrLf & "Press yes if you want to remove the section from this page.", vbYesNo Or 0 Or vbSystemModal, "Clear existing section?") = vbYes Then
          '            intWorkNow = 1
          '        End If
          '    Else
          '        intWorkNow = 1
          '    End If
          'but now we need the next line
110       intWorkNow = 1
          
120       If intWorkNow = 1 Then
              ' Store the title as a tag.
130           curslide.Tags.Add gblstrMessageTagTitle, strTitle
140       End If
          
          'the following line does NOT WORK on MAC!
          'Call ActiveWindow.View.slide.Tags.Add(gblstrMessageTagTitle, strTitle)
          
          
exit_proc:
150       Exit Sub
          
error_section:
160       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
170       MsgBox PI_Zh("65E06CD583B753D65F53524D9875780130028BF7900962E94E005F205E7B706F7247FF0C70B951FB7F168F91533A57DF540E91CD8BD53002"), , "SetSection: GetActiveSlide failed"
180       Resume exit_proc
190       Resume

error_unknown:
200       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
210       MsgBox PI_Zh("98757801672A63D04F9B621665E0654830028BF7900962E94E005F205E7B706F7247FF0C70B951FB7F168F91533A57DF540E91CD8BD53002"), , "SetSection: unknown slidenumber"
220       Resume exit_proc
230       Resume
          
ErrorHandler:
          If mPISilent Then PI_RaiseNativeFailure "SetSection", Erl, Err.Number, Err.Source, Err.Description
240       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
250       Select Case Err.Number
              Case -2147188160      'no active presentation
260               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "SetSection: Nothing to do!"
270               Resume exit_proc
280               Resume
290           Case Else
300               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "SetSection"
310               Resume Next
320               Resume
              
330       End Select
          
End Sub

' ----------------------------------------------------------------
' Procedure Name: SpltText
' Purpose: Splits the string into single values
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter strTextToSplt (String): string to be seperated
' Parameter strSeperator (String): seperator character
' Parameter intReturnWish (Integer, Optional): value of the array to return
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function SpltText(strTextToSplt As String, strSeperator As String, Optional intReturnWish As Integer)
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct SpltText:" & strTextToSplt & " / " & strSeperator         'Debug: Write active procedure to log
              

          Dim varArray() As String
          Dim i As Integer
          
30        varArray = Split(strTextToSplt, strSeperator)

40        If gblbolDebug = True Then
50            For i = LBound(varArray) To UBound(varArray)
60                If varArray(i) <> "" Then
70                    LogInfo varArray(i)
80                End If
90            Next i
100       End If
          
110       If intReturnWish <> 0 Then
120           SpltText = varArray(intReturnWish)
130       Else
140           SpltText = intReturnWish
150       End If
          
exit_proc:
160       Exit Function
          
ErrorHandler:
170       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
180       Select Case Err.Number
              Case -2147188160      'no active presentation
190               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "SpltText: Nothing to do!"
200               Resume exit_proc
210               Resume
220           Case Else
230               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "SpltText"
240               Resume Next
250               Resume
              
260       End Select
          
End Function

' ----------------------------------------------------------------
' Procedure Name: UsePPTSectionsList
' Purpose: get a list of sections defined in ppt
' Procedure Kind: Function
' Procedure Access: Public
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function UsePPTSectionsList()
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct UsePPTSectionsList"         'Debug: Write active procedure to log
              
          Dim i As Long
        
          Dim renderDeck As Presentation
          Set renderDeck = PI_RenderPresentation()
30        If gblbolDebug = True Then LogInfo "Scanning for PPT sections to use ..."
40        With renderDeck.SectionProperties
50            For i = 1 To .Count
60                If gblbolDebug = True Then LogInfo "Name: " & .Name(i) & "; firstslide: " & .FirstSlide(i) & "; slidecount: " & .SlidesCount(i)
70                If .FirstSlide(i) <> -1 Then             '.firstslide(i) is -1 if a section had no slides
                      'SetSection .Name(i), .FirstSlide(i)
                      'much easier:
75                    PI_RequireNativeTarget
80                    renderDeck.Slides(.FirstSlide(i)).Tags.Add gblstrMessageTagTitle, .Name(i)
90                End If
100           Next i
110       End With
        
          
exit_proc:

120       Exit Function
          
ErrorHandler:
          If mPISilent Then PI_RaiseNativeFailure "UsePPTSectionsList", Erl, Err.Number, Err.Source, Err.Description
130       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
140       Select Case Err.Number
              Case -2147188160      'no active presentation
150               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "UsePPTSectionsList: Nothing to do!"
160               Resume exit_proc
170               Resume
180           Case Else
190               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "UsePPTSectionsList"
200               Resume exit_proc
210               Resume
              
220       End Select

End Function





' Section backgrounds span the title and indicator rows. Text and links stay
' separate from the background, so indicators remain on top and clickable.

Public Sub PI_BandGetPressed(control As IRibbonControl, ByRef pressed)
    If control.id = "cbxSectionBand" Then
        pressed = (gblSectionBand = 1)
    ElseIf control.id = "cbxSectionBandContrast" Then
        pressed = (gblSectionBandContrast = 1)
    ElseIf control.id = "cbxSectionBandGradient" Then
        pressed = (gblSectionBandGradient = 1)
    ElseIf control.id = "cbxBandBackdrop" Then
        pressed = (gblSectionBandBackdrop = 1)
    ElseIf control.id = "cbxFlatBar" Then
        pressed = (gblFlatBar = 1)
    ElseIf control.id = "cbxFlatBarGlow" Then
        pressed = (gblFlatBarGlow = 1)
    ElseIf control.id = "cbxPanelShadow" Then
        pressed = (gblPanelShadow = 1)
    ElseIf control.id = "cbxLayeredNavigation" Then
        pressed = (gblLayeredNavigation = 1)
    ElseIf control.id = "cbxSectionNumbering" Then
        pressed = (gblSectionNumbering = 1)
    ElseIf control.id = "cbxCurrentTitleBold" Then
        pressed = (gblCurrentTitleBold = 1)
    ElseIf control.id = "cbxFutureBarOutline" Then
        pressed = (gblFutureBarOutline = 1)
    ElseIf control.id = "cbxSectionForegroundBars" Then
        pressed = (gblSectionForegroundBars = 1)
    End If
End Sub

Public Sub PI_BandOnToggle(control As IRibbonControl, pressed As Boolean)
    If gblbIndicatorUpdateRunning Then Exit Sub
    If Application.Presentations.Count = 0 Then Exit Sub
    If control.id = "cbxSectionBand" Then
        gblSectionBand = Abs(CInt(pressed))
    ElseIf control.id = "cbxSectionBandContrast" Then
        gblSectionBandContrast = Abs(CInt(pressed))
    ElseIf control.id = "cbxSectionBandGradient" Then
        gblSectionBandGradient = Abs(CInt(pressed))
    ElseIf control.id = "cbxBandBackdrop" Then
        gblSectionBandBackdrop = Abs(CInt(pressed))
    ElseIf control.id = "cbxFlatBar" Then
        gblFlatBar = Abs(CInt(pressed))
    ElseIf control.id = "cbxFlatBarGlow" Then
        gblFlatBarGlow = Abs(CInt(pressed))
    ElseIf control.id = "cbxPanelShadow" Then
        gblPanelShadow = Abs(CInt(pressed))
    ElseIf control.id = "cbxLayeredNavigation" Then
        gblLayeredNavigation = Abs(CInt(pressed))
    ElseIf control.id = "cbxSectionNumbering" Then
        gblSectionNumbering = Abs(CInt(pressed))
    ElseIf control.id = "cbxCurrentTitleBold" Then
        gblCurrentTitleBold = Abs(CInt(pressed))
    ElseIf control.id = "cbxFutureBarOutline" Then
        gblFutureBarOutline = Abs(CInt(pressed))
    ElseIf control.id = "cbxSectionForegroundBars" Then
        gblSectionForegroundBars = Abs(CInt(pressed))
    End If
    PersistPISettings SerializeProgressOptionsRibbon()
    RefreshRibbon
    If Not gobjRibbon Is Nothing Then gobjRibbon.Invalidate
    RefreshProgressBar
End Sub

Public Sub PI_BandGetPadding(control As IRibbonControl, ByRef value)
    value = CStr(gblSectionBandPadding)
End Sub

Public Sub PI_BandSetPadding(control As IRibbonControl, text As String)
    If gblbIndicatorUpdateRunning Then Exit Sub
    Dim value As Double
    If Application.Presentations.Count = 0 Then Exit Sub
    If Not IsNumeric(text) Then
        MsgBox PI_Zh("8BF78F93516500200030002081F30020003200300020768465705B57FF0C53554F4D4E3A78C5FF0800700074FF09"), vbExclamation
        Exit Sub
    End If
    value = CDbl(text)
    If value < 0 Or value > 20 Then
        MsgBox PI_Zh("8272575751858FB98DDD7684830356F4662F00200030002081F3002000320030002078C5FF0800700074FF09"), vbExclamation
        Exit Sub
    End If
    gblSectionBandPadding = value
    PersistPISettings SerializeProgressOptionsRibbon()
    RefreshRibbon
    If Not gobjRibbon Is Nothing Then gobjRibbon.Invalidate
    RefreshProgressBar
End Sub

Public Sub PI_BandGetTransparency(control As IRibbonControl, ByRef value)
    Select Case control.id
        Case "edtBandTransparencyCurrent"
            value = CStr(gblSectionBandTransparencyCurrent)
        Case "edtBandTransparencyLeft"
            value = CStr(gblSectionBandTransparencyLeft)
        Case "edtBandTransparencyRight"
            value = CStr(gblSectionBandTransparencyRight)
        Case "edtBandBackdropTransparency"
            value = CStr(gblSectionBandBackdropTransparency)
    End Select
End Sub

Public Sub PI_BandChooseColor(control As IRibbonControl)
    If gblbIndicatorUpdateRunning Then Exit Sub
    Dim value As String, key As String, title As String, channels() As String
    If Application.Presentations.Count = 0 Then Exit Sub
    Select Case control.id
        Case "btnBandLeftNear"
            value = gblSectionBandLeftNear
            key = "Band Left Near"
            title = PI_Zh("5DE64FA776F890BB8272")
        Case "btnBandLeftFar"
            value = gblSectionBandLeftFar
            key = "Band Left Far"
            title = PI_Zh("5DE64FA767008FDC8272")
        Case "btnBandRightNear"
            value = gblSectionBandRightNear
            key = "Band Right Near"
            title = PI_Zh("53F34FA776F890BB8272")
        Case "btnBandRightFar"
            value = gblSectionBandRightFar
            key = "Band Right Far"
            title = PI_Zh("53F34FA767008FDC8272")
        Case "btnBandBackdropColor"
            value = gblSectionBandBackdropColor
            key = "Band Backdrop"
            title = PI_Zh("80CC666F5B57989C8272")
        Case "btnAccentColor"
            value = gblAccentColor
            key = "Section Accent"
            title = PI_Zh("5F53524D7AE068078BB0989C8272")
        Case Else
            Exit Sub
    End Select
    channels = Split(value, ";")
    With frmThemeColors
        .lblColorSelected.Tag = "New color for " & key
        .lblColorSelected.Caption = PI_Zh("65B0989C8272")
        .Caption = title
        CalcColor
        .lblColorOld.BackColor = RGB(Val(channels(0)), Val(channels(1)), Val(channels(2)))
        .lblColorSelected.BackColor = .lblColorOld.BackColor
        .dotSeenRt = Val(channels(0))
        .dotSeenGt = Val(channels(1))
        .dotSeenBt = Val(channels(2))
        .Show
    End With
    RefreshRibbon
    If Not gobjRibbon Is Nothing Then gobjRibbon.Invalidate
    RefreshProgressBar
End Sub

Public Sub PI_BandSetTransparency(control As IRibbonControl, text As String)
    If gblbIndicatorUpdateRunning Then Exit Sub
    Dim value As Double
    If Application.Presentations.Count = 0 Then Exit Sub
    If Not IsNumeric(text) Then
        MsgBox PI_Zh("8BF78F93516500200030002081F3002000310030003000204E4B95F4768465705B57FF0C4E0D542B767E520653F73002"), vbExclamation
        Exit Sub
    End If
    value = CDbl(text)
    If value < 0 Or value > 100 Then
        MsgBox PI_Zh("900F660E5EA6830356F44E3A00200030002081F300200031003000303002"), vbExclamation
        Exit Sub
    End If
    Select Case control.id
        Case "edtBandTransparencyCurrent"
            gblSectionBandTransparencyCurrent = value
        Case "edtBandTransparencyLeft"
            gblSectionBandTransparencyLeft = value
        Case "edtBandTransparencyRight"
            gblSectionBandTransparencyRight = value
        Case "edtBandBackdropTransparency"
            gblSectionBandBackdropTransparency = value
        Case Else
            Exit Sub
    End Select
    PersistPISettings SerializeProgressOptionsRibbon()
    RefreshRibbon
    If Not gobjRibbon Is Nothing Then gobjRibbon.Invalidate
    RefreshProgressBar
End Sub

Public Sub PI_BarGetValue(control As IRibbonControl, ByRef value)
    Select Case control.id
        Case "edtFlatBarHeight"
            value = CStr(gblFlatBarHeight)
        Case "edtFlatBarGap"
            value = CStr(gblFlatBarGap)
        Case "edtFlatBarRoundness"
            value = CStr(gblFlatBarRoundness)
        Case "edtFlatBarGlowRadius"
            value = CStr(gblFlatBarGlowRadius)
        Case "edtFlatBarGlowTransparency"
            value = CStr(gblFlatBarGlowTransparency)
    End Select
End Sub

Public Sub PI_BarSetValue(control As IRibbonControl, text As String)
    If gblbIndicatorUpdateRunning Then Exit Sub
    Dim value As Double, minimum As Double, maximum As Double
    If Application.Presentations.Count = 0 Then Exit Sub
    minimum = 0
    Select Case control.id
        Case "edtFlatBarHeight"
            minimum = 0.5
            maximum = 12
        Case "edtFlatBarGap", "edtFlatBarGlowRadius"
            maximum = 20
        Case "edtFlatBarRoundness"
            maximum = 50
        Case "edtFlatBarGlowTransparency"
            maximum = 100
        Case Else
            Exit Sub
    End Select
    If Not IsNumeric(text) Then GoTo InvalidValue
    value = CDbl(text)
    If value < minimum Or value > maximum Then GoTo InvalidValue
    Select Case control.id
        Case "edtFlatBarHeight"
            gblFlatBarHeight = value
            If gblMarkerKind = 3 Then gbldblIndicatorSize = value
        Case "edtFlatBarGap"
            gblFlatBarGap = value
        Case "edtFlatBarRoundness"
            gblFlatBarRoundness = value
        Case "edtFlatBarGlowRadius"
            gblFlatBarGlowRadius = value
        Case "edtFlatBarGlowTransparency"
            gblFlatBarGlowTransparency = value
    End Select
    PersistPISettings SerializeProgressOptionsRibbon()
    RefreshRibbon
    If Not gobjRibbon Is Nothing Then gobjRibbon.Invalidate
    RefreshProgressBar
    Exit Sub
InvalidValue:
    MsgBox PI_Zh("8BF78F9351650020") & CStr(minimum) & PI_Zh("002081F30020") & CStr(maximum) & _
        PI_Zh("00204E4B95F476846570503C3002"), vbExclamation, PI_Zh("570689D26A2A6761")
End Sub

Public Function PI_BarClamp(ByVal value As Double, ByVal minimum As Double, ByVal maximum As Double) As Double
    If value < minimum Then value = minimum
    If value > maximum Then value = maximum
    PI_BarClamp = value
End Function

Public Function PI_MarkerHeight(opts As ProgressDotsOptions) As Double
    If opts.FlatBar = 1 And opts.MarkerKind = 3 Then
        PI_MarkerHeight = opts.dotSize
    ElseIf opts.FlatBar = 1 Then
        PI_MarkerHeight = PI_BarClamp(opts.FlatBarHeight, 0.5, 12)
    Else
        PI_MarkerHeight = opts.dotSize
    End If
End Function

Private Sub PI_GlyphDimensions(ByVal kind As Integer, ByVal height As Double, ByVal slot As Double, _
                               ByVal gap As Double, ByVal ordinal As Long, ByVal pageDelta As Long, _
                               ByRef along As Double, ByRef across As Double)
    Dim stateScale As Double
    along = slot - PI_BarClamp(gap, 0, slot / 2)
    across = height
    Select Case kind
        Case 0
            If across > along / 3 Then across = along / 3
        Case 1
            If pageDelta <> 0 Then across = across * PI_INACTIVE_SEGMENT_HEIGHT
        Case 2
            If along > height * 0.32 Then along = height * 0.32
            If pageDelta = 0 Then
                along = PI_BarClamp(along * 2.8, 0, slot)
            ElseIf (ordinal - 1) Mod 5 <> 0 Then
                across = across * 0.63
            End If
        Case 3, 5
            stateScale = 1
            If pageDelta < 0 Then stateScale = 0.54
            If pageDelta > 0 Then stateScale = 0.78
            If along > height Then along = height
            If kind = 5 Then across = along
            along = along * stateScale
            across = across * stateScale
        Case 4
            along = slot
            If pageDelta < 0 Then across = across * 0.5
            If pageDelta > 0 Then across = across * 0.25
    End Select
End Sub

Private Function PI_UsesCrimsonMarkers(opts As ProgressDotsOptions) As Boolean
    PI_UsesCrimsonMarkers = (opts.CrimsonType = 1 And opts.MarkerKind = 0)
End Function

Private Function PI_CrimsonMarkerWidth(opts As ProgressDotsOptions, ByVal sectionWidth As Double, _
                                      ByVal count As Long, ByRef gap As Double) As Double
    Dim span As Double
    span = sectionWidth * PI_CRIMSON_MARKER_SPAN
    gap = PI_BarClamp(opts.FlatBarGap * (1 - 0.25 * opts.PresetSize), 0, span / count / 2)
    PI_CrimsonMarkerWidth = (span - (count - 1) * gap) / count
End Function

Private Sub PI_PresetMarkerGeometry(opts As ProgressDotsOptions, ByVal sectionLeft As Double, _
                        ByVal sectionWidth As Double, ByVal ordinal As Long, ByVal count As Long, _
                        ByVal pageDelta As Long, ByVal rowHeight As Double, _
                        ByRef left As Double, ByRef width As Double, ByRef height As Double, ByRef slot As Double)
    Dim blend As Double, largeHeight As Double, gap As Double, largeCenter As Double
    Dim smallSlot As Double, smallWidth As Double, firstWidth As Double, lastWidth As Double
    Dim smallCenter As Double, firstCenter As Double, lastCenter As Double, unusedHeight As Double
    blend = opts.PresetSize / 2
    largeHeight = rowHeight / (1 - blend / 2)
    slot = sectionWidth / count
    PI_GlyphDimensions opts.MarkerKind, largeHeight, slot, opts.FlatBarGap, ordinal, pageDelta, width, height
    If PI_UsesCrimsonMarkers(opts) Then
        width = PI_CrimsonMarkerWidth(opts, sectionWidth, count, gap)
        slot = width + gap
        left = sectionLeft + sectionWidth * (1 - PI_CRIMSON_MARKER_SPAN) / 2 + (ordinal - 1) * slot
        height = height * (1 - blend / 2)
        Exit Sub
    End If
    largeCenter = sectionLeft + (ordinal - 0.5) * slot
    left = largeCenter - width / 2
    If blend = 0 Then Exit Sub

    smallSlot = 0.9 * slot
    If count = 1 Then
        smallWidth = 0.9 * sectionWidth
        smallCenter = sectionLeft + sectionWidth / 2
    Else
        ' The first and last visible strokes, including outlines, set the limits of the row.
        PI_GlyphDimensions opts.MarkerKind, largeHeight / 2, smallSlot, opts.FlatBarGap / 2, _
                           ordinal, pageDelta, smallWidth, unusedHeight
        PI_GlyphDimensions opts.MarkerKind, largeHeight / 2, smallSlot, opts.FlatBarGap / 2, _
                           1, pageDelta + 1 - ordinal, firstWidth, unusedHeight
        PI_GlyphDimensions opts.MarkerKind, largeHeight / 2, smallSlot, opts.FlatBarGap / 2, _
                           count, pageDelta + count - ordinal, lastWidth, unusedHeight
        firstCenter = sectionLeft + 0.05 * sectionWidth + firstWidth / 2
        lastCenter = sectionLeft + 0.95 * sectionWidth - lastWidth / 2
        smallCenter = firstCenter + (lastCenter - firstCenter) * (ordinal - 1) / (count - 1)
    End If
    width = width + blend * (smallWidth - width)
    height = height * (1 - blend / 2)
    left = largeCenter + blend * (smallCenter - largeCenter) - width / 2
End Sub

Public Sub PI_BarApplyGlow(ByVal target As Shape, opts As ProgressDotsOptions, ByVal isCurrent As Boolean)
    Dim effect As Object, failureNumber As Long, failureDescription As String
    On Error GoTo GlowUnavailable
    Set effect = target.Glow
    effect.Radius = 0
    effect.Transparency = 1
    If Not isCurrent Or opts.FlatBarGlow <> 1 Then Exit Sub
    effect.Color.RGB = target.Fill.ForeColor.RGB
    effect.Transparency = PI_BarClamp(opts.FlatBarGlowTransparency, 0, 100) / 100
    effect.Radius = PI_BarClamp(opts.FlatBarGlowRadius, 0, 20)
    Exit Sub
GlowUnavailable:
    failureNumber = Err.Number
    failureDescription = Err.Description
    If failureNumber = 438 Or failureNumber = 445 Or failureNumber = -2147467263 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    Err.Raise vbObjectError + 3022, "PI_BarApplyGlow", failureDescription
End Sub

Private Sub PI_SizePageMarker(ByVal marker As Shape, opts As ProgressDotsOptions, ByVal relativePage As Integer)
    Dim dimensions() As String, along As Double, across As Double, slot As Double
    Dim centerX As Double, centerY As Double, inset As Double
    dimensions = Split(marker.Tags("PI_GLYPH_SIZE"), "|")
    along = Val(dimensions(0))
    across = Val(dimensions(1))
    slot = Val(dimensions(2))
    Select Case opts.MarkerKind
        Case 0, 1
            If relativePage <> 0 Then across = across * PI_INACTIVE_SEGMENT_HEIGHT
        Case 2
            If relativePage = 0 Then
                along = PI_BarClamp(along * 2.8, 0, slot)
                across = opts.FlatBarHeight
            End If
        Case 3, 5
            If relativePage < 0 Then along = along * 0.54
            If relativePage > 0 Then along = along * 0.78
            across = along
        Case 4
            If relativePage < 0 Then across = across * 0.5
            If relativePage > 0 Then across = across * 0.25
    End Select
    ' The center stays in its allocated page slot on all four edges.
    centerX = marker.Left + marker.Width / 2
    centerY = marker.Top + marker.Height / 2
    If LCase$(marker.Tags("PI_OUTLINED")) = "yes" Then
        inset = marker.Line.Weight
        If inset > along / 3 Then inset = along / 3
        If inset > across / 3 Then inset = across / 3
        marker.Line.Weight = inset
    End If
    marker.LockAspectRatio = msoFalse
    If opts.DockEdge < 2 Then
        marker.Width = along - inset
        marker.Height = across - inset
    Else
        marker.Width = across - inset
        marker.Height = along - inset
    End If
    marker.Left = centerX - marker.Width / 2
    marker.Top = centerY - marker.Height / 2
End Sub

Private Sub PI_FinalizeNavigationGlow(ByVal navigation As Shape, opts As ProgressDotsOptions, _
                                       ByVal visibleIndex As Long, ByVal currentSection As Long)
    Dim index As Long, role As String, isCurrent As Boolean
    Dim relativePage As Integer, transparency As Double
    If opts.FlatBar <> 1 Then Exit Sub
    If navigation.Type = msoGroup Then
        For index = 1 To navigation.GroupItems.Count
            PI_FinalizeNavigationGlow navigation.GroupItems.Item(index), opts, visibleIndex, currentSection
        Next index
    Else
        role = LCase$(navigation.Tags("PI_ROLE"))
        If role = "indicator" Or role = "indicator-summary" Then
            relativePage = Sgn(CLng(Val(navigation.Tags("PI_PAGE"))) - visibleIndex)
            isCurrent = (role = "indicator" And relativePage = 0)
            If (PI_UsesCrimsonMarkers(opts) And Not isCurrent) Or _
                    (opts.PresetSize = 0 And opts.MarkerKind > 0 And opts.MarkerKind <> 3) Then
                PI_SizePageMarker navigation, opts, relativePage
            End If
            PI_BarApplyGlow navigation, opts, isCurrent
            If opts.CrimsonType = 1 Or (opts.LayeredNavigation = 1 And PI_UsesSectionBand(opts)) Then
                ' Apply opacity after the final fill color, outline, and dock transform.
                Select Case relativePage
                    Case -1: transparency = opts.transparencyIndicatorSeen
                    Case 0: transparency = opts.transparencyIndicatorActive
                    Case 1: transparency = opts.transparencyIndicatorUnseen
                End Select
                If opts.CrimsonType = 1 And relativePage <> 0 And _
                        opts.InactiveSectionBarTransparency <> PI_CRIMSON_NONCURRENT_PAGE_TRANSPARENCY Then
                    ' Only an explicitly configured section override needs the chapter lookup.
                    If transparency = PI_CRIMSON_NONCURRENT_PAGE_TRANSPARENCY And _
                            CLng(navigation.Tags("PI_SECTION")) <> currentSection Then
                        transparency = opts.InactiveSectionBarTransparency
                    End If
                End If
                transparency = transparency / 100
                If LCase$(navigation.Tags("PI_OUTLINED")) = "yes" Then
                    navigation.Line.Transparency = transparency
                Else
                    navigation.Fill.Transparency = transparency
                End If
            End If
        End If
    End If
End Sub

Public Sub PI_BandPreset(control As IRibbonControl)
    If gblbIndicatorUpdateRunning Then Exit Sub
    Dim spacingMode As String, previous As ProgressDotsOptions
    If Application.Presentations.Count = 0 Then Exit Sub
    previous = DeserializeProgressOptions(SerializeProgressOptionsRibbon())
    spacingMode = gblEvenSpaceSecOrInd
    PI_SetBandDefaults
    gblBandHeightMode = PI_PROPORTIONAL_BAND
    PI_RestoreStylePolicy previous
    PI_SetFullWidth PI_StyleAxisLength()
    If spacingMode = "Sections" Or spacingMode = "Shapes" Then gblEvenSpaceSecOrInd = spacingMode
    PersistPISettings SerializeProgressOptionsRibbon()
    RefreshRibbon
    If Not gobjRibbon Is Nothing Then gobjRibbon.Invalidate
End Sub

Private Sub PI_SetBandDefaults()
    gblPresetSize = 0
    gblCrimsonType = 0
    gblInactiveSectionBarTransparency = 0
    gblBandHeightMode = 0
    gblTitleAlignment = 0
    gblTitleCase = 0
    gblMarkerKind = 0
    gblPanelShadow = 0
    gblFlatRuleMode = 0
    gblFlatRuleWidth = 0.5
    gblFlatRuleColor = "107;114;128"
    gblUseThemeColors = 0
    gblSectionRoundness = 0
    gblHardShadowOffset = 0
    gblAccentMode = 0
    gblAccentWidth = 1.5
    gblAccentColor = "158;31;56"
    gblSectionNumbering = 0
    gblCurrentTitleBold = 0
    gblFutureBarOutline = 0
    gblSectionForegroundBars = 0
    gblLayeredNavigation = 0
    gblSectionBand = 1
    gblSectionBandPadding = 4
    gblSectionBandContrast = 1
    gblSectionBandGradient = 1
    gblSectionBandTransparencyCurrent = 0
    gblSectionBandTransparencyLeft = 0
    gblSectionBandTransparencyRight = 0
    gblSectionBandLeftNear = "96;112;132"
    gblSectionBandLeftFar = "203;210;219"
    gblSectionBandRightNear = "184;196;210"
    gblSectionBandRightFar = "229;233;238"
    gblSectionBandBackdrop = 0
    gblSectionBandBackdropColor = "255;255;255"
    gblSectionBandBackdropTransparency = 75
    gblSectionBandBackdropScale = 150
    gblFlatBar = 1
    gblFlatBarHeight = 2.5
    gblFlatBarGap = 3
    gblFlatBarRoundness = 50
    gblFlatBarGlow = 1
    gblFlatBarGlowRadius = 4
    gblFlatBarGlowTransparency = 55
    gblstrIndicatorStyle = "Full"
    gblEvenSpaceSecOrInd = "Sections"
    gblstrSlideIgnoreSections = "Full"
    gblstrSlideIgnoreIndicator = "All"
    gblintSectionPlacement = 0
    gblintGeneralProgressbarDirect = 0
    gblintUseTextBGCol = 1
    gblintUseTextBGCol_Current = 1
    gblintUseTextBGCol_UnSeen = 1
    gblstrFontBGColorSeen = "96;112;132"
    gblstrFontBGColorCurrent = "181;42;69"
    gblstrFontBGColorUnSeen = "229;233;238"
    gblstrFontColorSeen = "255;255;255"
    gblstrFontColorCurrent = "255;255;255"
    gblstrFontColorUnSeen = "65;75;85"
    gblTransparencyFontSeen = 0
    gblTransparencyFontActive = 0
    gblTransparencyFontUnseen = 0
    gblstrFontname = "Arial"
    gbldblFontSize = 12
    gblstrIndicatorType = "Oval"
    gblintIndicatorTypeID = 6
    gblshowCustomImage = 0
    gbldblIndicatorSize = 5
    gblstrIndicatorColorSeen = "255;255;255"
    gblstrIndicatorColorCurrent = "255;255;255"
    gblstrIndicatorColorUnSeen = "65;75;85"
    gblintUseShapeBG_Seen = 1
    gblintUseShapeBG_Current = 1
    gblintUseShapeBG_UnSeen = 0
    gblTransparencyIndicatorSeen = 35
    gblTransparencyIndicatorActive = 0
    gblTransparencyIndicatorUnseen = 35
    gblShapeLineWeight = 1
    gblShapeLineStyle = 1
    gblShapeLineDashStyle = 1
    gblintSectionSpacing = 0
    gblintIndicatorSqueeze = 0
    gblintSectionIndicatorGap = 3
End Sub

Public Sub PI_InitializeUniversalStyle(ByVal pageWidth As Double)
    PI_SetBandDefaults
    PI_ApplyStyleValues PI_StyleCommon()
    PI_ApplyStyleValues PI_StyleRecipe("universal_theme")
    gblBandHeightMode = PI_PROPORTIONAL_BAND
    gblintMarginTop = 0
    PI_SetFullWidth pageWidth
End Sub

Private Function PI_StyleAxisLength() As Double
    If Application.Presentations.Count = 0 Then Exit Function
    Dim renderDeck As Presentation
    Set renderDeck = PI_RenderPresentation()
    If gblDockEdge = 2 Or gblDockEdge = 3 Then
        PI_StyleAxisLength = renderDeck.PageSetup.SlideHeight
    Else
        PI_StyleAxisLength = renderDeck.PageSetup.SlideWidth
    End If
End Function

Public Sub PI_InitializeLayerStyle(ByVal pageWidth As Double)
    PI_SetBandDefaults
    PI_ApplyStyleValues PI_StyleCommon()
    PI_ApplyStyleValues PI_StyleRecipe("polyu_layered")
    gblBandHeightMode = PI_PROPORTIONAL_BAND
    PI_SetFullWidth pageWidth
End Sub

Public Sub PI_LayerPreset(control As IRibbonControl)
    If gblbIndicatorUpdateRunning Then Exit Sub
    Dim spacingMode As String, previous As ProgressDotsOptions
    If Application.Presentations.Count = 0 Then Exit Sub
    previous = DeserializeProgressOptions(SerializeProgressOptionsRibbon())
    spacingMode = gblEvenSpaceSecOrInd
    PI_InitializeLayerStyle PI_StyleAxisLength()
    PI_RestoreStylePolicy previous
    If spacingMode = "Sections" Or spacingMode = "Shapes" Then gblEvenSpaceSecOrInd = spacingMode
    PersistPISettings SerializeProgressOptionsRibbon()
    RefreshRibbon
    If Not gobjRibbon Is Nothing Then gobjRibbon.Invalidate
End Sub

Public Sub PI_LayerGetValue(control As IRibbonControl, ByRef value)
    Select Case control.id
        Case "edtNavigationWidth": value = CStr(gblNavigationWidth)
        Case "edtSectionBlockHeight"
            If gblBandHeightMode = PI_PROPORTIONAL_BAND Then
                Dim opts As ProgressDotsOptions, titleTop As Double, dotsTop As Double
                Dim bandHeight As Double, growth As Double, physicalBody As Double
                opts = getDefaultOptions()
                opts.DockEdge = 0
                PI_BandGeometry opts, titleTop, dotsTop, bandHeight, growth, physicalBody
                value = CStr(Round(physicalBody, 4))
            Else
                value = CStr(gblSectionBlockHeight)
            End If
        Case "edtRailWidth": value = CStr(gblRailWidth)
        Case "edtRailThickness": value = CStr(gblRailThickness)
        Case "edtLayerShadowBlur": value = CStr(gblLayerShadowBlur)
        Case "edtLayerShadowTransparency": value = CStr(gblLayerShadowTransparency)
        Case "edtLayerDepthStep": value = CStr(gblLayerDepthStep)
        Case "edtLayerShadowDistance": value = CStr(gblLayerShadowDistance)
        Case "edtAccentWidth": value = CStr(gblAccentWidth)
    End Select
End Sub

Private Sub PI_ApplyNamedPreset(ByVal styleID As String, previous As ProgressDotsOptions)
    Dim values As String, parts() As String, presetSize As Integer
    parts = Split(styleID, ":")
    If UBound(parts) > 0 Then
        Select Case parts(1)
            Case "medium": presetSize = 1
            Case "small": presetSize = 2
            Case Else: Err.Raise 5, "PI_ApplyNamedPreset", "Unknown preset size: " & parts(1)
        End Select
        styleID = parts(0)
    End If
    values = PI_StyleRecipe(styleID)
    If Len(values) = 0 Then Err.Raise 5, "PI_ApplyNamedPreset", "Unknown preset: " & styleID
    PI_SetBandDefaults
    PI_ApplyStyleValues PI_StyleCommon()
    PI_ApplyStyleValues values
    PI_ApplyPresetSize presetSize
    PI_RestoreStylePolicy previous
    gblEvenSpaceSecOrInd = previous.EvenSpaceSecOrInd
    PI_SetFullWidth PI_StyleAxisLength()
End Sub

Private Sub PI_ApplyPresetSize(ByVal presetSize As Integer)
    Dim blend As Double, largeHeight As Double, titleTop As Double, dotsTop As Double
    Dim largeOptions As ProgressDotsOptions
    If gblMarkerKind = 3 Then gbldblIndicatorSize = gblFlatBarHeight
    If presetSize = 0 Then
        gblBandHeightMode = PI_PROPORTIONAL_BAND
        Exit Sub
    End If
    blend = presetSize / 2
    largeOptions = getDefaultOptions()
    largeOptions.DockEdge = 0
    PI_BandGeometry largeOptions, titleTop, dotsTop, largeHeight
    If largeOptions.LayeredNavigation = 0 Then largeHeight = largeHeight - PI_BarClamp(largeOptions.HardShadowOffset, 0, largeHeight / 3)
    gblPresetSize = presetSize
    If presetSize = 2 Then
        If gblCrimsonType = 1 Then
            gblSectionBlockHeight = PI_CRIMSON_SMALL_BODY_HEIGHT
        Else
            gblSectionBlockHeight = PI_SMALL_BODY_HEIGHT
        End If
    Else
        gblSectionBlockHeight = (largeHeight + 7) / 2 - 4
        gblBandHeightMode = PI_PROPORTIONAL_BAND
    End If
    gbldblFontSize = gbldblFontSize + blend * (7 - gbldblFontSize)
    gblFlatBarHeight = gblFlatBarHeight * (1 - blend / 2)
    gbldblIndicatorSize = gbldblIndicatorSize * (1 - blend / 2)
End Sub

Public Sub PI_ModernPreset(control As IRibbonControl)
    Dim previous As ProgressDotsOptions
    If gblbIndicatorUpdateRunning Then Exit Sub
    If Application.Presentations.Count = 0 Then Exit Sub
    previous = DeserializeProgressOptions(SerializeProgressOptionsRibbon())
    PI_ApplyNamedPreset control.Tag, previous
    PersistPISettings SerializeProgressOptionsRibbon()
    RefreshRibbon
    If Not gobjRibbon Is Nothing Then gobjRibbon.Invalidate
End Sub

Private Sub PI_RestoreStylePolicy(previous As ProgressDotsOptions)
    gblstrSlideIgnoreSections = previous.ignoreSections
    gblstrSlideIgnoreIndicator = previous.ignoreIndicator
    gblintGeneralProgressbarDirect = previous.DirectUpdate
End Sub

Private Function PI_StyleCommon() As String
    PI_StyleCommon = "layered=0|height=34|rail=1|padding=4|textSize=12|textGap=3|barHeight=2.5|barGap=3|roundness=50|glow=0|glowRadius=2|glowTransparency=65|shadowBlur=0|shadowTransparency=80|shadowDistance=1.5|depthStep=0.75|seenBackground=#E6E8EC|currentBackground=#9E1F38|unseenBackground=#E6E8EC|seenForeground=#414B55|currentForeground=#FFFFFF|unseenForeground=#414B55|seenBarTransparency=18|currentBarTransparency=0|unseenBarTransparency=55|currentTransparency=0|leftTransparency=0|rightTransparency=0|contrast=0|gradient=0|leftNear=#DADDE2|leftFar=#F0F1F3|rightNear=#E3E5E9|rightFar=#F4F5F7|accentMode=0|accentWidth=1.5|" & _
        "accentColor=#9E1F38|numbering=0|currentBold=0|futureOutline=0|sectionBars=1|themeColors=0|sectionRoundness=0|sectionGap=0|hardShadow=0|fontName=Arial|titleAlignment=0|titleCase=0|flatRuleMode=1|flatRuleWidth=0.8|flatRuleColor=#9199A4|markerKind=0|panelShadow=0"
End Function

Private Function PI_StyleRecipe(ByVal styleID As String) As String
    Select Case styleID
        Case "universal_theme"
            PI_StyleRecipe = "layered=0|themeColors=1|seenBackground=#DCE1E7|currentBackground=#2563EB|unseenBackground=#F1F5F9|seenForeground=#334155|unseenForeground=#334155|accentMode=1|accentColor=#2563EB|currentBold=1|contrast=1|markerKind=4|flatRuleMode=1|barHeight=4"
        Case "expressive_pill"
            PI_StyleRecipe = "layered=0|sectionRoundness=50|sectionGap=6|height=38|seenBackground=#DBD4E4|currentBackground=#65558F|unseenBackground=#EDE6F5|seenForeground=#4A4458|unseenForeground=#4A4458|barHeight=6|currentBold=1|markerKind=3|flatRuleMode=0|futureOutline=1"
        Case "bento_segments"
            PI_StyleRecipe = "layered=0|sectionRoundness=18|sectionGap=4|height=38|seenBackground=#DDE2E7|currentBackground=#DCEAFE|unseenBackground=#F1F4F8|seenForeground=#3E4C5D|currentForeground=#163D73|unseenForeground=#3E4C5D|numbering=1|currentBold=1|markerKind=5|flatRuleMode=0|barHeight=6|futureOutline=1"
        Case "glass_dock"
            PI_StyleRecipe = "layered=0|sectionRoundness=40|sectionGap=5|height=38|seenBackground=#9AB1CE|currentBackground=#F8FBFF|unseenBackground=#A7BEDB|seenForeground=#314862|currentForeground=#1D3D60|unseenForeground=#314862|leftTransparency=65|rightTransparency=65|currentBold=1|markerKind=3|flatRuleMode=0|barHeight=5|futureOutline=1"
        Case "soft_minimal"
            PI_StyleRecipe = "layered=0|sectionRoundness=16|sectionGap=2|seenBackground=#E0E4DF|currentBackground=#E4E9E2|unseenBackground=#F2F4F0|seenForeground=#526054|currentForeground=#304236|unseenForeground=#526054|barHeight=4|currentBold=1|markerKind=4|flatRuleMode=0"
        Case "hard_offset"
            PI_StyleRecipe = "layered=0|sectionGap=6|hardShadow=2|seenBackground=#E5E3D6|currentBackground=#FFD449|unseenBackground=#FFFCEE|seenForeground=#151515|currentForeground=#151515|unseenForeground=#151515|accentMode=3|accentColor=#151515|accentWidth=1.5|numbering=1|currentBold=1|futureOutline=1|barHeight=6|roundness=0|markerKind=5|flatRuleMode=0"
        Case "retro_terminal"
            PI_StyleRecipe = "layered=0|fontName=Courier New|seenBackground=#16261E|currentBackground=#15281F|unseenBackground=#07140E|seenForeground=#93B7A2|currentForeground=#B9F7A3|unseenForeground=#93B7A2|numbering=1|currentBold=1|futureOutline=1|barHeight=6|roundness=0|markerKind=5|flatRuleMode=1"
        Case "serif_editorial"
            PI_StyleRecipe = "layered=0|fontName=Times New Roman|sectionGap=4|height=38|textSize=12.5|textGap=4|seenBackground=#E4E0DA|currentBackground=#FFFFFF|unseenBackground=#F8F5EF|seenForeground=#3F372F|currentForeground=#5C3F2B|unseenForeground=#3F372F|numbering=1|currentBold=1|barHeight=6|roundness=0|accentMode=4|accentColor=#8C6749|accentWidth=1|markerKind=2|flatRuleMode=3"
        Case "polyu_layered"
            PI_StyleRecipe = "shadowBlur=10|glow=1|glowRadius=4|glowTransparency=55|contrast=1|markerKind=0|flatRuleMode=0|barHeight=2.5|layered=1|shadowTransparency=62|shadowDistance=2.5|depthStep=0.8|flatRuleColor=#858F9A|flatRuleWidth=0.85|seenBackground=#DEDEE3|unseenBackground=#F2F2F7|inactiveSectionBarTransparency=85|seenBarTransparency=85|unseenBarTransparency=85|fontName=Inter Medium|currentBold=1|crimsonType=1"
        Case "quiet_layered"
            PI_StyleRecipe = "seenBackground=#DEDEE3|currentBackground=#FFFFFF|unseenBackground=#F2F2F7|seenForeground=#3C3C43|currentForeground=#9E1F38|unseenForeground=#3C3C43|shadowBlur=2|shadowTransparency=64|shadowDistance=1.3|depthStep=0.25|accentMode=1|glow=1|markerKind=3|flatRuleMode=0|barHeight=5|futureOutline=1|panelShadow=1|sectionGap=4"
        Case "minimal_underline"
            PI_StyleRecipe = "seenBackground=#EAEBEC|currentBackground=#FFFFFF|unseenBackground=#FFFFFF|currentForeground=#9E1F38|accentMode=1|accentWidth=2|markerKind=4|flatRuleMode=1|barHeight=4"
        Case "minimal_topline"
            PI_StyleRecipe = "seenBackground=#EAEBEC|currentBackground=#FAF3F5|unseenBackground=#FFFFFF|currentForeground=#9E1F38|accentMode=2|accentWidth=2.5|currentBold=1|markerKind=2|flatRuleMode=1|barHeight=6"
        Case "material_tonal"
            PI_StyleRecipe = "seenBackground=#E1DCE4|currentBackground=#E8DEF8|unseenBackground=#F4EFF6|seenForeground=#49454F|currentForeground=#21005D|unseenForeground=#49454F|accentMode=1|accentColor=#6750A4|accentWidth=2|currentBold=1|markerKind=3|flatRuleMode=0|barHeight=6|futureOutline=1|sectionGap=3|sectionRoundness=15"
        Case "material_filled"
            PI_StyleRecipe = "seenBackground=#E4DFE7|currentBackground=#6750A4|unseenBackground=#F7F2FA|seenForeground=#49454F|unseenForeground=#49454F|glow=1|glowTransparency=70|markerKind=1|flatRuleMode=0|barHeight=4|sectionGap=3|sectionRoundness=0"
        Case "fluent_soft"
            PI_StyleRecipe = "seenBackground=#E1E1E1|currentBackground=#EBF3FC|unseenBackground=#F5F5F5|seenForeground=#424242|currentForeground=#0F548C|unseenForeground=#424242|shadowBlur=3|shadowTransparency=63|shadowDistance=1.3|depthStep=0.25|accentMode=4|accentColor=#0F6CBD|accentWidth=2|currentBold=1|markerKind=4|flatRuleMode=0|barHeight=4|panelShadow=1|sectionGap=4"
        Case "fluent_outline"
            PI_StyleRecipe = "seenBackground=#E8E8E8|currentBackground=#EFF6FC|unseenBackground=#FFFFFF|seenForeground=#323130|currentForeground=#0F548C|unseenForeground=#323130|accentMode=3|accentColor=#0F6CBD|accentWidth=0.8|markerKind=5|flatRuleMode=4|barHeight=5|futureOutline=1"
        Case "editorial_numbered"
            PI_StyleRecipe = "seenBackground=#DFDDD8|currentBackground=#FFFFFF|unseenBackground=#F7F5F0|seenForeground=#191919|currentForeground=#9E1F38|unseenForeground=#191919|numbering=1|currentBold=1|accentMode=4|accentWidth=2|barHeight=6|roundness=0|markerKind=2|flatRuleMode=1"
        Case "swiss_mono"
            PI_StyleRecipe = "seenBackground=#EAEAEA|currentBackground=#111111|unseenBackground=#FFFFFF|seenForeground=#444444|unseenForeground=#444444|numbering=1|currentBold=1|barHeight=3|roundness=0|markerKind=1|flatRuleMode=4|futureOutline=1"
        Case "contrast_light"
            PI_StyleRecipe = "seenBackground=#E3E3E3|currentBackground=#111111|unseenBackground=#FFFFFF|seenForeground=#000000|unseenForeground=#000000|seenBarTransparency=0|unseenBarTransparency=0|currentBold=1|futureOutline=1|accentMode=3|accentColor=#FFD54F|barHeight=7|barGap=4|roundness=20|markerKind=5|flatRuleMode=4|flatRuleColor=#363636"
        Case "contrast_dark"
            PI_StyleRecipe = "seenBackground=#2F2F2F|currentBackground=#FFFFFF|unseenBackground=#151515|seenForeground=#FFFFFF|currentForeground=#111111|unseenForeground=#FFFFFF|seenBarTransparency=0|unseenBarTransparency=0|currentBold=1|futureOutline=1|barHeight=7|barGap=4|roundness=20|markerKind=3|flatRuleMode=4|flatRuleColor=#A6ADB4"
        Case "glass_light"
            PI_StyleRecipe = "seenBackground=#B9C3D0|currentBackground=#EEF5FF|unseenBackground=#CBD5E1|seenForeground=#263549|currentForeground=#1E3A5F|unseenForeground=#263549|leftTransparency=55|rightTransparency=55|shadowBlur=3|shadowTransparency=71|shadowDistance=1.3|depthStep=0|accentMode=3|accentColor=#93A9C4|accentWidth=0.75|markerKind=3|flatRuleMode=4|barHeight=5|futureOutline=1|panelShadow=1|sectionGap=4|flatRuleColor=#93A9C4"
        Case "glass_smoke"
            PI_StyleRecipe = "seenBackground=#3E4A5D|currentBackground=#20334E|unseenBackground=#263449|seenForeground=#FFFFFF|unseenForeground=#FFFFFF|leftTransparency=35|rightTransparency=35|shadowBlur=2|shadowTransparency=62|shadowDistance=1.3|depthStep=0|accentMode=3|accentColor=#9AB8DA|accentWidth=0.75|glow=1|markerKind=4|flatRuleMode=4|barHeight=4|panelShadow=1|sectionGap=4|flatRuleColor=#9AB8DA"
        Case "neon_dark"
            PI_StyleRecipe = "seenBackground=#262D3C|currentBackground=#17283D|unseenBackground=#111827|seenForeground=#CED7E4|currentForeground=#65E4FF|unseenForeground=#CED7E4|accentMode=1|accentColor=#65E4FF|accentWidth=1.2|glow=1|glowRadius=3|glowTransparency=45|currentBold=1|futureOutline=1|markerKind=3|flatRuleMode=1|barHeight=5"
        Case "blueprint"
            PI_StyleRecipe = "seenBackground=#203D5B|currentBackground=#123D66|unseenBackground=#0B2948|seenForeground=#C7DEF5|unseenForeground=#C7DEF5|accentMode=3|accentColor=#79C4F2|accentWidth=0.75|numbering=1|currentBold=1|futureOutline=1|barHeight=7|roundness=0|markerKind=2|flatRuleMode=7|flatRuleColor=#79C4F2"
        Case "duotone_gradient"
            PI_StyleRecipe = "currentBackground=#6D28D9|gradient=1|contrast=1|leftNear=#DDD6FE|leftFar=#F5F3FF|rightNear=#D6F0EE|rightFar=#F0FDFA|glow=1|markerKind=4|flatRuleMode=1|barHeight=5|seenBackground=#D4D7DB"
        Case "distance_gradient"
            PI_StyleRecipe = "gradient=1|contrast=1|leftNear=#A6B0BC|leftFar=#D7DDE4|rightNear=#DADDE2|rightFar=#F5F5F6|currentBold=1|glow=1|markerKind=3|flatRuleMode=1|barHeight=5|futureOutline=1|seenBackground=#D4D7DB"
        Case "paper_warm"
            PI_StyleRecipe = "seenBackground=#E0DCD2|currentBackground=#DED2BD|unseenBackground=#F4F0E6|seenForeground=#40382F|currentForeground=#3E3226|unseenForeground=#40382F|accentMode=2|accentColor=#8C6B45|accentWidth=1|shadowBlur=2|shadowTransparency=62|shadowDistance=1.3|depthStep=0.1|barHeight=6|roundness=20|currentBold=1|markerKind=2|flatRuleMode=8|panelShadow=1|sectionGap=4|fontName=Georgia|titleAlignment=1"
        Case "bold_graphic"
            PI_StyleRecipe = "seenBackground=#E5E3DB|currentBackground=#FFD449|unseenBackground=#FFFDF4|seenForeground=#151515|currentForeground=#151515|unseenForeground=#151515|accentMode=3|accentColor=#151515|accentWidth=2.5|numbering=1|currentBold=1|futureOutline=1|barHeight=7|roundness=0|markerKind=5|flatRuleMode=4"
        Case "stage_window"
            PI_StyleRecipe = "layered=0|height=38|padding=8|textSize=13|textGap=5|fontName=Arial|titleAlignment=1|titleCase=1|currentBold=1|numbering=0|sectionGap=0|sectionRoundness=0|seenBackground=#DCD9D2|currentBackground=#F33F2B|unseenBackground=#F5F2EA|seenForeground=#101010|currentForeground=#101010|unseenForeground=#101010|flatRuleMode=4|flatRuleWidth=1.2|flatRuleColor=#101010|barHeight=6|roundness=0|accentMode=0|glow=0|shadowBlur=0|gradient=0|markerKind=5|futureOutline=1"
        Case "engineering_register"
            PI_StyleRecipe = "layered=0|height=40|padding=10|textSize=12|textGap=5|fontName=Arial|titleAlignment=1|titleCase=0|currentBold=1|numbering=1|sectionGap=0|sectionRoundness=0|seenBackground=#D8D2C4|currentBackground=#A8B4BD|unseenBackground=#EFE8D8|seenForeground=#202020|currentForeground=#202020|unseenForeground=#202020|flatRuleMode=7|flatRuleWidth=0.45|flatRuleColor=#202020|accentMode=4|accentColor=#202020|accentWidth=2|barHeight=7|roundness=0|glow=0|shadowBlur=0|gradient=0|markerKind=2"
        Case "civic_corners"
            PI_StyleRecipe = "layered=0|height=40|padding=10|textSize=13|textGap=5|fontName=Arial|titleAlignment=0|titleCase=1|currentBold=1|numbering=0|sectionGap=0|sectionRoundness=0|seenBackground=#DEE0E3|currentBackground=#FFFFFF|unseenBackground=#F4F5F8|seenForeground=#30333B|currentForeground=#2640ED|unseenForeground=#30333B|flatRuleMode=5|flatRuleWidth=1.4|flatRuleColor=#2640ED|accentMode=0|barHeight=6|roundness=0|glow=0|shadowBlur=0|gradient=0|markerKind=3|futureOutline=1"
        Case "community_cells"
            PI_StyleRecipe = "layered=0|height=40|padding=8|textSize=13|textGap=6|fontName=Arial|titleAlignment=1|titleCase=1|currentBold=1|numbering=1|sectionGap=0|sectionRoundness=0|seenBackground=#CDCCC3|currentBackground=#FF571A|unseenBackground=#E4E3D9|seenForeground=#151515|currentForeground=#151515|unseenForeground=#151515|flatRuleMode=4|flatRuleWidth=1.8|flatRuleColor=#151515|accentMode=0|barHeight=7|roundness=0|glow=0|shadowBlur=0|gradient=0|markerKind=5|futureOutline=1"
        Case "quiet_colophon"
            PI_StyleRecipe = "layered=0|height=40|padding=12|textSize=12|textGap=7|fontName=Georgia|titleAlignment=0|titleCase=0|currentBold=0|numbering=0|sectionGap=0|sectionRoundness=0|seenBackground=#E6E4DE|currentBackground=#EEE9DE|unseenBackground=#F8F6F0|seenForeground=#55534C|currentForeground=#25251F|unseenForeground=#55534C|flatRuleMode=3|flatRuleWidth=0.35|flatRuleColor=#A5A094|accentMode=0|barHeight=3.5|roundness=0|glow=0|shadowBlur=0|gradient=0|markerKind=4"
        Case "specimen_index"
            PI_StyleRecipe = "layered=0|height=39|padding=8|textSize=12|textGap=5|fontName=Arial|titleAlignment=1|titleCase=0|currentBold=1|numbering=1|sectionGap=0|sectionRoundness=0|seenBackground=#E6E4E1|currentBackground=#292929|unseenBackground=#FBF9F5|seenForeground=#3E3E3E|currentForeground=#FFFFFF|unseenForeground=#3E3E3E|flatRuleMode=2|flatRuleWidth=0.45|flatRuleColor=#F25B42|accentMode=0|barHeight=6|roundness=0|glow=0|shadowBlur=0|gradient=0|markerKind=2"
        Case "botanical_labels"
            PI_StyleRecipe = "layered=0|height=42|padding=10|textSize=14|textGap=8|fontName=Calibri Light|titleAlignment=1|titleCase=0|currentBold=0|numbering=0|sectionGap=0|sectionRoundness=0|seenBackground=#E0E3DB|currentBackground=#DDE9D6|unseenBackground=#F1F3EC|seenForeground=#52614F|currentForeground=#233D2A|unseenForeground=#52614F|flatRuleMode=8|accentMode=0|barHeight=5|roundness=0|glow=0|shadowBlur=0|gradient=0|markerKind=3|futureOutline=1"
        Case "archive_label"
            PI_StyleRecipe = "layered=0|height=40|padding=9|textSize=12|textGap=6|fontName=Courier New|titleAlignment=1|titleCase=0|currentBold=0|numbering=1|sectionGap=0|sectionRoundness=0|seenBackground=#DCD9D1|currentBackground=#FFFFFF|unseenBackground=#F0EDE5|seenForeground=#383633|currentForeground=#222222|unseenForeground=#383633|flatRuleMode=8|flatRuleWidth=0.45|flatRuleColor=#77736A|accentMode=4|accentWidth=2|accentColor=#AD342A|barHeight=5|roundness=0|glow=0|shadowBlur=0|gradient=0|markerKind=5|futureOutline=1"
        Case "coastal_bands"
            PI_StyleRecipe = "layered=0|height=38|padding=8|textSize=12|textGap=5|fontName=Arial|titleAlignment=1|titleCase=0|currentBold=0|numbering=0|sectionGap=0|sectionRoundness=0|seenBackground=#4B745F|currentBackground=#2556D8|unseenBackground=#EDE6CB|seenForeground=#FFFFFF|currentForeground=#FFFFFF|unseenForeground=#24382D|flatRuleMode=3|flatRuleWidth=2.5|flatRuleColor=#889680|accentMode=0|barHeight=5|roundness=0|glow=0|shadowBlur=0|gradient=0|markerKind=4"
        Case "swimlane_outline"
            PI_StyleRecipe = "layered=0|height=38|padding=8|textSize=12|textGap=5|fontName=Arial|titleAlignment=2|titleCase=0|currentBold=1|numbering=0|sectionGap=0|sectionRoundness=0|seenBackground=#E6EDF5|currentBackground=#E7F1F9|unseenBackground=#FFFFFF|seenForeground=#185CA0|currentForeground=#185CA0|unseenForeground=#185CA0|flatRuleMode=4|flatRuleWidth=0.55|flatRuleColor=#185CA0|accentMode=2|accentWidth=1.5|accentColor=#185CA0|barHeight=5|roundness=0|glow=0|shadowBlur=0|gradient=0|markerKind=4"
        Case "newspaper_columns"
            PI_StyleRecipe = "layered=0|height=40|padding=8|textSize=14|textGap=5|fontName=Georgia|titleAlignment=1|titleCase=0|currentBold=1|numbering=0|sectionGap=0|sectionRoundness=0|seenBackground=#E9E9E9|currentBackground=#F0F0EC|unseenBackground=#FFFFFF|seenForeground=#343434|currentForeground=#101010|unseenForeground=#343434|flatRuleMode=1|flatRuleWidth=0.45|flatRuleColor=#888888|accentMode=2|accentWidth=1.2|accentColor=#101010|barHeight=6|roundness=0|glow=0|shadowBlur=0|gradient=0|markerKind=2"
        Case "cabinet_grid"
            PI_StyleRecipe = "layered=0|height=38|padding=8|textSize=12|textGap=7|fontName=Arial|titleAlignment=1|titleCase=0|currentBold=0|numbering=1|sectionGap=0|sectionRoundness=0|seenBackground=#D0D0CE|currentBackground=#151515|unseenBackground=#E6E6E3|seenForeground=#202020|currentForeground=#FFFFFF|unseenForeground=#202020|flatRuleMode=4|flatRuleWidth=2|flatRuleColor=#777975|accentMode=0|barHeight=4|roundness=0|glow=0|shadowBlur=0|gradient=0|markerKind=1|futureOutline=1"
    End Select
End Function

Private Sub PI_ApplyStyleValues(ByVal values As String)
    Dim item As Variant, pair() As String, value As String
    For Each item In Split(values, "|")
        pair = Split(CStr(item), "=")
        If UBound(pair) = 1 Then
            value = pair(1)
            Select Case pair(0)
                Case "markerKind": gblMarkerKind = Val(value)
                Case "panelShadow": gblPanelShadow = Val(value)
                Case "themeColors": gblUseThemeColors = Val(value)
                Case "sectionRoundness": gblSectionRoundness = Val(value)
                Case "sectionGap": gblintSectionSpacing = Val(value)
                Case "hardShadow": gblHardShadowOffset = Val(value)
                Case "fontName": gblstrFontname = value
                Case "crimsonType": gblCrimsonType = Val(value)
                Case "titleAlignment": gblTitleAlignment = Val(value)
                Case "titleCase": gblTitleCase = Val(value)
                Case "flatRuleMode": gblFlatRuleMode = Val(value)
                Case "flatRuleWidth": gblFlatRuleWidth = Val(value)
                Case "flatRuleColor": gblFlatRuleColor = PI_StyleColor(value)
                Case "layered": gblLayeredNavigation = Val(value)
                Case "height": gblSectionBlockHeight = Val(value)
                Case "rail": gblRailThickness = Val(value)
                Case "padding": gblSectionBandPadding = Val(value)
                Case "textSize": gbldblFontSize = Val(value)
                Case "textGap": gblintSectionIndicatorGap = Val(value)
                Case "barHeight"
                    gblFlatBarHeight = Val(value)
                    If gblMarkerKind = 3 Then gbldblIndicatorSize = gblFlatBarHeight
                Case "dotSize": gbldblIndicatorSize = Val(value)
                Case "barGap": gblFlatBarGap = Val(value)
                Case "roundness": gblFlatBarRoundness = Val(value)
                Case "glow": gblFlatBarGlow = Val(value)
                Case "glowRadius": gblFlatBarGlowRadius = Val(value)
                Case "glowTransparency": gblFlatBarGlowTransparency = Val(value)
                Case "shadowBlur": gblLayerShadowBlur = Val(value)
                Case "shadowTransparency": gblLayerShadowTransparency = Val(value)
                Case "shadowDistance": gblLayerShadowDistance = Val(value)
                Case "depthStep": gblLayerDepthStep = Val(value)
                Case "inactiveSectionBarTransparency": gblInactiveSectionBarTransparency = Val(value)
                Case "seenBarTransparency": gblTransparencyIndicatorSeen = Val(value)
                Case "currentBarTransparency": gblTransparencyIndicatorActive = Val(value)
                Case "unseenBarTransparency": gblTransparencyIndicatorUnseen = Val(value)
                Case "currentTransparency": gblSectionBandTransparencyCurrent = Val(value)
                Case "leftTransparency": gblSectionBandTransparencyLeft = Val(value)
                Case "rightTransparency": gblSectionBandTransparencyRight = Val(value)
                Case "contrast": gblSectionBandContrast = Val(value)
                Case "gradient": gblSectionBandGradient = Val(value)
                Case "accentMode": gblAccentMode = Val(value)
                Case "accentWidth": gblAccentWidth = Val(value)
                Case "numbering": gblSectionNumbering = Val(value)
                Case "currentBold": gblCurrentTitleBold = Val(value)
                Case "futureOutline": gblFutureBarOutline = Val(value)
                Case "sectionBars": gblSectionForegroundBars = Val(value)
                Case "seenBackground"
                    gblstrFontBGColorSeen = PI_StyleColor(value)
                Case "currentBackground"
                    gblstrFontBGColorCurrent = PI_StyleColor(value)
                Case "unseenBackground"
                    gblstrFontBGColorUnSeen = PI_StyleColor(value)
                Case "seenForeground"
                    gblstrFontColorSeen = PI_StyleColor(value)
                    gblstrIndicatorColorSeen = PI_StyleColor(value)
                Case "currentForeground"
                    gblstrFontColorCurrent = PI_StyleColor(value)
                    gblstrIndicatorColorCurrent = PI_StyleColor(value)
                Case "unseenForeground"
                    gblstrFontColorUnSeen = PI_StyleColor(value)
                    gblstrIndicatorColorUnSeen = PI_StyleColor(value)
                Case "leftNear"
                    gblSectionBandLeftNear = PI_StyleColor(value)
                Case "leftFar"
                    gblSectionBandLeftFar = PI_StyleColor(value)
                Case "rightNear"
                    gblSectionBandRightNear = PI_StyleColor(value)
                Case "rightFar"
                    gblSectionBandRightFar = PI_StyleColor(value)
                Case "accentColor"
                    gblAccentColor = PI_StyleColor(value)
            End Select
        End If
    Next item
End Sub

Private Function PI_StyleColor(ByVal value As String) As String
    If Len(value) = 7 And Left$(value, 1) = "#" Then
        PI_StyleColor = CStr(CLng("&H" & Mid$(value, 2, 2))) & ";" & _
            CStr(CLng("&H" & Mid$(value, 4, 2))) & ";" & CStr(CLng("&H" & Mid$(value, 6, 2)))
    Else
        PI_StyleColor = value
    End If
End Function

Private Sub PI_SetFullWidth(ByVal pageWidth As Double)
    gblintMarginLeft = 0
    gblintMarginRight = 0
    gblNavigationWidth = 0
    gblRailWidth = pageWidth
End Sub

Public Sub PI_FullWidth(control As IRibbonControl)
    If gblbIndicatorUpdateRunning Then Exit Sub
    Dim opts As ProgressDotsOptions
    If Application.Presentations.Count = 0 Then Exit Sub
    opts = DeserializeProgressOptions(SerializeProgressOptionsRibbon())
    PI_SetFullWidth PI_RenderAxisLength(opts, ActivePresentation)
    PersistPISettings SerializeProgressOptionsRibbon()
    RefreshRibbon
    If Not gobjRibbon Is Nothing Then gobjRibbon.Invalidate
    RefreshProgressBar
End Sub

Public Sub PI_AccentGetIndex(control As IRibbonControl, ByRef index)
    index = gblAccentMode
End Sub

Public Sub PI_AccentSetIndex(control As IRibbonControl, selectedID As String, selectedIndex As Integer)
    If gblbIndicatorUpdateRunning Then Exit Sub
    If Application.Presentations.Count = 0 Then Exit Sub
    gblAccentMode = CInt(PI_BarClamp(selectedIndex, 0, 4))
    PersistPISettings SerializeProgressOptionsRibbon()
    RefreshRibbon
    If Not gobjRibbon Is Nothing Then gobjRibbon.Invalidate
    RefreshProgressBar
End Sub

Private Sub PI_AddSectionAccent(ByVal targetSlide As Slide, opts As ProgressDotsOptions, _
                                ByVal x As Double, ByVal width As Double, ByVal height As Double, _
                                ByVal sectionIndex As Long, ByVal shapeTag As String, ByVal linkTarget As String)
    Dim accent As Shape, path As FreeformBuilder, weight As Double, y As Double, inset As Double
    Dim failureDescription As String
    If opts.AccentMode < 1 Or opts.AccentMode > 4 Then Exit Sub
    If width <= 0 Or height <= 0 Then Exit Sub
    On Error GoTo AccentError
    If opts.LayeredNavigation <> 1 And opts.HardShadowOffset > 0 Then
        width = width - PI_BarClamp(opts.HardShadowOffset, 0, width / 3)
        height = height - PI_BarClamp(opts.HardShadowOffset, 0, height / 3)
    End If
    weight = PI_BarClamp(opts.AccentWidth, 0.5, 4)
    If weight > width / 3 Then weight = width / 3
    If weight > height / 3 Then weight = height / 3
    y = opts.yMargin
    Select Case opts.AccentMode
        Case 1
            Set accent = targetSlide.Shapes.AddShape(msoShapeRectangle, x, y + height - weight, width, weight)
        Case 2
            Set accent = targetSlide.Shapes.AddShape(msoShapeRectangle, x, y, width, weight)
        Case 4
            Set accent = targetSlide.Shapes.AddShape(msoShapeRectangle, x, y, weight, height)
        Case 3
            inset = weight / 2
            Set path = targetSlide.Shapes.BuildFreeform(msoEditingCorner, x + inset, y + height - inset)
            path.AddNodes msoSegmentLine, msoEditingAuto, x + inset, y + inset
            path.AddNodes msoSegmentLine, msoEditingAuto, x + width - inset, y + inset
            path.AddNodes msoSegmentLine, msoEditingAuto, x + width - inset, y + height - inset
            If opts.LayeredNavigation <> 1 Then
                path.AddNodes msoSegmentLine, msoEditingAuto, x + inset, y + height - inset
            End If
            Set accent = path.ConvertToShape
    End Select
    accent.Tags.Add shapeTag, "yes"
    accent.Tags.Add "PI_ROLE", "section-accent"
    accent.Tags.Add "PI_SECTION", CStr(sectionIndex)
    PI_ClearSurfaceEffects accent
    If opts.AccentMode = 3 Then
        accent.Fill.Visible = msoFalse
        accent.Line.Visible = msoTrue
        accent.Line.ForeColor.RGB = ColorToRGB(opts.AccentColor)
        accent.Line.Transparency = 0
        accent.Line.Weight = weight
        accent.Line.Style = msoLineSingle
        accent.Line.DashStyle = msoLineSolid
    Else
        accent.Fill.Solid
        accent.Fill.Visible = msoTrue
        accent.Fill.ForeColor.RGB = ColorToRGB(opts.AccentColor)
        accent.Fill.Transparency = 0
    End If
    accent.ActionSettings(ppMouseClick).Action = ppActionHyperlink
    accent.ActionSettings(ppMouseClick).Hyperlink.SubAddress = linkTarget
    Exit Sub
AccentError:
    failureDescription = Err.Description
    On Error GoTo 0
    Err.Raise vbObjectError + 3023, "PI_AddSectionAccent", failureDescription
End Sub

Private Sub PI_OutlineFutureBar(ByVal target As Shape)
    Dim weight As Double, alpha As Double, color As Long, failureDescription As String
    On Error GoTo OutlineError
    weight = 0.75
    If weight > target.Height / 3 Then weight = target.Height / 3
    If weight > target.Width / 3 Then weight = target.Width / 3
    alpha = target.Fill.Transparency
    color = target.Fill.ForeColor.RGB
    target.LockAspectRatio = msoFalse
    target.Left = target.Left + weight / 2
    target.Top = target.Top + weight / 2
    target.Width = target.Width - weight
    target.Height = target.Height - weight
    ' A transparent fill keeps the hollow marker's entire interior clickable.
    target.Fill.Solid
    target.Fill.Visible = msoTrue
    target.Fill.Transparency = 1
    target.Line.Visible = msoTrue
    target.Line.ForeColor.RGB = color
    target.Line.Transparency = alpha
    target.Line.Weight = weight
    target.Line.Style = msoLineSingle
    target.Line.DashStyle = msoLineSolid
    target.Tags.Add "PI_OUTLINED", "yes"
    Exit Sub
OutlineError:
    failureDescription = Err.Description
    On Error GoTo 0
    Err.Raise vbObjectError + 3023, "PI_OutlineFutureBar", failureDescription
End Sub

Public Sub PI_LayerSetValue(control As IRibbonControl, text As String)
    If gblbIndicatorUpdateRunning Then Exit Sub
    Dim value As Double, minimum As Double, maximum As Double
    If Application.Presentations.Count = 0 Then Exit Sub
    minimum = 0
    Select Case control.id
        Case "edtNavigationWidth", "edtRailWidth": maximum = 4032
        Case "edtSectionBlockHeight": maximum = 200
        Case "edtRailThickness": minimum = 0.5: maximum = 30
        Case "edtLayerShadowBlur", "edtLayerShadowDistance": maximum = 20
        Case "edtLayerShadowTransparency": maximum = 100
        Case "edtLayerDepthStep": maximum = 5
        Case "edtAccentWidth": minimum = 0.5: maximum = 4
        Case Else: Exit Sub
    End Select
    If Not IsNumeric(text) Then GoTo InvalidValue
    value = CDbl(text)
    If value < minimum Or value > maximum Then GoTo InvalidValue
    Select Case control.id
        Case "edtNavigationWidth": gblNavigationWidth = value
        Case "edtSectionBlockHeight": PI_SetManualBandHeight value
        Case "edtRailWidth": gblRailWidth = value
        Case "edtRailThickness": gblRailThickness = value
        Case "edtLayerShadowBlur": gblLayerShadowBlur = value
        Case "edtLayerShadowTransparency": gblLayerShadowTransparency = value
        Case "edtLayerDepthStep": gblLayerDepthStep = value
        Case "edtLayerShadowDistance": gblLayerShadowDistance = value
        Case "edtAccentWidth": gblAccentWidth = value
    End Select
    PersistPISettings SerializeProgressOptionsRibbon()
    RefreshRibbon
    If Not gobjRibbon Is Nothing Then gobjRibbon.Invalidate
    RefreshProgressBar
    Exit Sub
InvalidValue:
    MsgBox PI_Zh("8BF78F9351650020") & CStr(minimum) & PI_Zh("002081F30020") & CStr(maximum) & _
        PI_Zh("00204E4B95F476846570503C3002"), vbExclamation, PI_Zh("5C4253E05BFC822A")
End Sub

Public Sub PI_NavigationGeometry(opts As ProgressDotsOptions, ByVal pageWidth As Double, _
                                ByRef navigationLeft As Double, ByRef navigationWidth As Double, _
                                ByRef railLeft As Double, ByRef railWidth As Double)
    navigationLeft = PI_BarClamp(opts.LeftMargin, 0, pageWidth)
    If opts.NavigationWidth > 0 Then
        navigationWidth = opts.NavigationWidth
    Else
        navigationWidth = pageWidth - navigationLeft - opts.RightMargin
    End If
    navigationWidth = PI_BarClamp(navigationWidth, 0, pageWidth - navigationLeft)
    railWidth = opts.RailWidth
    If railWidth < navigationWidth Then railWidth = navigationWidth
    If railWidth > pageWidth Then railWidth = pageWidth
    railLeft = navigationLeft + (navigationWidth - railWidth) / 2
    railLeft = PI_BarClamp(railLeft, 0, pageWidth - railWidth)
End Sub

Private Function PI_LayerDepth(sections As ProgressDotsSectionInfo, _
                               ByVal sectionIndex As Long, ByVal currentSlide As Long) As Long
    Dim index As Long
    Dim targetOrdinal As Long, currentOrdinal As Long
    For index = 1 To sections.lngSecSections
        If sections.intArySecVisible(index) = 1 Then
            If index <= sectionIndex Then targetOrdinal = targetOrdinal + 1
            If sections.lngArySecStart(index) <= currentSlide Then currentOrdinal = currentOrdinal + 1
        End If
    Next index
    PI_LayerDepth = Abs(targetOrdinal - currentOrdinal)
End Function

Private Sub PI_StyleLayerBand(ByVal targetSlide As Slide, ByRef band As Shape, _
                             opts As ProgressDotsOptions, sections As ProgressDotsSectionInfo, _
                             ByVal state As Integer, ByVal sectionIndex As Long, ByVal fullHeight As Double, _
                             ByVal railLeft As Double, ByVal railWidth As Double, _
                             ByVal shapeTag As String, ByVal linkTarget As String)
    Dim foreground As Shape, path As FreeformBuilder
    Dim leftEdge As Double, rightEdge As Double, railTop As Double, railBottom As Double
    Dim depth As Long, shadowDistance As Double
    Dim failureDescription As String
    On Error GoTo LayerError
110     leftEdge = band.Left
120     rightEdge = band.Left + band.Width
130     railTop = opts.yMargin + fullHeight
140     railBottom = railTop + PI_BarClamp(opts.RailThickness, 0.5, 30)
150     shadowDistance = PI_BarClamp(opts.LayerShadowDistance, 0, 20)
160     If state = 1 Then
        ' One outline prevents doubled opacity and an internal shadow at the joint.
180         Set path = targetSlide.Shapes.BuildFreeform(msoEditingCorner, leftEdge, opts.yMargin)
190         path.AddNodes msoSegmentLine, msoEditingAuto, rightEdge, opts.yMargin
200         path.AddNodes msoSegmentLine, msoEditingAuto, rightEdge, railTop
210         If railLeft + railWidth > rightEdge Then path.AddNodes msoSegmentLine, msoEditingAuto, railLeft + railWidth, railTop
220         path.AddNodes msoSegmentLine, msoEditingAuto, railLeft + railWidth, railBottom
230         path.AddNodes msoSegmentLine, msoEditingAuto, railLeft, railBottom
240         If railLeft < leftEdge Then path.AddNodes msoSegmentLine, msoEditingAuto, railLeft, railTop
250         path.AddNodes msoSegmentLine, msoEditingAuto, leftEdge, railTop
260         path.AddNodes msoSegmentLine, msoEditingAuto, leftEdge, opts.yMargin
270         Set foreground = path.ConvertToShape
280         foreground.Tags.Add shapeTag, "yes"
290         foreground.Tags.Add "PI_ROLE", "layer-foreground"
300         foreground.Tags.Add "PI_SECTION", CStr(sectionIndex)
            foreground.Tags.Add "PI_DEPTH", "0"
310         foreground.ActionSettings(ppMouseClick).Action = ppActionHyperlink
320         foreground.ActionSettings(ppMouseClick).Hyperlink.SubAddress = linkTarget
330         foreground.Name = "PI current chapter and rail"
340         PI_PaintSectionBackground foreground, opts, 1, 0
350         PI_ApplyLayerShadow foreground, opts, -shadowDistance
360         band.Delete
370         Set band = foreground
    Else
390         depth = PI_LayerDepth(sections, sectionIndex, gblLastPIslideNum)
400         band.Height = fullHeight
410         band.Tags.Add "PI_DEPTH", CStr(depth)
            ' Offset has three settings; stacking uses every visible section's exact depth.
420         PI_ApplyLayerShadow band, opts, shadowDistance + PI_BarClamp(depth - 1, 0, 2) * PI_BarClamp(opts.LayerDepthStep, 0, 5)
    End If
    Exit Sub
LayerError:
    failureDescription = Err.Description
    PI_FirstError "PI_StyleLayerBand", Erl, Err.Number, Err.Source, Err.Description
    On Error GoTo 0
    Err.Raise vbObjectError + 3023, "PI_StyleLayerBand", failureDescription
End Sub

Private Sub PI_ApplyLayerShadow(ByVal target As Shape, opts As ProgressDotsOptions, ByVal offsetY As Double)
    Dim effect As Object, blur As Double, alpha As Double
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo ShadowError
50     Set effect = target.Shadow
60     effect.Visible = msoFalse
70     blur = PI_BarClamp(opts.LayerShadowBlur, 0, 20)
80     alpha = PI_BandClampPercent(opts.LayerShadowTransparency) / 100
90     If blur <= 0 Or alpha >= 1 Then Exit Sub
100     If target.Fill.Visible <> msoTrue Then Exit Sub
110     If target.Fill.Transparency >= 1 Then Exit Sub
    With effect
        ' Set the preset before overriding its geometry and appearance.
130         .Visible = msoTrue
140         .Type = msoShadow21
150         .Style = msoShadowStyleOuterShadow
160         .Size = 100
170         .RotateWithShape = msoFalse
180         .ForeColor.RGB = RGB(0, 0, 0)
190         .Blur = blur
200         .OffsetX = 0
210         .OffsetY = offsetY
220         .Transparency = alpha
    End With
    Exit Sub
ShadowError:
    failureNumber = Err.Number
    failureDescription = Err.Description
    PI_FirstError "PI_ApplyLayerShadow", Erl, Err.Number, Err.Source, Err.Description
    On Error GoTo 0
    Err.Raise vbObjectError + 3023, "PI_ApplyLayerShadow", _
        PI_Zh("539F751F591696345F718BBE7F6E59318D25FF1A") & CStr(failureNumber) & " " & failureDescription
End Sub

Private Sub PI_OrderLayeredShapes(ByVal targetSlide As Slide, ByVal firstShape As Long, _
                                  ByVal shapeTag As String, ByVal sectionCount As Long)
    Dim generated() As Shape, index As Long, count As Long, depth As Long, role As String
    Dim surfacesByDepth() As Collection, surface As Shape
    Dim failureDescription As String
    On Error GoTo LayerError
50     count = targetSlide.Shapes.Count - firstShape + 1
60     If count < 1 Then Exit Sub
70     ReDim generated(1 To count)
       ReDim surfacesByDepth(0 To sectionCount)
80     For index = 1 To count
90         Set generated(index) = targetSlide.Shapes(firstShape + index - 1)
           If UCase$(generated(index).Tags(shapeTag)) = "YES" Then
               role = LCase$(generated(index).Tags("PI_ROLE"))
               If role = "section-band" Or role = "section-backdrop" Then
                   depth = CLng(Val(generated(index).Tags("PI_DEPTH")))
                   If depth > 0 Then
                       If surfacesByDepth(depth) Is Nothing Then Set surfacesByDepth(depth) = New Collection
                       surfacesByDepth(depth).Add generated(index)
                   End If
               End If
           End If
    Next index
    ' Draw the farther panels first so shadows from nearer panels fall across them.
120     For depth = sectionCount To 1 Step -1
            If Not surfacesByDepth(depth) Is Nothing Then
                For Each surface In surfacesByDepth(depth)
170                 surface.ZOrder msoBringToFront
                Next surface
            End If
    Next depth
220     For index = 1 To count
230         If LCase$(generated(index).Tags("PI_ROLE")) = "layer-foreground" Then generated(index).ZOrder msoBringToFront
    Next index
250     For index = 1 To count
260         role = LCase$(generated(index).Tags("PI_ROLE"))
270         If role = "section-hit-area" Or role = "layer-hit-area" Or role = "section-body" Or _
                role = "section-marker-row" Then generated(index).ZOrder msoBringToFront
    Next index
290     For index = 1 To count
300         If LCase$(generated(index).Tags("PI_ROLE")) = "section-accent" Then generated(index).ZOrder msoBringToFront
    Next index
320     For index = 1 To count
330         role = LCase$(generated(index).Tags("PI_ROLE"))
340         If role = "section-rule" Then generated(index).ZOrder msoBringToFront
    Next index
360     For index = 1 To count
370         role = LCase$(generated(index).Tags("PI_ROLE"))
380         If role = "section-title" Or role = "section-counter" Or role = "indicator" Or _
            role = "indicator-summary" Then generated(index).ZOrder msoBringToFront
    Next index
    Exit Sub
LayerError:
    failureDescription = Err.Description
    PI_FirstError "PI_OrderLayeredShapes", Erl, Err.Number, Err.Source, Err.Description
    On Error GoTo 0
    Err.Raise vbObjectError + 3023, "PI_OrderLayeredShapes", failureDescription
End Sub

Private Sub PI_ClipLayeredBackgrounds(ByVal targetSlide As Slide, ByVal firstShape As Long, _
                                     ByVal shapeTag As String, ByVal railBottom As Double)
    Dim original() As Shape, index As Long, matchIndex As Long, count As Long
    Dim role As String, backdrop As Shape, failureDescription As String, anyShadow As Boolean
    On Error GoTo ClipError
60     count = targetSlide.Shapes.Count - firstShape + 1
70     If count < 1 Then Exit Sub
80     ReDim original(1 To count)
90     For index = 1 To count
100         Set original(index) = targetSlide.Shapes(firstShape + index - 1)
110         role = LCase$(original(index).Tags("PI_ROLE"))
120         If role = "section-band" Or role = "layer-foreground" Then
130             If original(index).Shadow.Visible = msoTrue Then anyShadow = True
        End If
    Next index
160     If Not anyShadow Then Exit Sub
170     For index = 1 To count
180         role = LCase$(original(index).Tags("PI_ROLE"))
190         If role = "section-band" Or role = "layer-foreground" Then
200             Set backdrop = Nothing
210             For matchIndex = 1 To count
220                 If LCase$(original(matchIndex).Tags("PI_ROLE")) = "section-backdrop" Then
230                     If original(matchIndex).Tags("PI_SECTION") = original(index).Tags("PI_SECTION") Then
240                         Set backdrop = original(matchIndex)
                        Exit For
                    End If
                End If
            Next matchIndex
290             If original(index).Shadow.Visible = msoTrue Then
300                 PI_CaptureLayerSurface targetSlide, original(index), backdrop, shapeTag, railBottom
            Else
320                 PI_AddNativeHitRegion original(index), shapeTag
            End If
        End If
    Next index
360     For index = 1 To count
370         If LCase$(original(index).Tags("PI_ROLE")) = "section-backdrop" Then
380             If UCase$(original(index).Tags("PI_CAPTURED")) = "YES" Then original(index).Delete
        End If
    Next index
    Exit Sub
ClipError:
    failureDescription = Err.Description
    PI_FirstError "PI_ClipLayeredBackgrounds", Erl, Err.Number, Err.Source, Err.Description
    On Error GoTo 0
    Err.Raise vbObjectError + 3023, "PI_ClipLayeredBackgrounds", failureDescription
End Sub

Private Function PI_SurfaceFile(Optional ByVal leaf As String = "surface.png") As String
    Dim root As String, separator As String, runDirectory As String
#If Mac Then
    separator = "/"
    root = Environ$("TMPDIR")
#Else
    separator = "\"
    root = Environ$("TEMP")
#End If
    If Len(mPISurfaceDirectory) = 0 Then
        If Len(mPIScratchRoot) > 0 Then root = mPIScratchRoot
        If Right$(root, 1) <> separator Then root = root & separator
        mPIRunSerial = mPIRunSerial + 1
        runDirectory = root & "PI_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & _
            CStr(CLng(Timer * 100)) & "_" & CStr(mPIRunSerial)
        MkDir runDirectory
        mPISurfaceDirectory = runDirectory
    End If
    PI_SurfaceFile = mPISurfaceDirectory & separator & leaf
End Function

Private Function PI_EmbedShapePNG(ByVal targetSlide As Slide, ByVal source As Shape, _
                                  Optional ByVal cacheSection As Long = 0) As Shape
    Dim imagePath As String, picture As Shape
    Dim failureNumber As Long, failureSource As String, failureDescription As String
    On Error GoTo ExportError
    If cacheSection > 0 Then
10      imagePath = PI_SurfaceFile("section_" & CStr(cacheSection) & ".png")
    Else
        imagePath = PI_SurfaceFile()
    End If
    If Len(mPITraceFile) > 0 Then PI_Trace "png.export"
20  source.Export imagePath, ppShapeFormatPNG
    If Len(mPITraceFile) > 0 Then PI_Trace "png.embed"
30  Set picture = targetSlide.Shapes.AddPicture(imagePath, msoFalse, msoTrue, 0, 0, -1, -1)
    If cacheSection > 0 Then
        mPISectionSlots(cacheSection).surfacePath = imagePath
    Else
40      Kill imagePath
    End If
    Set PI_EmbedShapePNG = picture
    Exit Function
ExportError:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    PI_FirstError "PI_EmbedShapePNG", Erl, failureNumber, failureSource, failureDescription
    On Error Resume Next
    If Not picture Is Nothing Then picture.Delete
    If Len(imagePath) > 0 Then
        If Len(Dir$(imagePath)) > 0 Then Kill imagePath
    End If
    On Error GoTo 0
    Err.Raise failureNumber, failureSource, failureDescription
End Function

Private Sub PI_CaptureLayerSurface(ByVal targetSlide As Slide, ByVal source As Shape, _
                                   ByVal backdrop As Shape, ByVal shapeTag As String, ByVal railBottom As Double)
    Const captureScale As Double = 4
    Dim bodyCopy As Shape, backdropCopy As Shape, frame As Shape, captureGroup As Shape, picture As Shape
    Dim duplicates As ShapeRange, captureMembers As ShapeRange, shadowEffect As Object
    Dim shadowVisible As Boolean, surfaceVisible As Boolean
    Dim shadowBlur As Double, shadowOffsetX As Double, shadowOffsetY As Double
    Dim shadowAlpha As Double, shadowSize As Double, shadowColor As Long
    Dim guard As Double, maxOffset As Double, ex As Double, ey As Double, ew As Double, eh As Double
    Dim fx As Double, fy As Double, fw As Double, fh As Double, rightEdge As Double, bottomEdge As Double
    Dim expectedAspect As Double, pastedAspect As Double, tolerance As Double
    Dim cacheSection As Long
    Dim role As String, failureNumber As Long, failureDescription As String
    On Error GoTo CaptureError
140     role = LCase$(source.Tags("PI_ROLE"))
    If Len(mPITraceFile) > 0 Then PI_Trace "shadow.begin", CLng(source.Tags("PI_SECTION"))
150     surfaceVisible = False
160     If source.Fill.Visible = msoTrue Then surfaceVisible = (source.Fill.Transparency < 1)
170     If Not backdrop Is Nothing Then surfaceVisible = True
180     Set shadowEffect = source.Shadow
190     shadowVisible = (shadowEffect.Visible = msoTrue)
200     If Not shadowVisible Then Exit Sub
210     If shadowVisible Then
220         shadowBlur = shadowEffect.Blur
230         shadowOffsetX = shadowEffect.OffsetX
240         shadowOffsetY = shadowEffect.OffsetY
250         shadowAlpha = shadowEffect.Transparency
260         shadowSize = shadowEffect.Size
270         shadowColor = shadowEffect.ForeColor.RGB
    End If
290     maxOffset = Abs(shadowOffsetX)
300     If Abs(shadowOffsetY) > maxOffset Then maxOffset = Abs(shadowOffsetY)
310     guard = 4 * shadowBlur + maxOffset + 2
320     ex = source.Left - guard
330     ey = source.Top - guard
340     ew = source.Width + 2 * guard
350     bottomEdge = source.Top + source.Height
360     If railBottom > bottomEdge Then bottomEdge = railBottom
370     eh = bottomEdge + guard - ey
380     fx = ex
390     If fx < 0 Then fx = 0
400     fy = ey
410     If fy < 0 Then fy = 0
420     rightEdge = ex + ew
430     If rightEdge > mPIRenderCanvasWidth Then rightEdge = mPIRenderCanvasWidth
440     bottomEdge = railBottom
450     If bottomEdge > mPIRenderCanvasHeight Then bottomEdge = mPIRenderCanvasHeight
460     fw = rightEdge - fx
470     fh = bottomEdge - fy
480     If Not surfaceVisible Or fw <= 0 Or fh <= 0 Then GoTo MakeHitRegion

    If mPIStoreLayout Then cacheSection = CLng(source.Tags("PI_SECTION"))
    If mPIReadLayout Then
    If Len(mPITraceFile) > 0 Then PI_Trace "shadow.embed.cached", cacheSection
        Set picture = targetSlide.Shapes.AddPicture(mPISectionSlots(cacheSection).surfacePath, msoFalse, msoTrue, 0, 0, -1, -1)
        picture.Tags.Add shapeTag, "yes"
        mPISurfaceHits = mPISurfaceHits + 1
        GoTo CropViewport
    End If

    ' Rebase and enlarge copies; keep the original geometry and shadow values untouched.
    mPIStage = "capture section " & source.Tags("PI_SECTION") & " " & role
    mPIParameters = "surface=" & source.Width & "x" & source.Height & " expanded=" & ew & "x" & eh & _
        " viewport=" & fx & "," & fy & "," & fw & "," & fh & " blur=" & shadowBlur & _
        " offset=" & shadowOffsetX & "," & shadowOffsetY & " alpha=" & shadowAlpha & " scale=" & captureScale
510     Set duplicates = source.Duplicate
520     Set bodyCopy = duplicates.Item(1)
530     bodyCopy.Tags.Add shapeTag, "yes"
540     bodyCopy.Name = PI_CaptureTemporaryName(targetSlide, "body", bodyCopy.Id)
550     PI_ClearSurfaceEffects bodyCopy
560     bodyCopy.LockAspectRatio = msoFalse
570     bodyCopy.Width = source.Width * captureScale
580     bodyCopy.Height = source.Height * captureScale
590     bodyCopy.Left = (source.Left - ex) * captureScale
600     bodyCopy.Top = (source.Top - ey) * captureScale
610     Set shadowEffect = bodyCopy.Shadow
620     If shadowVisible Then
625         shadowEffect.Visible = msoTrue
630         shadowEffect.Type = msoShadow21
640         shadowEffect.Style = msoShadowStyleOuterShadow
650         shadowEffect.Size = shadowSize
660         shadowEffect.RotateWithShape = msoFalse
670         shadowEffect.Blur = shadowBlur * captureScale
680         shadowEffect.OffsetX = shadowOffsetX * captureScale
690         shadowEffect.OffsetY = shadowOffsetY * captureScale
700         shadowEffect.Transparency = shadowAlpha
710         shadowEffect.ForeColor.RGB = shadowColor
    Else
740         shadowEffect.Visible = msoFalse
    End If
760     If Not backdrop Is Nothing Then
770         Set duplicates = backdrop.Duplicate
780         Set backdropCopy = duplicates.Item(1)
790         backdropCopy.Tags.Add shapeTag, "yes"
800         backdropCopy.Name = PI_CaptureTemporaryName(targetSlide, "backdrop", backdropCopy.Id)
810         backdropCopy.LockAspectRatio = msoFalse
820         backdropCopy.Width = backdrop.Width * captureScale
830         backdropCopy.Height = backdrop.Height * captureScale
840         backdropCopy.Left = (backdrop.Left - ex) * captureScale
850         backdropCopy.Top = (backdrop.Top - ey) * captureScale
860         PI_ClearSurfaceEffects backdropCopy
    End If
880     Set frame = targetSlide.Shapes.AddShape(msoShapeRectangle, 0, 0, ew * captureScale, eh * captureScale)
890     frame.Tags.Add shapeTag, "yes"
900     frame.Name = PI_CaptureTemporaryName(targetSlide, "frame", frame.Id)
910     frame.Fill.Solid
920     frame.Fill.Visible = msoTrue
930     frame.Fill.Transparency = 1
940     PI_ClearSurfaceEffects frame
950     If backdropCopy Is Nothing Then
960         Set captureMembers = targetSlide.Shapes.Range(Array(bodyCopy.Name, frame.Name))
    Else
980         Set captureMembers = targetSlide.Shapes.Range(Array(bodyCopy.Name, backdropCopy.Name, frame.Name))
    End If
1000     PI_VerifyCaptureMembers captureMembers, bodyCopy, backdropCopy, frame
    If Len(mPITraceFile) > 0 Then PI_Trace "shadow.group"
1010     Set captureGroup = captureMembers.Group
1020     Set captureMembers = Nothing
1030     Set bodyCopy = Nothing
1040     Set backdropCopy = Nothing
1050     Set frame = Nothing
1060     captureGroup.Tags.Add shapeTag, "yes"
1070     tolerance = 0.1
1080     If ew * captureScale / 100000 > tolerance Then tolerance = ew * captureScale / 100000
1090     If Abs(captureGroup.Left) > tolerance Or Abs(captureGroup.Top) > tolerance Or _
        Abs(captureGroup.Width - ew * captureScale) > tolerance Or _
        Abs(captureGroup.Height - eh * captureScale) > tolerance Then
1120         Err.Raise vbObjectError + 3023, "PI_CaptureLayerSurface", PI_Zh("635583B78FB9754C4E0D5339914D3002")
    End If
1140     PI_YieldRender
1160     Set picture = PI_EmbedShapePNG(targetSlide, captureGroup, cacheSection)
1210     picture.Tags.Add shapeTag, "yes"
1220     If picture.Width <= 0 Or picture.Height <= 0 Then
1230         Err.Raise vbObjectError + 3023, "PI_CaptureLayerSurface", PI_Zh("5BFC822A80CC666F00200050004E004700205C3A5BF865E065483002")
    End If
1250     expectedAspect = ew / eh
1260     pastedAspect = picture.Width / picture.Height
1270     If Abs(pastedAspect / expectedAspect - 1) > 0.01 Then
1280         Err.Raise vbObjectError + 3023, "PI_CaptureLayerSurface", PI_Zh("5BFC822A80CC666F00200050004E004700206BD44F8B4E0D5339914D3002")
    End If
    If Len(mPITraceFile) > 0 Then PI_Trace "shadow.delete.group"
1300     captureGroup.Delete
1310     Set captureGroup = Nothing
1320     Set bodyCopy = Nothing
1330     Set backdropCopy = Nothing
1340     Set frame = Nothing

    If mPIRunActive Then mPICaptures = mPICaptures + 1
CropViewport:
    ' The viewport clips the baked silhouette; no second fill or native shadow is added.
    If Len(mPITraceFile) > 0 Then PI_Trace "shadow.crop"
1370     picture.LockAspectRatio = msoFalse
    With picture.PictureFormat.Crop
1390         .PictureWidth = ew
1400         .PictureHeight = eh
1410         .ShapeWidth = fw
1420         .ShapeHeight = fh
1430         .ShapeLeft = fx
1440         .ShapeTop = fy
1450         .PictureOffsetX = (ex + ew / 2) - (fx + fw / 2)
1460         .PictureOffsetY = (ey + eh / 2) - (fy + fh / 2)
    End With
1480     PI_ClearSurfaceEffects picture
1490     picture.Tags.Add "PI_ROLE", role
1500     picture.Tags.Add "PI_SECTION", source.Tags("PI_SECTION")
        If Len(source.Tags("PI_BODY_PT")) > 0 Then
            picture.Tags.Add "PI_BODY_PT", source.Tags("PI_BODY_PT")
            picture.Tags.Add "PI_ROW_PT", source.Tags("PI_ROW_PT")
        End If
1510     If Len(source.Tags("PI_DEPTH")) > 0 Then picture.Tags.Add "PI_DEPTH", source.Tags("PI_DEPTH")
1520     picture.Tags.Add "PI_RENDER", "clipped-surface"
1540     picture.Name = "PI clipped surface " & source.Tags("PI_SECTION")
1550     picture.ActionSettings(ppMouseClick).Action = ppActionNone
1560     picture.ActionSettings(ppMouseOver).Action = ppActionNone
1570     If Not backdrop Is Nothing Then backdrop.Tags.Add "PI_CAPTURED", "yes"

MakeHitRegion:
    If Len(mPITraceFile) > 0 Then PI_Trace "shadow.hitregion"
1600     source.Fill.Solid
1610     source.Fill.Visible = msoTrue
1620     source.Fill.Transparency = 1
1630     PI_ClearSurfaceEffects source
1640     If role = "layer-foreground" Then
1650         source.Tags.Add "PI_ROLE", "layer-hit-area"
    Else
1670         source.Tags.Add "PI_ROLE", "section-hit-area"
    End If
    Exit Sub
CaptureError:
    failureNumber = Err.Number
    failureDescription = Err.Description
    PI_FirstError "PI_CaptureLayerSurface", Erl, Err.Number, Err.Source, Err.Description
    On Error Resume Next
    If Not picture Is Nothing Then picture.Delete
    If Not captureGroup Is Nothing Then captureGroup.Delete
    If Not frame Is Nothing Then frame.Delete
    If Not backdropCopy Is Nothing Then backdropCopy.Delete
    If Not bodyCopy Is Nothing Then bodyCopy.Delete
    On Error GoTo 0
    Err.Raise vbObjectError + 3023, "PI_CaptureLayerSurface", _
        PI_Zh("5BFC822A80CC666F88C1520759318D25FF1A") & CStr(failureNumber) & " " & failureDescription
End Sub

Private Function PI_CaptureTemporaryName(ByVal targetSlide As Slide, ByVal role As String, _
                                          ByVal shapeId As Long) As String
    Dim capturePrefix As String, candidate As String, index As Long, suffix As Long, collision As Boolean
    capturePrefix = "PI_CAPTURE_" & role & "_" & CStr(shapeId)
    candidate = capturePrefix
    Do
        collision = False
        For index = 1 To targetSlide.Shapes.Count
            If StrComp(targetSlide.Shapes(index).Name, candidate, vbTextCompare) = 0 Then
                collision = True
                Exit For
            End If
        Next index
        If Not collision Then Exit Do
        suffix = suffix + 1
        candidate = capturePrefix & "_" & CStr(suffix)
    Loop
    PI_CaptureTemporaryName = candidate
End Function

Private Sub PI_VerifyCaptureMembers(ByVal members As ShapeRange, ByVal body As Shape, _
                                     ByVal backdrop As Shape, ByVal frame As Shape)
    Dim index As Long, expectedCount As Long, memberId As Long
    Dim expectedMask As Long, seenMask As Long, memberMask As Long, backdropId As Long
    expectedCount = 2
    expectedMask = 3
    backdropId = -1
    If Not backdrop Is Nothing Then
        expectedCount = 3
        expectedMask = 7
        backdropId = backdrop.Id
    End If
    If members.Count <> expectedCount Then GoTo InvalidMembers
    For index = 1 To members.Count
        memberId = members.Item(index).Id
        If memberId = body.Id Then
            memberMask = 1
        ElseIf memberId = frame.Id Then
            memberMask = 2
        ElseIf memberId = backdropId Then
            memberMask = 4
        Else
            GoTo InvalidMembers
        End If
        If (seenMask And memberMask) <> 0 Then GoTo InvalidMembers
        seenMask = seenMask Or memberMask
    Next index
    If seenMask <> expectedMask Then GoTo InvalidMembers
    Exit Sub
InvalidMembers:
    Err.Raise vbObjectError + 3023, "PI_VerifyCaptureMembers", PI_Zh("5BFC822A635583B75BF98C614E0D5339914D3002")
End Sub

Private Sub PI_AddNativeHitRegion(ByVal source As Shape, ByVal shapeTag As String)
    Dim duplicates As ShapeRange, hit As Shape, role As String, failureDescription As String
    On Error GoTo HitError
40     role = LCase$(source.Tags("PI_ROLE"))
50     Set duplicates = source.Duplicate
60     Set hit = duplicates.Item(1)
70     hit.Tags.Add shapeTag, "yes"
80     hit.Left = source.Left
90     hit.Top = source.Top
100     hit.Fill.Solid
110     hit.Fill.Visible = msoTrue
120     hit.Fill.Transparency = 1
130     PI_ClearSurfaceEffects hit
140     If role = "layer-foreground" Then
150         hit.Tags.Add "PI_ROLE", "layer-hit-area"
    Else
170         hit.Tags.Add "PI_ROLE", "section-hit-area"
    End If
    Exit Sub
HitError:
    failureDescription = Err.Description
    PI_FirstError "PI_AddNativeHitRegion", Erl, Err.Number, Err.Source, Err.Description
    On Error Resume Next
    If Not hit Is Nothing Then hit.Delete
    On Error GoTo 0
    Err.Raise vbObjectError + 3023, "PI_AddNativeHitRegion", failureDescription
End Sub

Private Sub PI_ClearSurfaceEffects(ByVal target As Shape)
    Dim effect As Object
    On Error GoTo SurfaceError
40     target.Line.Visible = msoFalse
50     target.Shadow.Visible = msoFalse
60     Set effect = target.Glow
70     effect.Radius = 0
80     effect.Transparency = 1
90     Set effect = target.SoftEdge
100     effect.Radius = 0
    Exit Sub
SurfaceError:
    PI_FirstError "PI_ClearSurfaceEffects", Erl, Err.Number, Err.Source, Err.Description
    Err.Raise mPIFirstNumber, mPIFirstSource, mPIFirstDescription
End Sub

Private Sub PI_DisableLayerGroupShadow(ByVal target As Shape)
    Dim failureDescription As String
    On Error GoTo GroupError
40     target.Shadow.Visible = msoFalse
    Exit Sub
GroupError:
    failureDescription = Err.Description
    PI_FirstError "PI_DisableLayerGroupShadow", Erl, Err.Number, Err.Source, Err.Description
    On Error GoTo 0
    Err.Raise vbObjectError + 3023, "PI_DisableLayerGroupShadow", failureDescription
End Sub

Public Function PI_UsesSectionBand(opts As ProgressDotsOptions) As Boolean
    PI_UsesSectionBand = (opts.SectionBand = 1 And _
        opts.PIStyle = "Full" And opts.ignoreSections = "Full" And _
        opts.ignoreIndicator <> "None")
End Function

Public Function PI_BandTitleHeight(fontSize As Double) As Double
    PI_BandTitleHeight = fontSize * 1.35 + 2
End Function

Private Sub PI_AddCompactRegion(ByVal target As Slide, opts As ProgressDotsOptions, _
            ByVal left As Double, ByVal top As Double, ByVal width As Double, ByVal height As Double, _
            ByVal section As Long, ByVal state As Integer, ByVal position As Double, ByVal role As String, _
            ByVal shapeTag As String, ByVal subAddress As String, ByVal paint As Boolean)
    Dim region As Shape
    Set region = target.Shapes.AddShape(msoShapeRectangle, left, top, width, height)
    PI_ClearSurfaceEffects region
    If paint Then
        PI_PaintSectionBackground region, opts, state, position
    Else
        ' The single layered silhouette beneath these chapter and marker hit areas owns their fill.
        region.Fill.Solid
        region.Fill.Transparency = 1
    End If
    region.Tags.Add shapeTag, "yes"
    region.Tags.Add "PI_ROLE", role
    region.Tags.Add "PI_SECTION", CStr(section)
    region.ActionSettings(ppMouseClick).Action = ppActionHyperlink
    region.ActionSettings(ppMouseClick).Hyperlink.SubAddress = subAddress
End Sub

Private Sub PI_SetManualBandHeight(ByVal value As Double)
    gblSectionBlockHeight = value
    gblBandHeightMode = PI_EXACT_BAND
End Sub

Private Function PI_CompactBodyHeight(opts As ProgressDotsOptions) As Double
    PI_CompactBodyHeight = opts.SectionBlockHeight
    If opts.SectionBlockHeight = 0 Then PI_CompactBodyHeight = PI_BandTitleHeight(opts.textHeight)
End Function

Public Sub PI_BandGeometry(opts As ProgressDotsOptions, ByRef titleTop As Double, _
                          ByRef dotsTop As Double, ByRef bandHeight As Double, _
                          Optional ByRef presetGrowth As Double = 0, Optional ByRef physicalBody As Double = 0)
    Dim padding As Double, gap As Double, rowHeight As Double, markerHeight As Double
    Dim counterHeight As Double, hardOffset As Double, bodyHeight As Double
    Dim contentShift As Double, reverse As Boolean
    padding = opts.SectionBandPadding
    gap = opts.SectionIndicatorGap
    If gap < 0 Then gap = 0
    rowHeight = PI_BandTitleHeight(opts.textHeight)
    markerHeight = PI_MarkerHeight(opts)
    counterHeight = PI_CounterRowHeight(opts)
    presetGrowth = 0

    If opts.PresetSize > 0 And opts.DockEdge < 2 Then
        bodyHeight = PI_CompactBodyHeight(opts)
        If opts.LayeredNavigation = 0 Then hardOffset = PI_BarClamp(opts.HardShadowOffset, 0, bodyHeight / 3)
        bandHeight = bodyHeight + hardOffset + counterHeight + markerHeight
        If opts.SectionPlacement = 0 Then
            titleTop = opts.yMargin
            dotsTop = opts.yMargin + bandHeight - markerHeight
        Else
            dotsTop = opts.yMargin
            titleTop = dotsTop + markerHeight + counterHeight
        End If
        If opts.PresetSize = 1 Then titleTop = titleTop + (bodyHeight - rowHeight) / 2
        physicalBody = bodyHeight
    Else
        If opts.SectionPlacement = 0 Then
            titleTop = opts.yMargin + padding
            dotsTop = titleTop + rowHeight + counterHeight + gap
        Else
            dotsTop = opts.yMargin + padding
            titleTop = dotsTop + markerHeight + gap
        End If
        bandHeight = 2 * padding + rowHeight + counterHeight + gap + markerHeight
        If opts.DockEdge >= 2 Then
            If opts.SidebarWidth > bandHeight Then bandHeight = opts.SidebarWidth
        ElseIf opts.BandHeightMode = PI_EXACT_BAND And opts.SectionBlockHeight > 0 Then
            bandHeight = opts.SectionBlockHeight
            If opts.LayeredNavigation = 0 Then bandHeight = bandHeight + PI_BarClamp(opts.HardShadowOffset, 0, bandHeight / 2)
        ElseIf opts.SectionBlockHeight > bandHeight Then
            bandHeight = opts.SectionBlockHeight
        End If
        physicalBody = bandHeight
        If opts.LayeredNavigation = 0 Then physicalBody = physicalBody - PI_BarClamp(opts.HardShadowOffset, 0, bandHeight / 3)
    End If

    If opts.BandHeightMode <> PI_PROPORTIONAL_BAND Or opts.PresetSize = 2 Or opts.DockEdge >= 2 Then Exit Sub
    ' Grow from the chapter face, excluding its hard shadow, marker row and rail.
    presetGrowth = 0.3 * physicalBody
    physicalBody = physicalBody + presetGrowth
    bandHeight = bandHeight + presetGrowth
    reverse = (opts.SectionPlacement <> 0)
    ' Large bottom-docked text is translated with its original reading order.
    If opts.PresetSize = 0 And opts.DockEdge = 1 Then reverse = Not reverse
    contentShift = 2 * presetGrowth / 3
    If reverse Then contentShift = presetGrowth / 3
    titleTop = titleTop + contentShift
    dotsTop = dotsTop + contentShift
End Sub

Public Function PI_SectionBackground(opts As ProgressDotsOptions, ByVal state As Integer, _
                         ByRef enabled As Boolean, ByRef alpha As Double, _
                         Optional ByVal position As Double = -1) As Long
    Select Case state
        Case 0
            enabled = (opts.useTextColorBG = 1)
            alpha = opts.transparencyFontSeen / 100
            PI_SectionBackground = ColorToRGB(opts.seenTextColorBG)
        Case 1
            enabled = (opts.useTextColorBG_Current = 1)
            alpha = opts.transparencyFontActive / 100
            PI_SectionBackground = ColorToRGB(opts.CurrentTextColorBG)
        Case Else
            enabled = (opts.useTextColorBG_UnSeen = 1)
            alpha = opts.transparencyFontUnseen / 100
            PI_SectionBackground = ColorToRGB(opts.unseenTextColorBG)
    End Select
    If PI_UsesSectionBand(opts) Then
        Select Case state
            Case 0: alpha = PI_BandClampPercent(opts.SectionBandTransparencyLeft) / 100
            Case 1: alpha = PI_BandClampPercent(opts.SectionBandTransparencyCurrent) / 100
            Case Else: alpha = PI_BandClampPercent(opts.SectionBandTransparencyRight) / 100
        End Select
    End If
    If opts.SectionBandGradient = 1 And position >= 0 Then
        Select Case state
            Case 0
                PI_SectionBackground = PI_BandCompositeColor(ColorToRGB(opts.SectionBandLeftNear), position, ColorToRGB(opts.SectionBandLeftFar))
                alpha = PI_BandClampPercent(opts.SectionBandTransparencyLeft) / 100
            Case 1
                alpha = PI_BandClampPercent(opts.SectionBandTransparencyCurrent) / 100
            Case Else
                PI_SectionBackground = PI_BandCompositeColor(ColorToRGB(opts.SectionBandRightNear), position, ColorToRGB(opts.SectionBandRightFar))
                alpha = PI_BandClampPercent(opts.SectionBandTransparencyRight) / 100
        End Select
    End If
End Function

Public Sub PI_PaintSectionBackground(target As Shape, opts As ProgressDotsOptions, ByVal state As Integer, _
                                    Optional ByVal position As Double = -1)
    Dim enabled As Boolean, alpha As Double, color As Long
    color = PI_SectionBackground(opts, state, enabled, alpha, position)
    target.Line.Visible = msoFalse
    If enabled Then
        target.Fill.Solid
        target.Fill.ForeColor.RGB = color
        target.Fill.Transparency = alpha
        target.Fill.Visible = msoTrue
    Else
        target.Fill.Visible = msoFalse
    End If
End Sub

Public Function PI_BandClampPercent(ByVal value As Double) As Double
    If value < 0 Then value = 0
    If value > 100 Then value = 100
    PI_BandClampPercent = value
End Function

Public Function PI_BandClampBackdropScale(ByVal value As Double) As Double
    If value < 110 Then value = 110
    If value > 300 Then value = 300
    PI_BandClampBackdropScale = value
End Function

Public Sub PI_BandGetBackdropScale(control As IRibbonControl, ByRef value)
    value = CStr(gblSectionBandBackdropScale)
End Sub

Public Sub PI_BandSetBackdropScale(control As IRibbonControl, text As String)
    If gblbIndicatorUpdateRunning Then Exit Sub
    Dim value As Double
    If Application.Presentations.Count = 0 Then Exit Sub
    If Not IsNumeric(text) Then
        MsgBox PI_Zh("8BF78F9351650020003100310030002081F3002000330030003000204E4B95F4768465705B573002"), vbExclamation
        Exit Sub
    End If
    value = CDbl(text)
    If value < 110 Or value > 300 Then
        MsgBox PI_Zh("80CC666F5B57653E5927830356F44E3A0020003100310030002081F3002000330030003000253002"), vbExclamation
        Exit Sub
    End If
    gblSectionBandBackdropScale = value
    PersistPISettings SerializeProgressOptionsRibbon()
    RefreshRibbon
    If Not gobjRibbon Is Nothing Then gobjRibbon.Invalidate
    RefreshProgressBar
End Sub

Public Sub PI_AddBandBackdrop(ByVal targetSlide As Slide, ByVal band As Shape, _
                             ByVal heading As Shape, opts As ProgressDotsOptions, _
                             ByVal chapterName As String, ByVal state As Integer, _
                             ByVal sectionIndex As Long, ByVal shapeTag As String, _
                             ByVal linkTarget As String)
    Dim source As Shape, picture As Shape
    Dim textWidth As Double, textHeight As Double, coverScale As Double
    Dim pictureWidth As Double, pictureHeight As Double
    Dim glyphLeft As Double, glyphTop As Double, offsetX As Double, offsetY As Double
    Dim farEastTypeface As String
    Dim savedNumber As Long, savedDescription As String

    If opts.SectionBandBackdrop <> 1 Or state = 1 Then Exit Sub
    If Len(Trim$(chapterName)) = 0 Then Exit Sub
    If band.Width <= 0 Or band.Height <= 0 Then Exit Sub
    If opts.SectionBandBackdropTransparency >= 100 Then Exit Sub
    On Error GoTo CropError

    ' Export the transparent text directly. The picture viewport clips the glyphs,
    ' leaving the underlying band and slide background visible between strokes.
    Set source = targetSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 1, 1)
    source.Fill.Visible = msoFalse
    source.Line.Visible = msoFalse
    farEastTypeface = heading.TextFrame2.TextRange.Font.NameFarEast
    If Len(farEastTypeface) = 0 Then farEastTypeface = opts.fontName
    With source.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoFalse
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = chapterName
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.ParagraphFormat.SpaceBefore = 0
        .TextRange.ParagraphFormat.SpaceAfter = 0
        With .TextRange.Font
            .Name = opts.fontName
            .NameFarEast = farEastTypeface
            .Size = 96
            .Bold = msoTrue
            .Fill.Solid
            .Fill.ForeColor.RGB = ColorToRGB(opts.SectionBandBackdropColor)
            .Fill.Transparency = PI_BandClampPercent(opts.SectionBandBackdropTransparency) / 100
        End With
        .AutoSize = msoAutoSizeShapeToFitText
        textWidth = .TextRange.BoundWidth
        textHeight = .TextRange.BoundHeight
        glyphLeft = .TextRange.BoundLeft
        glyphTop = .TextRange.BoundTop
    End With
    If textWidth <= 0 Or textHeight <= 0 Then
        source.Delete
        Exit Sub
    End If

    coverScale = band.Width / textWidth
    If band.Height / textHeight > coverScale Then coverScale = band.Height / textHeight
    coverScale = coverScale * PI_BandClampBackdropScale(opts.SectionBandBackdropScale) / 100
    pictureWidth = source.Width * coverScale
    pictureHeight = source.Height * coverScale
    ' Center the glyph bounds, not the font's surrounding line-box whitespace.
    offsetX = (source.Left + source.Width / 2 - glyphLeft - textWidth / 2) * coverScale
    offsetY = (source.Top + source.Height / 2 - glyphTop - textHeight / 2) * coverScale

    Set picture = PI_EmbedShapePNG(targetSlide, source)
    source.Delete
    Set source = Nothing
    picture.LockAspectRatio = msoFalse
    With picture.PictureFormat.Crop
        .PictureWidth = pictureWidth
        .PictureHeight = pictureHeight
        .ShapeWidth = band.Width
        .ShapeHeight = band.Height
        .ShapeLeft = band.Left
        .ShapeTop = band.Top
        .PictureOffsetX = offsetX
        .PictureOffsetY = offsetY
    End With
    picture.Line.Visible = msoFalse
    picture.Name = "PI cropped title " & CStr(sectionIndex)
    picture.AlternativeText = chapterName
    picture.Tags.Add shapeTag, "yes"
    picture.Tags.Add "PI_ROLE", "section-backdrop"
    picture.Tags.Add "PI_SECTION", CStr(sectionIndex)
    If opts.LayeredNavigation = 1 Then picture.Tags.Add "PI_DEPTH", band.Tags("PI_DEPTH")
    picture.ActionSettings(ppMouseClick).Action = ppActionHyperlink
    picture.ActionSettings(ppMouseClick).Hyperlink.SubAddress = linkTarget
    ' The heading is the previous permanent shape; stay above the band and
    ' within the contiguous range that DrawProgressBarOnSlide groups afterward.
    picture.ZOrder msoBringToFront
    Do While picture.ZOrderPosition > heading.ZOrderPosition
        picture.ZOrder msoSendBackward
    Loop
    If mPIRunActive Then mPICaptures = mPICaptures + 1
    Exit Sub

CropError:
    savedNumber = Err.Number
    savedDescription = Err.Description
    On Error Resume Next
    If Not picture Is Nothing Then picture.Delete
    If Not source Is Nothing Then source.Delete
    On Error GoTo 0
    Err.Raise vbObjectError + 3020, "PI_AddBandBackdrop", _
        PI_Zh("88C1520780CC666F5B57751F621059318D25FF1A") & CStr(savedNumber) & " " & savedDescription
End Sub

Public Function PI_BandGradientPosition(sections As ProgressDotsSectionInfo, _
                                      ByVal sectionIndex As Long, ByVal currentSlide As Long) As Double
    Dim index As Long, visibleOrdinal As Long
    Dim targetOrdinal As Long, currentOrdinal As Long, currentSection As Long
    Dim distance As Long, sideCount As Long
    PI_BandGradientPosition = -1
    ' Find the current physical section before counting only displayed bands.
    For index = 1 To sections.lngSecSections
        If currentSlide >= sections.lngArySecStart(index) Then currentSection = index
    Next index
    For index = 1 To sections.lngSecSections
        If sections.intArySecVisible(index) = 1 Then
            visibleOrdinal = visibleOrdinal + 1
            If index = sectionIndex Then targetOrdinal = visibleOrdinal
            If index = currentSection Then currentOrdinal = visibleOrdinal
        End If
    Next index
    If targetOrdinal > 0 And currentOrdinal > 0 Then
        distance = Abs(targetOrdinal - currentOrdinal)
        If targetOrdinal < currentOrdinal Then
            sideCount = currentOrdinal - 1
        Else
            sideCount = visibleOrdinal - currentOrdinal
        End If
        ' The nearest band uses the near color. A lone band also uses the near color.
        PI_BandGradientPosition = 0
        If distance > 0 And sideCount > 1 Then PI_BandGradientPosition = (distance - 1) / (sideCount - 1)
    End If
End Function

Private Function PI_BandSlideBackground(ByVal slideNum As Long) As Long
    Dim presentation As Presentation
    Set presentation = PI_RenderPresentation()
    PI_BandSlideBackground = RGB(255, 255, 255)
    If slideNum < 1 Or slideNum > presentation.Slides.Count Then Exit Function
    ' A single RGB value is available for solid backgrounds, including theme colors.
    With presentation.Slides(slideNum).Background.Fill
        If .Type = msoFillSolid Then PI_BandSlideBackground = .ForeColor.RGB
    End With
End Function

Private Function PI_BandCompositeColor(ByVal color As Long, ByVal transparency As Double, ByVal backdrop As Long) As Long
    Dim red As Long, green As Long, blue As Long
    red = CLng((color And &HFF&) * (1 - transparency) + (backdrop And &HFF&) * transparency)
    green = CLng(((color \ 256) And &HFF&) * (1 - transparency) + ((backdrop \ 256) And &HFF&) * transparency)
    blue = CLng(((color \ 65536) And &HFF&) * (1 - transparency) + ((backdrop \ 65536) And &HFF&) * transparency)
    PI_BandCompositeColor = RGB(red, green, blue)
End Function

Public Function PI_ContrastColor(backgroundColor As Long) As Long
    Dim red As Double, green As Double, blue As Double, luminance As Double
    red = PI_LinearChannel(backgroundColor And &HFF&)
    green = PI_LinearChannel((backgroundColor \ 256) And &HFF&)
    blue = PI_LinearChannel((backgroundColor \ 65536) And &HFF&)
    luminance = 0.2126 * red + 0.7152 * green + 0.0722 * blue
    If luminance > 0.179 Then
        PI_ContrastColor = RGB(20, 27, 36)
    Else
        PI_ContrastColor = RGB(255, 255, 255)
    End If
End Function

Private Function PI_LinearChannel(value As Long) As Double
    Dim normalized As Double
    normalized = value / 255
    If normalized <= 0.04045 Then
        PI_LinearChannel = normalized / 12.92
    Else
        PI_LinearChannel = ((normalized + 0.055) / 1.055) ^ 2.4
    End If
End Function

Public Sub PI_BandTextContrast(target As Shape, opts As ProgressDotsOptions, ByVal state As Integer, _
                               ByVal position As Double, ByVal slideNum As Long)
    Dim enabled As Boolean, alpha As Double, color As Long
    If opts.SectionBandContrast <> 1 Then Exit Sub
    If Not PI_UsesSectionBand(opts) Then Exit Sub
    color = PI_SectionBackground(opts, state, enabled, alpha, position)
    If Not enabled Then Exit Sub
    color = PI_BandCompositeColor(color, alpha, PI_BandSlideBackground(slideNum))
    target.TextFrame.TextRange.Font.Color.RGB = PI_ContrastColor(color)
End Sub

Public Sub PI_BandDotContrast(target As Shape, opts As ProgressDotsOptions, ByVal state As Integer, _
                              Optional ByVal position As Double = -1, Optional ByVal slideNum As Long = 0)
    Dim enabled As Boolean, alpha As Double, color As Long
    If opts.SectionBandContrast <> 1 Then Exit Sub
    If gblshowCustomImage <> 0 And opts.FlatBar <> 1 Then Exit Sub
    color = PI_SectionBackground(opts, state, enabled, alpha, position)
    If Not enabled Then Exit Sub
    If PI_UsesSectionBand(opts) Or opts.SectionBandGradient = 1 Or opts.LayeredNavigation = 1 Then
        color = PI_BandCompositeColor(color, alpha, PI_BandSlideBackground(slideNum))
    End If
    color = PI_ContrastColor(color)
    target.Fill.ForeColor.RGB = color
    target.Line.ForeColor.RGB = color
End Sub


Public Function PI_IsSilentGeneration() As Boolean
    PI_IsSilentGeneration = mPISilent
End Function
