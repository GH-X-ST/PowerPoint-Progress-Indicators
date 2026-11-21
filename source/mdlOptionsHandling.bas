Attribute VB_Name = "mdlOptionsHandling"
' ------------------------------------------------------
' Name: mdlOptionsHandling
' Kind: Modul
' Purpose: Read and write options from and to settings (presentation). Split options-string into different values, combine values to new options-string.
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
'     : 2018-04-29: integrated Active Slide / Section color options
' ------------------------------------------------------
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: DeserializeProgressOptions
' Purpose: Split options-string into different values
' Procedure Kind: Function
' Procedure Access: Public
' Parameter s (String): option to use
' Return Type: ProgressDotsOptions
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function DeserializeProgressOptions(s As String) As ProgressDotsOptions
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct DeserializeProgressOptions: " & s         'Debug: Write active procedure to log
              
          ' Split the string into words
          Dim words() As String
30        words() = Split(s)
          
          ' Use those words to populate theOptions.
          Dim theOptions As ProgressDotsOptions
          
40        If gblIndicatorSettingsVersion = "" Then
50            MsgBox PI_Zh("63D24EF69700898191CD65B0521D59CB5316FF0C8BF791CD65B0542F52A800200050006F0077006500720050006F0069006E00743002"), vbCritical, PI_Zh("9700898191CD542F")
60            gblIndicatorSettingsVersion = "something is wrong, use defaults!"
70        End If
80        If words(0) <> gblIndicatorSettingsVersion Then
90            theOptions = getDefaultOptions()
100       Else
              
110           theOptions.fontName = Replace(words(11), "#", " ")
120           theOptions.textHeight = CDbl(words(4))
130           theOptions.seenTextColor = DeserializeColor(words(9))
140           theOptions.unseenTextColor = DeserializeColor(words(10))
              
150           theOptions.dotShape = words(12)
160           theOptions.dotShapeID = CInt(words(13))
170           theOptions.dotSize = CDbl(words(5))
180           theOptions.seenDotColor = DeserializeColor(words(7))
190           theOptions.unseenDotColor = DeserializeColor(words(8))
              
200           theOptions.LeftMargin = CDbl(words(1))
210           theOptions.RightMargin = CDbl(words(2))
220           theOptions.yMargin = CDbl(words(3))
              
230           theOptions.SectionGap = CDbl(words(6))
              
240           theOptions.ignoreSkipped = CInt(words(14))
250           theOptions.ignoreHidden = CInt(words(15))
260           theOptions.ignoreSections = Replace(words(16), "#", " ")

270           theOptions.DirectUpdate = CInt(words(17))
              
280           theOptions.IndicatorSpaceAdjustment = CInt(words(18))

290           theOptions.useTextColorBG = CInt(words(19))
300           theOptions.useTextColorBG_Current = CInt(words(20))
310           theOptions.useTextColorBG_UnSeen = CInt(words(21))
320           theOptions.useShapeColorBG_Seen = CInt(words(22))
330           theOptions.useShapeColorBG_Current = CInt(words(23))
340           theOptions.useShapeColorBG_UnSeen = CInt(words(24))

350           theOptions.seenTextColorBG = DeserializeColor(words(25))
360           theOptions.unseenTextColorBG = DeserializeColor(words(26))

370           theOptions.IndicatorSqueeze = CInt(words(27))

380           theOptions.SectionPlacement = CInt(words(28))
              
390           theOptions.SectionIndicatorGap = CInt(words(29))


400           theOptions.CurrentDotColor = DeserializeColor(words(30))
410           theOptions.CurrentTextColorBG = DeserializeColor(words(31))
420           theOptions.CurrentTextColor = DeserializeColor(words(32))


430           theOptions.ShapeLineWeight = CInt(words(33))
440           theOptions.ShapeLineStyle = CInt(words(34))
450           theOptions.ShapeLineDashStyle = CInt(words(35))

460           theOptions.ignoreIndicator = Replace(words(36), "#", " ")


470           theOptions.transparencyIndicatorSeen = CInt(words(37))
480           theOptions.transparencyIndicatorActive = CInt(words(38))
490           theOptions.transparencyIndicatorUnseen = CInt(words(39))
          
500           theOptions.transparencyFontSeen = CInt(words(40))
510           theOptions.transparencyFontActive = CInt(words(41))
520           theOptions.transparencyFontUnseen = CInt(words(42))

530           theOptions.EvenSpaceSecOrInd = Replace(words(43), "#", " ")

540           theOptions.PIStyle = Replace(words(44), "#", " ")
              theOptions.SectionBand = 0
              theOptions.SectionBandPadding = 4
              theOptions.SectionBandContrast = 0
              If UBound(words) >= 45 Then theOptions.SectionBand = Val(words(45))
              If UBound(words) >= 46 Then theOptions.SectionBandPadding = Val(words(46))
              If UBound(words) >= 47 Then theOptions.SectionBandContrast = Val(words(47))
              If theOptions.SectionBandPadding < 0 Then theOptions.SectionBandPadding = 0
              If theOptions.SectionBandPadding > 20 Then theOptions.SectionBandPadding = 20
              theOptions.SectionBandGradient = 0
              theOptions.SectionBandTransparencyCurrent = 0
              theOptions.SectionBandTransparencyLeft = 0
              theOptions.SectionBandTransparencyRight = 0
              theOptions.SectionBandLeftNear = DeserializeColor("96;112;132")
              theOptions.SectionBandLeftFar = DeserializeColor("203;210;219")
              theOptions.SectionBandRightNear = DeserializeColor("184;196;210")
              theOptions.SectionBandRightFar = DeserializeColor("229;233;238")
              If UBound(words) >= 48 Then theOptions.SectionBandGradient = Val(words(48))
              If UBound(words) >= 49 Then theOptions.SectionBandTransparencyCurrent = PI_BandClampPercent(Val(words(49)))
              If UBound(words) >= 50 Then theOptions.SectionBandTransparencyLeft = PI_BandClampPercent(Val(words(50)))
              If UBound(words) >= 51 Then theOptions.SectionBandTransparencyRight = PI_BandClampPercent(Val(words(51)))
              If UBound(words) >= 52 Then theOptions.SectionBandLeftNear = DeserializeColor(words(52))
              If UBound(words) >= 53 Then theOptions.SectionBandLeftFar = DeserializeColor(words(53))
              If UBound(words) >= 54 Then theOptions.SectionBandRightNear = DeserializeColor(words(54))
              If UBound(words) >= 55 Then theOptions.SectionBandRightFar = DeserializeColor(words(55))
              theOptions.SectionBandBackdrop = 0
              theOptions.SectionBandBackdropColor = DeserializeColor("255;255;255")
              theOptions.SectionBandBackdropTransparency = 75
              theOptions.SectionBandBackdropScale = 150
              If UBound(words) >= 56 Then theOptions.SectionBandBackdrop = Val(words(56))
              If UBound(words) >= 57 Then theOptions.SectionBandBackdropColor = DeserializeColor(words(57))
              If UBound(words) >= 58 Then theOptions.SectionBandBackdropTransparency = PI_BandClampPercent(Val(words(58)))
              If UBound(words) >= 59 Then theOptions.SectionBandBackdropScale = PI_BandClampBackdropScale(Val(words(59)))
              theOptions.FlatBar = 0
              theOptions.FlatBarHeight = 2.5
              theOptions.FlatBarGap = 3
              theOptions.FlatBarRoundness = 50
              theOptions.FlatBarGlow = 1
              theOptions.FlatBarGlowRadius = 4
              theOptions.FlatBarGlowTransparency = 55
              If UBound(words) >= 60 Then theOptions.FlatBar = Abs(CInt(Val(words(60)) = 1))
              If UBound(words) >= 61 Then theOptions.FlatBarHeight = PI_BarClamp(Val(words(61)), 0.5, 12)
              If UBound(words) >= 62 Then theOptions.FlatBarGap = PI_BarClamp(Val(words(62)), 0, 20)
              If UBound(words) >= 63 Then theOptions.FlatBarRoundness = PI_BarClamp(Val(words(63)), 0, 50)
              If UBound(words) >= 64 Then theOptions.FlatBarGlow = Abs(CInt(Val(words(64)) = 1))
              If UBound(words) >= 65 Then theOptions.FlatBarGlowRadius = PI_BarClamp(Val(words(65)), 0, 20)
              If UBound(words) >= 66 Then theOptions.FlatBarGlowTransparency = PI_BarClamp(Val(words(66)), 0, 100)
              theOptions.LayeredNavigation = 0
              theOptions.NavigationWidth = 0
              theOptions.SectionBlockHeight = 34
              theOptions.RailWidth = 0
              theOptions.RailThickness = 1
              theOptions.LayerShadowBlur = 6
              theOptions.LayerShadowTransparency = 80
              theOptions.LayerDepthStep = 0.75
              theOptions.LayerShadowDistance = 1.5
              If UBound(words) >= 67 Then theOptions.LayeredNavigation = Abs(CInt(Val(words(67)) = 1))
              If UBound(words) >= 68 Then theOptions.NavigationWidth = PI_BarClamp(Val(words(68)), 0, 4032)
              If UBound(words) >= 69 Then theOptions.SectionBlockHeight = PI_BarClamp(Val(words(69)), 0, 200)
              If UBound(words) >= 70 Then theOptions.RailWidth = PI_BarClamp(Val(words(70)), 0, 4032)
              If UBound(words) >= 71 Then theOptions.RailThickness = PI_BarClamp(Val(words(71)), 0.5, 30)
              If UBound(words) >= 72 Then theOptions.LayerShadowBlur = PI_BarClamp(Val(words(72)), 0, 20)
              If UBound(words) >= 73 Then theOptions.LayerShadowTransparency = PI_BarClamp(Val(words(73)), 0, 100)
              If UBound(words) >= 74 Then theOptions.LayerDepthStep = PI_BarClamp(Val(words(74)), 0, 5)
              If UBound(words) >= 75 Then theOptions.LayerShadowDistance = PI_BarClamp(Val(words(75)), 0, 20)
              theOptions.AccentMode = 0
              theOptions.AccentWidth = 1.5
              theOptions.AccentColor = DeserializeColor("158;31;56")
              theOptions.SectionNumbering = 0
              theOptions.CurrentTitleBold = 0
              theOptions.FutureBarOutline = 0
              If UBound(words) >= 76 Then theOptions.AccentMode = CInt(PI_BarClamp(Val(words(76)), 0, 4))
              If UBound(words) >= 77 Then theOptions.AccentWidth = PI_BarClamp(Val(words(77)), 0.5, 4)
              If UBound(words) >= 78 Then theOptions.AccentColor = DeserializeColor(words(78))
              If UBound(words) >= 79 Then theOptions.SectionNumbering = Abs(CInt(Val(words(79)) = 1))
              If UBound(words) >= 80 Then theOptions.CurrentTitleBold = Abs(CInt(Val(words(80)) = 1))
              If UBound(words) >= 81 Then theOptions.FutureBarOutline = Abs(CInt(Val(words(81)) = 1))
              theOptions.SectionForegroundBars = 0
              If UBound(words) >= 82 Then theOptions.SectionForegroundBars = Abs(CInt(Val(words(82)) = 1))
              theOptions.DockEdge = 0
              theOptions.SidebarWidth = 96
              theOptions.DisplayBindingMode = 0
              theOptions.DisplayBindingId = ""
              theOptions.UseThemeColors = 0
              theOptions.SectionRoundness = 0
              theOptions.ReuseNavigation = 1
              theOptions.HardShadowOffset = 0
              If UBound(words) >= 83 Then theOptions.DockEdge = CInt(PI_BarClamp(Val(words(83)), 0, 3))
              If UBound(words) >= 84 Then theOptions.SidebarWidth = PI_BarClamp(Val(words(84)), 48, 240)
              If UBound(words) >= 85 Then theOptions.DisplayBindingMode = CInt(PI_BarClamp(Val(words(85)), 0, 2))
              If UBound(words) >= 86 Then
                  If words(86) <> "-" Then theOptions.DisplayBindingId = words(86)
              End If
              If UBound(words) >= 87 Then theOptions.UseThemeColors = Abs(CInt(Val(words(87)) = 1))
              If UBound(words) >= 88 Then theOptions.SectionRoundness = PI_BarClamp(Val(words(88)), 0, 50)
              If UBound(words) >= 89 Then theOptions.ReuseNavigation = Abs(CInt(Val(words(89)) = 1))
              If UBound(words) >= 90 Then theOptions.HardShadowOffset = PI_BarClamp(Val(words(90)), 0, 8)
              theOptions.CounterMode = 0
              theOptions.DensityMode = 0
              theOptions.MinMarkerLength = 4
              theOptions.TitleOverflow = 0
              theOptions.TitleAlignment = 0
              theOptions.TitleCase = 0
              theOptions.FlatRuleMode = 0
              theOptions.FlatRuleWidth = 0.5
              theOptions.FlatRuleColor = DeserializeColor("107;114;128")
              If UBound(words) >= 91 Then theOptions.CounterMode = CInt(PI_BarClamp(Val(words(91)), 0, 5))
              If UBound(words) >= 92 Then theOptions.DensityMode = CInt(PI_BarClamp(Val(words(92)), 0, 2))
              If UBound(words) >= 93 Then theOptions.MinMarkerLength = PI_BarClamp(Val(words(93)), 2, 20)
              If UBound(words) >= 94 Then theOptions.TitleOverflow = CInt(PI_BarClamp(Val(words(94)), 0, 2))
              If UBound(words) >= 95 Then theOptions.TitleAlignment = CInt(PI_BarClamp(Val(words(95)), 0, 2))
              If UBound(words) >= 96 Then theOptions.TitleCase = CInt(PI_BarClamp(Val(words(96)), 0, 1))
              If UBound(words) >= 97 Then theOptions.FlatRuleMode = CInt(PI_BarClamp(Val(words(97)), 0, 8))
              If UBound(words) >= 98 Then theOptions.FlatRuleWidth = PI_BarClamp(Val(words(98)), 0.25, 3)
              If UBound(words) >= 99 Then theOptions.FlatRuleColor = DeserializeColor(words(99))
              theOptions.MarkerKind = 0
              theOptions.PanelShadow = 0
              If UBound(words) >= 100 Then theOptions.MarkerKind = CInt(Val(words(100)))
              If UBound(words) >= 101 Then theOptions.PanelShadow = CInt(Val(words(101)))
              If UBound(words) >= 102 Then
                  theOptions.PresetSize = CInt(Val(words(102)))
              ElseIf theOptions.MarkerKind = 3 And theOptions.dotSize = 5 Then
                  ' Existing node presets used the bar-height setting to size the shape.
                  theOptions.dotSize = theOptions.FlatBarHeight
              End If
              If UBound(words) >= 103 Then theOptions.CrimsonType = CInt(Val(words(103)))
              If UBound(words) >= 104 Then
                  theOptions.BandHeightMode = CInt(Val(words(104)))
                  If theOptions.PresetSize = 2 And theOptions.CrimsonType = 1 And _
                          theOptions.BandHeightMode = 0 Then
                      Select Case theOptions.SectionBlockHeight
                          Case PI_SMALL_BODY_HEIGHT, 16, 14.17
                              theOptions.SectionBlockHeight = PI_CRIMSON_SMALL_BODY_HEIGHT
                      End Select
                  End If
              End If
              If UBound(words) >= 105 Then
                  theOptions.InactiveSectionBarTransparency = CInt(words(105))
              ElseIf theOptions.CrimsonType = 1 Then
                  theOptions.InactiveSectionBarTransparency = PI_CRIMSON_NONCURRENT_PAGE_TRANSPARENCY
                  If words(25) = "212;215;219" Then theOptions.seenTextColorBG = DeserializeColor("222;222;227")
                  If words(26) = "230;232;236" Then theOptions.unseenTextColorBG = DeserializeColor("242;242;247")
              End If
              If theOptions.CrimsonType = 1 And UBound(words) < 106 Then
                  ' Earlier presets used 70% for other slides and 50% for other sections.
                  If theOptions.transparencyIndicatorSeen = 70 Then theOptions.transparencyIndicatorSeen = PI_CRIMSON_NONCURRENT_PAGE_TRANSPARENCY
                  If theOptions.transparencyIndicatorUnseen = 70 Then theOptions.transparencyIndicatorUnseen = PI_CRIMSON_NONCURRENT_PAGE_TRANSPARENCY
                  If theOptions.InactiveSectionBarTransparency = 50 Then theOptions.InactiveSectionBarTransparency = PI_CRIMSON_NONCURRENT_PAGE_TRANSPARENCY
              End If


550       End If
          
          ' Return it
560       DeserializeProgressOptions = theOptions
          
exit_proc:
570       Exit Function
          
ErrorHandler:
          If PI_IsSilentGeneration() Then PI_RaiseNativeFailure "DeserializeProgressOptions", Erl, Err.Number, Err.Source, Err.Description
580       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
590       Select Case Err.Number
          
              Case Else
600               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "DeserializeProgressOptions"
610               Resume Next
620               Resume
              
630       End Select
          
End Function

' ----------------------------------------------------------------
' Procedure Name: getDefaultOptions
' Purpose: grab global default values or use settings from a previous progress indicator
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: ProgressDotsOptions
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function getDefaultOptions() As ProgressDotsOptions
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct getDefaultOptions"         'Debug: Write active procedure to log
          
          Dim strRGBArray() As String
          
30        With getDefaultOptions:

40            .PIStyle = gblstrIndicatorStyle

              'font
50            .fontName = gblstrFontname
60            .textHeight = gbldblFontSize
70            .DirectUpdate = gblintGeneralProgressbarDirect
              
80            strRGBArray = Split(gblstrFontColorSeen, ";")
90            .seenTextColor.red = Val(strRGBArray(0))
100           .seenTextColor.green = Val(strRGBArray(1))
110           .seenTextColor.blue = Val(strRGBArray(2))
              
120           strRGBArray = Split(gblstrFontColorUnSeen, ";")
130           .unseenTextColor.red = Val(strRGBArray(0))
140           .unseenTextColor.green = Val(strRGBArray(1))
150           .unseenTextColor.blue = Val(strRGBArray(2))
                 
              'indicator
160           .dotShape = gblstrIndicatorType
170           .dotShapeID = gblintIndicatorTypeID
180           .dotSize = gbldblIndicatorSize
              
190           strRGBArray = Split(gblstrIndicatorColorSeen, ";")
200           .seenDotColor.red = Val(strRGBArray(0))
210           .seenDotColor.green = Val(strRGBArray(1))
220           .seenDotColor.blue = Val(strRGBArray(2))
              
230           strRGBArray = Split(gblstrIndicatorColorUnSeen, ";")
240           .unseenDotColor.red = Val(strRGBArray(0))
250           .unseenDotColor.green = Val(strRGBArray(1))
260           .unseenDotColor.blue = Val(strRGBArray(2))
              
              'margin + sections
270           .yMargin = gblintMarginTop
280           .LeftMargin = gblintMarginLeft
290           .RightMargin = gblintMarginRight
          
300           .SectionGap = gblintSectionSpacing
310           .IndicatorSpaceAdjustment = gblintIndicatorSpacingAdjustment
320           .IndicatorSqueeze = gblintIndicatorSqueeze
                      
              'ignore
330           .ignoreSkipped = gblintSlideIgnoreSkipped
340           .ignoreHidden = gblintSlideIgnoreHidden
350           .ignoreSections = gblstrSlideIgnoreSections
              
              'BG colors of section labels
360           .useTextColorBG = gblintUseTextBGCol
370           .useTextColorBG_Current = gblintUseTextBGCol_Current
380           .useTextColorBG_UnSeen = gblintUseTextBGCol_UnSeen
390           .useShapeColorBG_Seen = gblintUseShapeBG_Seen
400           .useShapeColorBG_Current = gblintUseShapeBG_Current
410           .useShapeColorBG_UnSeen = gblintUseShapeBG_UnSeen

420           strRGBArray = Split(gblstrFontBGColorSeen, ";")
430           .seenTextColorBG.red = Val(strRGBArray(0))
440           .seenTextColorBG.green = Val(strRGBArray(1))
450           .seenTextColorBG.blue = Val(strRGBArray(2))

460           strRGBArray = Split(gblstrFontBGColorUnSeen, ";")
470           .unseenTextColorBG.red = Val(strRGBArray(0))
480           .unseenTextColorBG.green = Val(strRGBArray(1))
490           .unseenTextColorBG.blue = Val(strRGBArray(2))
         
500           .SectionPlacement = gblintSectionPlacement
              
510           .SectionIndicatorGap = gblintSectionIndicatorGap
             
              'Active/Current Colors
520           strRGBArray = Split(gblstrIndicatorColorCurrent, ";")
530           .CurrentDotColor.red = Val(strRGBArray(0))
540           .CurrentDotColor.green = Val(strRGBArray(1))
550           .CurrentDotColor.blue = Val(strRGBArray(2))
           
560           strRGBArray = Split(gblstrFontBGColorCurrent, ";")
570           .CurrentTextColorBG.red = Val(strRGBArray(0))
580           .CurrentTextColorBG.green = Val(strRGBArray(1))
590           .CurrentTextColorBG.blue = Val(strRGBArray(2))
                            
600           strRGBArray = Split(gblstrFontColorCurrent, ";")
610           .CurrentTextColor.red = Val(strRGBArray(0))
620           .CurrentTextColor.green = Val(strRGBArray(1))
630           .CurrentTextColor.blue = Val(strRGBArray(2))
                            
              'shape: ShapeLineStyle etc.
640           .ShapeLineWeight = gblShapeLineWeight
650           .ShapeLineStyle = gblShapeLineStyle
660           .ShapeLineDashStyle = gblShapeLineDashStyle
          
670           .ignoreIndicator = gblstrSlideIgnoreIndicator
              
              
680           .transparencyIndicatorSeen = gblTransparencyIndicatorSeen
690           .transparencyIndicatorActive = gblTransparencyIndicatorActive
700           .transparencyIndicatorUnseen = gblTransparencyIndicatorUnseen
          
710           .transparencyFontSeen = gblTransparencyFontSeen
720           .transparencyFontActive = gblTransparencyFontActive
730           .transparencyFontUnseen = gblTransparencyFontUnseen

740           .EvenSpaceSecOrInd = gblEvenSpaceSecOrInd
              .SectionBand = gblSectionBand
              .SectionBandPadding = gblSectionBandPadding
              .SectionBandContrast = gblSectionBandContrast
              .SectionBandGradient = gblSectionBandGradient
              .SectionBandTransparencyCurrent = gblSectionBandTransparencyCurrent
              .SectionBandTransparencyLeft = gblSectionBandTransparencyLeft
              .SectionBandTransparencyRight = gblSectionBandTransparencyRight
              .SectionBandLeftNear = DeserializeColor(gblSectionBandLeftNear)
              .SectionBandLeftFar = DeserializeColor(gblSectionBandLeftFar)
              .SectionBandRightNear = DeserializeColor(gblSectionBandRightNear)
              .SectionBandRightFar = DeserializeColor(gblSectionBandRightFar)
              .SectionBandBackdrop = gblSectionBandBackdrop
              .SectionBandBackdropColor = DeserializeColor(gblSectionBandBackdropColor)
              .SectionBandBackdropTransparency = gblSectionBandBackdropTransparency
              .SectionBandBackdropScale = gblSectionBandBackdropScale
              .FlatBar = gblFlatBar
              .FlatBarHeight = gblFlatBarHeight
              .FlatBarGap = gblFlatBarGap
              .FlatBarRoundness = gblFlatBarRoundness
              .FlatBarGlow = gblFlatBarGlow
              .FlatBarGlowRadius = gblFlatBarGlowRadius
              .FlatBarGlowTransparency = gblFlatBarGlowTransparency
              .LayeredNavigation = gblLayeredNavigation
              .NavigationWidth = gblNavigationWidth
              .SectionBlockHeight = gblSectionBlockHeight
              .RailWidth = gblRailWidth
              .RailThickness = gblRailThickness
              .LayerShadowBlur = gblLayerShadowBlur
              .LayerShadowTransparency = gblLayerShadowTransparency
              .LayerDepthStep = gblLayerDepthStep
              .LayerShadowDistance = gblLayerShadowDistance
              .AccentMode = gblAccentMode
              .AccentWidth = gblAccentWidth
              .AccentColor = DeserializeColor(gblAccentColor)
              .SectionNumbering = gblSectionNumbering
              .CurrentTitleBold = gblCurrentTitleBold
              .FutureBarOutline = gblFutureBarOutline
              .SectionForegroundBars = gblSectionForegroundBars
              .DockEdge = gblDockEdge
              .SidebarWidth = gblSidebarWidth
              .DisplayBindingMode = gblDisplayBindingMode
              .DisplayBindingId = gblDisplayBindingId
              .UseThemeColors = gblUseThemeColors
              .SectionRoundness = gblSectionRoundness
              .ReuseNavigation = gblReuseNavigation
              .HardShadowOffset = gblHardShadowOffset
              .CounterMode = gblCounterMode
              .DensityMode = gblDensityMode
              .MinMarkerLength = gblMinMarkerLength
              .TitleOverflow = gblTitleOverflow
              .TitleAlignment = gblTitleAlignment
              .TitleCase = gblTitleCase
              .FlatRuleMode = gblFlatRuleMode
              .FlatRuleWidth = gblFlatRuleWidth
              .FlatRuleColor = DeserializeColor(gblFlatRuleColor)
              .MarkerKind = gblMarkerKind
              .PanelShadow = gblPanelShadow
              .PresetSize = gblPresetSize
              .CrimsonType = gblCrimsonType
              .InactiveSectionBarTransparency = gblInactiveSectionBarTransparency
              .BandHeightMode = gblBandHeightMode
              
              
              
750       End With

exit_proc:
760       Exit Function
          
ErrorHandler:
          If PI_IsSilentGeneration() Then PI_RaiseNativeFailure "getDefaultOptions", Erl, Err.Number, Err.Source, Err.Description
770       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
780       Select Case Err.Number
          
              Case Else
790               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "getDefaultOptions"
800               Resume Next
810               Resume
              
820       End Select

End Function

Sub InitializePIVersion()
          'Version for settings-format saved in presentation
10        gblIndicatorSettingsVersion = "Version20231115_1200"
20        gblProgramVersionBase = "1.960"   'this is checked at online version check, see cmdWebCheck_Click in frmHelp
30        gblProgramVersion = gblProgramVersionBase & " CN.11.10 32b/64b" & " / Settings: " & gblIndicatorSettingsVersion
          
End Sub

' ----------------------------------------------------------------
' Procedure Name: Initialize
' Purpose: set all base variables to their default values.
'          This happens automatically, when the macro is being started.
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub Initialize()
10        On Error Resume Next
20        If gblbolDebug = True Then LogInfo "Sub Initialize"         'Debug: Write active procedure to log
              
          'Initialize all variables here that are needed somewhere

30        InitializePIVersion
          
          'define some variables so it will be easier in the furture to adjust code
40        gblstrPISectionHidden = "PISectionHidden"
50        gblstrSectionNameHidden = "PISectionnameHidden"
60        gblstrMessageIgnore = "Progress Indicator: Ignore this slide"
70        gblstrMessageRepeat = "Progress Indicator: Repeat previous indicator"
80        gblstrMessageMore = "...more..."
90        gblstrMessageTagTitle = "Progress Indicator title"
100       gblCreatePressed = False

110       gblShowPreview = 0  'Preview-Button Value
          
          'initial / default ribbon settings
120       gblstrIndicatorStyle = "Full"   'Style: Full (Indicator and Sections, Indicator only, Sections only
          
130       gblintGeneralProgressbarDirect = 0      ' update directly after change in ribbon? 1=true, 0=false
          
140       gblstrFontname = "Calibri"
150       gbldblFontSize = 12
160       gblstrFontColorSeen = "79;129;189"        ' color values in RGB, in DEC format - not HEX!
170       gblstrFontColorUnSeen = "220;230;242"     ' color values in RGB, in DEC format - not HEX!
          
180       gblstrIndicatorType = "Oval"
190       gblintIndicatorTypeID = 6       ' Number to select in listbox (not MSO shape number)!
200       gbldblIndicatorSize = 20
210       gblstrIndicatorColorSeen = "90;200;30"    ' color values in RGB, in DEC format - not HEX!
220       gblstrIndicatorColorUnSeen = "10;255;0"   ' color values in RGB, in DEC format - not HEX!
          
230       gblintMarginTop = 5
240       gblintMarginLeft = 0
250       gblintMarginRight = 0
          
260       gblintSectionSpacing = 10
270       gblintIndicatorSpacingAdjustment = 0
280       gblintIndicatorSqueeze = 0
          
290       gblintSlideIgnoreSkipped = 1
300       gblintSlideIgnoreHidden = 1
310       gblstrSlideIgnoreSections = "Full" 'Full|Space Only|None
          
          'BG Colors for sections   0=no BG color, 1= use BG color
320       gblintUseTextBGCol = 0
330       gblintUseTextBGCol_Current = 0
340       gblintUseTextBGCol_UnSeen = 0

          'BG Colors for shapes   0=no BG color, 1= use BG color
350       gblintUseShapeBG_Seen = 1
360       gblintUseShapeBG_Current = 1
370       gblintUseShapeBG_UnSeen = 1

          'fore/textcolors for sections
380       gblstrFontBGColorSeen = "90;200;30"        'use colors of the indicator, color values in RGB, in DEC format - not HEX!
390       gblstrFontBGColorUnSeen = "10;255;0"    'use colors of the indicator, color values in RGB, in DEC format - not HEX!

400       gblstrIndicatorColorCurrent = "10;255;0"         'use colors of the indicator, color values in RGB, in DEC format - not HEX!
410       gblstrFontBGColorCurrent = "220;230;242"         'use colors of the text, color values in RGB, in DEC format - not HEX!
420       gblstrFontColorCurrent = "220;230;242"        'use colors of the text, color values in RGB, in DEC format - not HEX!

430       gblintSectionPlacement = 0

440       gblintSectionIndicatorGap = 0


          'shapes for dropdown
450       gblvarShapePossible = Array( _
              "149", "10pointStar", "150", "12pointStar", "94", "16pointStar", "95", "24pointStar", "96", "32pointStar", "91", "4pointStar", "92", "5pointStar", "147", "6pointStar", "148", "7pointStar", "93", "8pointStar", "129", "ActionButtonBackorPrevious", "131", "ActionButtonBeginning", "125", "ActionButtonCustom", "134", "ActionButtonDocument", "132", "ActionButtonEnd", "130", "ActionButtonForwardorNext", "127", "ActionButtonHelp", "126", "ActionButtonHome", "128", "ActionButtonInformation", "136", "ActionButtonMovie", "133", "ActionButtonReturn", "135", "ActionButtonSound", "25", "Arc", "137", "Balloon", "41", "BentArrow", "44", "BentUpArrow", "15", "Bevel", "20", "BlockArc", "13", "Can", "182", "ChartPlus", "181", "ChartStar", "180", "ChartX", "52", "Chevron", "161", "Chord", "60", "CircularArrow", "179", "Cloud", "108", "CloudCallout", "162", "Corner", "169", "CornerTabs", "11", "Cross", "14", "Cube", "48", "CurvedDownArrow", "100", "CurvedDownRibbon", _
              "46", "CurvedLeftArrow", "45", "CurvedRightArrow", "47", "CurvedUpArrow", "99", "CurvedUpRibbon", "144", "Decagon", "141", "DiagonalStripe", "4", "Diamond", "146", "Dodecagon", "18", "Donut", "27", "DoubleBrace", "26", "DoubleBracket", "104", "DoubleWave", "36", "DownArrow", "56", "DownArrowCallout", "98", "DownRibbon", "89", "Explosion1", "90", "Explosion2", "62", "FlowchartAlternateProcess", "75", "FlowchartCard", "79", "FlowchartCollate", "73", "FlowchartConnector", "64", "FlowchartData", "63", "FlowchartDecision", "84", "FlowchartDelay", "87", "FlowchartDirectAccessStorage", "88", "FlowchartDisplay", "67", "FlowchartDocument", "81", "FlowchartExtract", "66", "FlowchartInternalStorage", "86", "FlowchartMagneticDisk", "71", "FlowchartManualInput", "72", "FlowchartManualOperation", "82", "FlowchartMerge", "68", "FlowchartMultidocument", "139", "FlowchartOfflineStorage", "74", "FlowchartOffpageConnector", "78", "FlowchartOr", _
              "65", "FlowchartPredefinedProcess", "70", "FlowchartPreparation", "61", "FlowchartProcess", "76", "FlowchartPunchedTape", "85", "FlowchartSequentialAccessStorage", "80", "FlowchartSort", "83", "FlowchartStoredData", "77", "FlowchartSummingJunction", "69", "FlowchartTerminator", "16", "FoldedCorner", "158", "Frame", "174", "Funnel", "172", "Gear6", "173", "Gear9", "159", "HalfFrame", "21", "Heart", "145", "Heptagon", "10", "Hexagon", "102", "HorizontalScroll", "7", "IsoscelesTriangle", "34", "LeftArrow", "54", "LeftArrowCallout", "31", "LeftBrace", "29", "LeftBracket", "176", "LeftCircularArrow", "37", "LeftRightArrow", "57", "LeftRightArrowCallout", "177", "LeftRightCircularArrow", "140", "LeftRightRibbon", "40", "LeftRightUpArrow", "43", "LeftUpArrow", "22", "LightningBolt", "109", "LineCallout1", "113", "LineCallout1AccentBar", "121", "LineCallout1BorderandAccentBar", "117", "LineCallout1NoBorder", "110", "LineCallout2", _
              "114", "LineCallout2AccentBar", "122", "LineCallout2BorderandAccentBar", "118", "LineCallout2NoBorder", "111", "LineCallout3", "115", "LineCallout3AccentBar", "123", "LineCallout3BorderandAccentBar", "119", "LineCallout3NoBorder", "112", "LineCallout4", "116", "LineCallout4AccentBar", "124", "LineCallout4BorderandAccentBar", "120", "LineCallout4NoBorder", "183", "LineInverse", "166", "MathDivide", "167", "MathEqual", "164", "MathMinus", "165", "MathMultiply", "168", "MathNotEqual", "163", "MathPlus", "24", "Moon", "143", "NonIsoscelesTrapezoid", "19", "NoSymbol", "50", "NotchedRightArrow", "138", "NotPrimitive", "6", "Octagon", "9", "Oval", "107", "OvalCallout", "2", "Parallelogram", "51", "Pentagon", "142", "Pie", "175", "PieWedge", "28", "Plaque", "171", "PlaqueTabs", "39", "QuadArrow", "59", "QuadArrowCallout", "1", "Rectangle", "105", "RectangularCallout", "12", "RegularPentagon", "33", "RightArrow", "53", "RightArrowCallout", _
              "32", "RightBrace", "30", "RightBracket", "8", "RightTriangle", "151", "Round1Rectangle", "157", "Round2DiagRectangle", "152", "Round2SameRectangle", "5", "RoundedRectangle", "106", "RoundedRectangularCallout", "17", "SmileyFace", "155", "Snip1Rectangle", "157", "Snip2DiagRectangle", "156", "Snip2SameRectangle", "154", "SnipRoundRectangle", "170", "SquareTabs", "49", "StripedRightArrow", "23", "Sun", "178", "SwooshArrow", "160", "Tear", "3", "Trapezoid", "35", "UpArrow", "55", "UpArrowCallout", "38", "UpDownArrow", "58", "UpDownArrowCallout", "97", "UpRibbon", "42", "UTurnArrow", "101", "VerticalScroll", "103", "Wave" _
              )
              
          'for a list of all shapes in the ribbon dropdown
460       gblvarShapeListIDs = Array( _
              "array0_blank", "93", "4", "18", "173", "10", "7", "9", "2", "1", "33", "5", "3", "103", _
              "9991", _
              "9", "149", "150", "94", "95", "96", "91", "92", "147", "148", "93", "129", "131", "125", "134", "132", "130", "127", "126", "128", "136", "133", "135", "25", "137", "41", "44", "15", "20", "13", "182", "181", "180", "52", "161", "60", "179", "108", _
              "162", "169", "11", "14", "48", "100", "46", "45", "47", "99", "144", "141", "4", "146", "18", "27", "26", "104", "36", "56", "98", "89", "90", "62", "75", "79", "73", "64", "63", "84", "87", "88", "67", "81", "66", "86", "71", "72", "82", "68", "139", "74", "78", "65", "70", "61", "76", "85", "80", "83", "77", "69", "16", _
              "158", "174", "172", "173", "159", "21", "145", "10", "102", "7", "34", "54", "31", "29", "176", "37", "57", "177", "140", "40", "43", "22", "109", "113", "121", "117", "110", "114", "122", "118", "111", "115", "123", "119", "112", "116", "124", "120", "183", "166", "167", "164", "165", "168", "163", "24", "143", "19", _
              "50", "138", "6", "9", "107", "2", "51", "142", "175", "28", "171", "39", "59", "1", "105", "12", "33", "53", "32", "30", "8", "151", "157", "152", "5", "106", "17", "155", "157", "156", "154", "170", "49", "23", "178", "160", "3", "35", "55", "38", "58", "97", "42", "101", "103" _
              )
          
470       gblvarShapeListNames = Array( _
              "array0_blank", "8pointStar", "Diamond", "Donut", "Gear9", "Hexagon", "IsoscelesTriangle", "Oval", "Parallelogram", "Rectangle", "RightArrow", "RoundedRectangle", "Trapezoid", "Wave", _
              "-custom image-", _
              "Oval", "10pointStar", "12pointStar", "16pointStar", "24pointStar", "32pointStar", "4pointStar", "5pointStar", "6pointStar", "7pointStar", "8pointStar", "ActionButtonBackorPrevious", "ActionButtonBeginning", "ActionButtonCustom", "ActionButtonDocument", "ActionButtonEnd", "ActionButtonForwardorNext", "ActionButtonHelp", "ActionButtonHome", "ActionButtonInformation", _
              "ActionButtonMovie", "ActionButtonReturn", "ActionButtonSound", "Arc", "Balloon", "BentArrow", "BentUpArrow", "Bevel", "BlockArc", "Can", "ChartPlus", "ChartStar", "ChartX", "Chevron", "Chord", "CircularArrow", "Cloud", "CloudCallout", "Corner", "CornerTabs", "Cross", "Cube", "CurvedDownArrow", "CurvedDownRibbon", "CurvedLeftArrow", "CurvedRightArrow", "CurvedUpArrow", "CurvedUpRibbon", "Decagon", "DiagonalStripe", "Diamond", "Dodecagon", "Donut", "DoubleBrace", "DoubleBracket", "DoubleWave", "DownArrow", "DownArrowCallout", _
              "DownRibbon", "Explosion1", "Explosion2", "FlowchartAlternateProcess", "FlowchartCard", "FlowchartCollate", "FlowchartConnector", "FlowchartData", "FlowchartDecision", "FlowchartDelay", "FlowchartDirectAccessStorage", "FlowchartDisplay", "FlowchartDocument", "FlowchartExtract", "FlowchartInternalStorage", "FlowchartMagneticDisk", "FlowchartManualInput", "FlowchartManualOperation", "FlowchartMerge", "FlowchartMultidocument", "FlowchartOfflineStorage", "FlowchartOffpageConnector", "FlowchartOr", _
              "FlowchartPredefinedProcess", "FlowchartPreparation", "FlowchartProcess", "FlowchartPunchedTape", "FlowchartSequentialAccessStorage", "FlowchartSort", "FlowchartStoredData", "FlowchartSummingJunction", "FlowchartTerminator", "FoldedCorner", "Frame", "Funnel", "Gear6", "Gear9", "HalfFrame", "Heart", "Heptagon", "Hexagon", "HorizontalScroll", "IsoscelesTriangle", "LeftArrow", "LeftArrowCallout", "LeftBrace", "LeftBracket", "LeftCircularArrow", "LeftRightArrow", "LeftRightArrowCallout", "LeftRightCircularArrow", _
              "LeftRightRibbon", "LeftRightUpArrow", "LeftUpArrow", "LightningBolt", "LineCallout1", "LineCallout1AccentBar", "LineCallout1BorderandAccentBar", "LineCallout1NoBorder", "LineCallout2", "LineCallout2AccentBar", "LineCallout2BorderandAccentBar", "LineCallout2NoBorder", "LineCallout3", "LineCallout3AccentBar", "LineCallout3BorderandAccentBar", "LineCallout3NoBorder", "LineCallout4", "LineCallout4AccentBar", "LineCallout4BorderandAccentBar", "LineCallout4NoBorder", "LineInverse", "MathDivide", "MathEqual", _
              "MathMinus", "MathMultiply", "MathNotEqual", "MathPlus", "Moon", "NonIsoscelesTrapezoid", "NoSymbol", "NotchedRightArrow", "NotPrimitive", "Octagon", "Oval", "OvalCallout", "Parallelogram", "Pentagon", "Pie", "PieWedge", "Plaque", "PlaqueTabs", "QuadArrow", "QuadArrowCallout", "Rectangle", "RectangularCallout", "RegularPentagon", "RightArrow", "RightArrowCallout", "RightBrace", "RightBracket", "RightTriangle", "Round1Rectangle", "Round2DiagRectangle", "Round2SameRectangle", "RoundedRectangle", "RoundedRectangularCallout", _
              "SmileyFace", "Snip1Rectangle", "Snip2DiagRectangle", "Snip2SameRectangle", "SnipRoundRectangle", "SquareTabs", "StripedRightArrow", "Sun", "SwooshArrow", "Tear", "Trapezoid", "UpArrow", "UpArrowCallout", "UpDownArrow", "UpDownArrowCallout", "UpRibbon", "UTurnArrow", "VerticalScroll", "Wave" _
              )
          
          'define base line forms   'saved to settings is new in v1.839
480       gblShapeLineWeight = 1
490       gblShapeLineStyle = 1
500       gblShapeLineDashStyle = 1
          
510       gblstrSlideIgnoreIndicator = "All"  ' All|All|Active Section|None

520       gblTransparencyIndicatorSeen = 0
530       gblTransparencyIndicatorActive = 0
540       gblTransparencyIndicatorUnseen = 0
          
550       gblTransparencyFontSeen = 0
560       gblTransparencyFontActive = 0
570       gblTransparencyFontUnseen = 0


580       gblshowCustomImage = 0      ' 0 = use regular symbol; 1 = use custom images (seen, active, unseen)
590       gblCustomImage_Seen = ""
600       gblCustomImage_Active = ""
610       gblCustomImage_UnSeen = ""


620       gblEvenSpaceSecOrInd = "Shapes" ' Alternative: Sections

630       gbliDrawStart = 1
640       gbliDrawEnd = 999
          gblDockEdge = 0
          gblSidebarWidth = 96
          gblDisplayBindingMode = 0
          gblDisplayBindingId = ""
          gblUseThemeColors = 1
          gblSectionRoundness = 0
          gblReuseNavigation = 1
          gblHardShadowOffset = 0
          gblCounterMode = 0
          gblDensityMode = 0
          gblMinMarkerLength = 4
          gblTitleOverflow = 1
          gblTitleAlignment = 0
          gblTitleCase = 0
          gblFlatRuleMode = 0
          gblMarkerKind = 0
          gblPanelShadow = 0
          gblFlatRuleWidth = 0.5
          gblFlatRuleColor = "107;114;128"
          PI_ClearRuleCache
          PI_InitializeUniversalStyle 0
          If Application.Presentations.Count > 0 Then gblRailWidth = ActivePresentation.PageSetup.SlideWidth


exit_proc:
650       Exit Sub
          
ErrorHandler:
660       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
670       Select Case Err.Number
              Case -2147188160      'no active presentation
680               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "Initialize: Nothing to do!"
690               Resume exit_proc
700               Resume
710           Case Else
720               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "Initialize"
730               Resume Next
740               Resume
              
750       End Select
End Sub

' ----------------------------------------------------------------
' Procedure Name: optionsFromRibbon
' Purpose: Grab new values set from the ribbon, write them to global variables for further use
' Procedure Kind: Function
' Procedure Access: Public
' Parameter strOption (String): which part of the indicator ribbon was clicked (checkbox, button etc.)
' Parameter strValue (String): new value of the corresponding variable
' Return Type: ProgressDotsOptions
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function optionsFromRibbon(strOption As String, strValue As String) As ProgressDotsOptions
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct optionsFromRibbon: " & strOption & " / " & strValue               'Debug: Write active procedure to log
              
          Dim strRGBArray() As String
          Dim s As String
          Dim intCorrected  As Integer
          
30        intCorrected = 0    'no errors so far
          
40        If gblbolDebug = True Then LogInfo "optionsFromRibbon Param 1: " & strOption & ", Value: " & strValue
          
50        Select Case strOption

              Case "ddcGeneralProgressbarStyle"
60                optionsFromRibbon.PIStyle = strValue
70                gblstrIndicatorStyle = optionsFromRibbon.PIStyle

80            Case "ddcGeneralProgressbarEvenSpace"
90                optionsFromRibbon.EvenSpaceSecOrInd = strValue
100               gblEvenSpaceSecOrInd = optionsFromRibbon.EvenSpaceSecOrInd
              
110           Case "cmbMarginLeft"   '"LeftMargin"
120               optionsFromRibbon.LeftMargin = CDbl(strValue)
130               gblintMarginLeft = optionsFromRibbon.LeftMargin
                  
140           Case "cmbMarginRight"   '"RightMargin"
150               optionsFromRibbon.RightMargin = CDbl(strValue)
160               gblintMarginRight = optionsFromRibbon.RightMargin
                  
170           Case "cmbMarginTop"  '"yMargin"
180               optionsFromRibbon.yMargin = CDbl(strValue)
190               gblintMarginTop = optionsFromRibbon.yMargin
                  
200           Case "cmbSectionSpacingSpace"   '"SectionGap"
210               optionsFromRibbon.SectionGap = Val(strValue)
220               gblintSectionSpacing = optionsFromRibbon.SectionGap
                  
230           Case "cmbIndicatorSpaceAdjustment"
                  '            MsgBox "This feature is not working yet", , "Be patient, the future is bright"
240               optionsFromRibbon.IndicatorSpaceAdjustment = Val(strValue)
250               gblintIndicatorSpacingAdjustment = optionsFromRibbon.IndicatorSpaceAdjustment
                  
260           Case "cmbIndicatorSpaceSqueeze"
270               optionsFromRibbon.IndicatorSqueeze = Val(strValue)
280               gblintIndicatorSqueeze = optionsFromRibbon.IndicatorSqueeze
                  
290           Case "cmbSectionIndicatorGap"
300               optionsFromRibbon.SectionIndicatorGap = Val(strValue)
                  If gblSectionBand = 1 And optionsFromRibbon.SectionIndicatorGap < 0 Then optionsFromRibbon.SectionIndicatorGap = 0
310               gblintSectionIndicatorGap = optionsFromRibbon.SectionIndicatorGap
                  
320           Case "cmbFontName"  '"fontName"
330               optionsFromRibbon.fontName = strValue
340               gblstrFontname = optionsFromRibbon.fontName
                  
350           Case "cmbFontSize"  '"textHeight"
                  'error handling first
360               If Val(strValue) < 0 Then   'prevent negative values
370                   strValue = Val(strValue) * -1
380                   intCorrected = 1
390               End If
400               If Val(strValue) = 0 Then
410                   strValue = "12"   'prevent text values
420                   intCorrected = 1
430               End If
                  'set values
440               optionsFromRibbon.textHeight = Val(strValue)
450               gbldblFontSize = optionsFromRibbon.textHeight
                  
                  'Do Colors!
              Case "New color for Band Left Near"
                  optionsFromRibbon.SectionBandLeftNear = DeserializeColor(strValue)
                  gblSectionBandLeftNear = strValue
              Case "New color for Band Left Far"
                  optionsFromRibbon.SectionBandLeftFar = DeserializeColor(strValue)
                  gblSectionBandLeftFar = strValue
              Case "New color for Band Right Near"
                  optionsFromRibbon.SectionBandRightNear = DeserializeColor(strValue)
                  gblSectionBandRightNear = strValue
              Case "New color for Band Right Far"
                  optionsFromRibbon.SectionBandRightFar = DeserializeColor(strValue)
                  gblSectionBandRightFar = strValue
              Case "New color for Band Backdrop"
                  optionsFromRibbon.SectionBandBackdropColor = DeserializeColor(strValue)
                  gblSectionBandBackdropColor = strValue
              Case "New color for Section Accent"
                  optionsFromRibbon.AccentColor = DeserializeColor(strValue)
                  gblAccentColor = strValue
460           Case "New color for Color Font Seen"
470               strRGBArray = Split(strValue, ";")
480               optionsFromRibbon.seenTextColor.red = Val(strRGBArray(0))
490               optionsFromRibbon.seenTextColor.green = Val(strRGBArray(1))
500               optionsFromRibbon.seenTextColor.blue = Val(strRGBArray(2))
510               gblstrFontColorSeen = strValue
              
520           Case "New color for Color Font Unseen"
530               strRGBArray = Split(strValue, ";")
540               optionsFromRibbon.unseenTextColor.red = Val(strRGBArray(0))
550               optionsFromRibbon.unseenTextColor.green = Val(strRGBArray(1))
560               optionsFromRibbon.unseenTextColor.blue = Val(strRGBArray(2))
570               gblstrFontColorUnSeen = strValue
              
580           Case "New color for Color Font Seen Background"
590               strRGBArray = Split(strValue, ";")
600               optionsFromRibbon.seenTextColorBG.red = Val(strRGBArray(0))
610               optionsFromRibbon.seenTextColorBG.green = Val(strRGBArray(1))
620               optionsFromRibbon.seenTextColorBG.blue = Val(strRGBArray(2))
630               gblstrFontBGColorSeen = strValue
              
640           Case "New color for Color Font Unseen Background"
650               strRGBArray = Split(strValue, ";")
660               optionsFromRibbon.unseenTextColorBG.red = Val(strRGBArray(0))
670               optionsFromRibbon.unseenTextColorBG.green = Val(strRGBArray(1))
680               optionsFromRibbon.unseenTextColorBG.blue = Val(strRGBArray(2))
690               gblstrFontBGColorUnSeen = strValue
              
700           Case "New color for Color Indicator Seen"
710               strRGBArray = Split(strValue, ";")
720               optionsFromRibbon.seenDotColor.red = Val(strRGBArray(0))
730               optionsFromRibbon.seenDotColor.green = Val(strRGBArray(1))
740               optionsFromRibbon.seenDotColor.blue = Val(strRGBArray(2))
750               gblstrIndicatorColorSeen = strValue
              
760           Case "New color for Color Indicator Unseen"
770               strRGBArray = Split(strValue, ";")
780               optionsFromRibbon.unseenDotColor.red = Val(strRGBArray(0))
790               optionsFromRibbon.unseenDotColor.green = Val(strRGBArray(1))
800               optionsFromRibbon.unseenDotColor.blue = Val(strRGBArray(2))
810               gblstrIndicatorColorUnSeen = strValue
                  
                  
820           Case "New color for Color Indicator Active"
830               strRGBArray = Split(strValue, ";")
840               optionsFromRibbon.CurrentDotColor.red = Val(strRGBArray(0))
850               optionsFromRibbon.CurrentDotColor.green = Val(strRGBArray(1))
860               optionsFromRibbon.CurrentDotColor.blue = Val(strRGBArray(2))
870               gblstrIndicatorColorCurrent = strValue
                  
                  
880           Case "New color for Color Font Active Background"
890               strRGBArray = Split(strValue, ";")
900               optionsFromRibbon.CurrentTextColorBG.red = Val(strRGBArray(0))
910               optionsFromRibbon.CurrentTextColorBG.green = Val(strRGBArray(1))
920               optionsFromRibbon.CurrentTextColorBG.blue = Val(strRGBArray(2))
930               gblstrFontBGColorCurrent = strValue
                  
                  
940           Case "New color for Color Font Active"
950               strRGBArray = Split(strValue, ";")
960               optionsFromRibbon.CurrentTextColor.red = Val(strRGBArray(0))
970               optionsFromRibbon.CurrentTextColor.green = Val(strRGBArray(1))
980               optionsFromRibbon.CurrentTextColor.blue = Val(strRGBArray(2))
990               gblstrFontColorCurrent = strValue
                                                
                  
1000          Case "cmbIndicatorSize" '"dotSize"
                  'error handling first
1010              If Val(strValue) < 0 Then   'prevent negative values
1020                  strValue = Val(strValue) * -1
1030                  intCorrected = 1
1040              End If
1050              If Val(strValue) = 0 Then
1060                  strValue = "12"   'prevent text values
1070                  intCorrected = 1
1080              End If
                  'set values
1090              optionsFromRibbon.dotSize = Val(strValue)               '* 0.2
1100              gbldblIndicatorSize = optionsFromRibbon.dotSize
                                    
1110          Case "ebx_DrawStart"
1120              If IsNumeric(strValue) Then
1130                  If Val(strValue) > 0 And Val(strValue) < 1000 Then
                          'all ok
1140                  Else
1150                      intCorrected = 1
1160                  End If
1170              Else
1180                  intCorrected = 1
1190              End If
1200              If intCorrected = 1 Then strValue = "1"
1210              gbliDrawStart = Val(strValue)

1220          Case "ebx_DrawEnd"
1230              If IsNumeric(strValue) Then
1240                  If Val(strValue) > 0 And Val(strValue) < 1000 Then
                          'all ok
1250                  Else
1260                      intCorrected = 1
1270                  End If
1280              Else
1290                  intCorrected = 1
1300              End If
1310              If intCorrected = 1 Then strValue = "999"
1320              gbliDrawEnd = Val(strValue)
                  
1330          Case "ddcIndicatorType"
1340              optionsFromRibbon.dotShapeID = CInt(strValue) 'this listitem will be displayed as default / active
1350              optionsFromRibbon.dotShape = retShapeName(CInt(strValue))
1360              gblstrIndicatorType = optionsFromRibbon.dotShape
1370              gblintIndicatorTypeID = optionsFromRibbon.dotShapeID
              
1380          Case "cbxSectionplacement"
1390              optionsFromRibbon.SectionPlacement = CInt(strValue)
1400              gblintSectionPlacement = optionsFromRibbon.SectionPlacement
              
1410          Case "cbxFontColorUseBG"
1420              optionsFromRibbon.useTextColorBG = CInt(strValue)
1430              gblintUseTextBGCol = optionsFromRibbon.useTextColorBG
          
1440          Case "cbxFontColorUseBG_Current"
1450              optionsFromRibbon.useTextColorBG_Current = CInt(strValue)
1460              gblintUseTextBGCol_Current = optionsFromRibbon.useTextColorBG_Current
          
1470          Case "cbxFontColorUseBG_UnSeen"
1480              optionsFromRibbon.useTextColorBG_UnSeen = CInt(strValue)
1490              gblintUseTextBGCol_UnSeen = optionsFromRibbon.useTextColorBG_UnSeen

1500          Case "cmbFontTransparencySeen"
1510              optionsFromRibbon.transparencyFontSeen = Val(strValue)
1520              gblTransparencyFontSeen = optionsFromRibbon.transparencyFontSeen
                    
1530          Case "cmbFontTransparencyActive"
1540              optionsFromRibbon.transparencyFontActive = Val(strValue)
1550              gblTransparencyFontActive = optionsFromRibbon.transparencyFontActive
               
1560          Case "cmbFontTransparencyUnseen"
1570              optionsFromRibbon.transparencyFontUnseen = Val(strValue)
1580              gblTransparencyFontUnseen = optionsFromRibbon.transparencyFontUnseen
          
1590          Case "cbxIndicatorFill_Seen"
1600              optionsFromRibbon.useShapeColorBG_Seen = CInt(strValue)
1610              gblintUseShapeBG_Seen = optionsFromRibbon.useShapeColorBG_Seen
          
1620          Case "cbxIndicatorFill_Current"
1630              optionsFromRibbon.useShapeColorBG_Current = CInt(strValue)
1640              gblintUseShapeBG_Current = optionsFromRibbon.useShapeColorBG_Current
          
1650          Case "cbxIndicatorFill_UnSeen"
1660              optionsFromRibbon.useShapeColorBG_UnSeen = CInt(strValue)
1670              gblintUseShapeBG_UnSeen = optionsFromRibbon.useShapeColorBG_UnSeen
                    
1680          Case "cmbIndicatorTransparencySeen"
1690              optionsFromRibbon.transparencyIndicatorSeen = Val(strValue)
1700              gblTransparencyIndicatorSeen = optionsFromRibbon.transparencyIndicatorSeen
                    
1710          Case "cmbIndicatorTransparencyActive"
1720              optionsFromRibbon.transparencyIndicatorActive = Val(strValue)
1730              gblTransparencyIndicatorActive = optionsFromRibbon.transparencyIndicatorActive
               
1740          Case "cmbIndicatorTransparencyUnseen"
1750              optionsFromRibbon.transparencyIndicatorUnseen = Val(strValue)
1760              gblTransparencyIndicatorUnseen = optionsFromRibbon.transparencyIndicatorUnseen
                    
1770          Case "cbxGeneralProgressbarDirect"
1780              optionsFromRibbon.DirectUpdate = CInt(strValue)
1790              gblintGeneralProgressbarDirect = optionsFromRibbon.DirectUpdate
              
1800          Case "cbxIgnoreIgnored"
1810              optionsFromRibbon.ignoreSkipped = CInt(strValue)
1820              gblintSlideIgnoreSkipped = optionsFromRibbon.ignoreSkipped
                  
1830          Case "cbxIgnoreHidden"
1840              optionsFromRibbon.ignoreHidden = CInt(strValue)
1850              gblintSlideIgnoreHidden = optionsFromRibbon.ignoreHidden
                  
                  '          Case "cbxIgnoreIndicator"
                  '              optionsFromRibbon.ignoreIndicator = CInt(strValue)
                  '              gblstrSlideIgnoreIndicator = optionsFromRibbon.ignoreIndicator
1860          Case "ddcIgnoreIndicator"
1870              optionsFromRibbon.ignoreIndicator = strValue
1880              gblstrSlideIgnoreIndicator = optionsFromRibbon.ignoreIndicator
                  
                  
                  '          Case "cbxIgnoreSections"
                  '              optionsFromRibbon.ignoreSections = CInt(strValue)
                  '              gblstrSlideIgnoreSections = optionsFromRibbon.ignoreSections
                                    
1890          Case "ddcIgnoreSections"
1900              optionsFromRibbon.ignoreSections = strValue
1910              gblstrSlideIgnoreSections = optionsFromRibbon.ignoreSections
                                    
                                    
1920          Case "btn_OutlinePlus"
1930              optionsFromRibbon.ShapeLineWeight = CInt(strValue)
1940              gblShapeLineWeight = optionsFromRibbon.ShapeLineWeight

1950          Case "btn_OutlineMinus"
1960              optionsFromRibbon.ShapeLineWeight = CInt(strValue)
1970              gblShapeLineWeight = optionsFromRibbon.ShapeLineWeight

1980          Case "btn_Style"
1990              optionsFromRibbon.ShapeLineStyle = CInt(strValue)
2000              gblShapeLineStyle = optionsFromRibbon.ShapeLineStyle

2010          Case "btn_Dash"
2020              optionsFromRibbon.ShapeLineDashStyle = CInt(strValue)
2030              gblShapeLineDashStyle = optionsFromRibbon.ShapeLineDashStyle
                                             
                  
2040          Case Else
2050              MsgBox PI_Zh("914D7F6E503C65E06CD58BC6522BFF1A") & vbCrLf & PI_Zh("914D7F6E9879FF1A") & strOption & vbCrLf & PI_Zh("914D7F6E503CFF1A") & strValue, , PI_Zh("914D7F6E503C65E06548")
2060      End Select
              
2070      If gblbolDebug = True Then LogInfo "Saving '" & strOption & "' option '" & strValue & "' from ribbon..."
2080      s = SerializeProgressOptionsRibbon
2090      If gblbolDebug = True Then LogInfo "New: " & s

2100      PersistPISettings s
          
2110      If intCorrected = 1 Then
2120          MsgBox PI_Zh("521A8F9351657684503C65E06548FF0C63D24EF65DF28C0365748BE5503C30028BF768C067E58BBE7F6E3002"), vbOKOnly Or vbExclamation Or vbSystemModal, PI_Zh("8BBE7F6E5DF28C036574")
2130      End If
          
exit_proc:
2140      Exit Function
                   
          
ErrorHandler:
2150      If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
2160      Select Case Err.Number

              Case -2147188160      'no active presentation
2170              MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "optionsFromRibbon: Nothing to do!"
2180              Resume exit_proc
2190              Resume
                
2200          Case 13
2210              MsgBox PI_Zh("8F935165503C65E065483002") & vbCrLf & PI_Zh("63D24EF65DF291CD7F6E8BE598794EE56062590D8FD0884C3002") & vbCrLf & PI_Zh("8BF791CD65B08F93516567096548503C3002"), vbOKOnly Or vbCritical Or vbSystemModal, PI_Zh("8F93516565E06548")
2220              Resume Next
              
2230          Case Else
2240              MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "optionsFromRibbon"
2250              Resume Next
2260              Resume
              
2270      End Select
          

End Function


' ----------------------------------------------------------------
' Procedure Name: optionsToRibbon
' Purpose: reflect new values (e.g. font name / color) to ribbon
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter theOptions (ProgressDotsOptions): which field in the ribbon (and/or global variable) needs to be updated
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub optionsToRibbon(theOptions As ProgressDotsOptions)
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Sub optionsToRibbon (theOptions As ProgressDotsOptions)"         'Debug: Write active procedure to log
          
30        gblintGeneralProgressbarDirect = CInt(theOptions.DirectUpdate)
          
40        gblstrIndicatorStyle = theOptions.PIStyle
          
50        gblstrFontname = theOptions.fontName
60        gbldblFontSize = CDbl(theOptions.textHeight)
70        gblstrFontColorSeen = theOptions.seenTextColor.red & ";" & theOptions.seenTextColor.green & ";" & theOptions.seenTextColor.blue
80        gblstrFontColorUnSeen = theOptions.unseenTextColor.red & ";" & theOptions.unseenTextColor.green & ";" & theOptions.unseenTextColor.blue
          
90        gblstrIndicatorType = theOptions.dotShape
100       gblintIndicatorTypeID = theOptions.dotShapeID
110       gbldblIndicatorSize = CDbl(theOptions.dotSize)
120       gblstrIndicatorColorSeen = theOptions.seenDotColor.red & ";" & theOptions.seenDotColor.green & ";" & theOptions.seenDotColor.blue
130       gblstrIndicatorColorUnSeen = theOptions.unseenDotColor.red & ";" & theOptions.unseenDotColor.green & ";" & theOptions.unseenDotColor.blue

140       gblintMarginTop = CInt(theOptions.yMargin)
150       gblintMarginLeft = CInt(theOptions.LeftMargin)
160       gblintMarginRight = CInt(theOptions.RightMargin)
          
170       gblintSectionSpacing = theOptions.SectionGap
          
180       gblintSlideIgnoreSkipped = CInt(theOptions.ignoreSkipped)
190       gblintSlideIgnoreHidden = CInt(theOptions.ignoreHidden)
200       gblstrSlideIgnoreSections = theOptions.ignoreSections
          
210       gblintUseTextBGCol = CInt(theOptions.useTextColorBG)
220       gblintUseTextBGCol_Current = CInt(theOptions.useTextColorBG_Current)
230       gblintUseTextBGCol_UnSeen = CInt(theOptions.useTextColorBG_UnSeen)

240       gblintUseShapeBG_Seen = CInt(theOptions.useShapeColorBG_Seen)
250       gblintUseShapeBG_Current = CInt(theOptions.useShapeColorBG_Current)
260       gblintUseShapeBG_UnSeen = CInt(theOptions.useShapeColorBG_UnSeen)

270       gblstrFontBGColorSeen = theOptions.seenTextColorBG.red & ";" & theOptions.seenTextColorBG.green & ";" & theOptions.seenTextColorBG.blue
280       gblstrFontBGColorUnSeen = theOptions.unseenTextColorBG.red & ";" & theOptions.unseenTextColorBG.green & ";" & theOptions.unseenTextColorBG.blue

290       gblintIndicatorSqueeze = CInt(theOptions.IndicatorSqueeze)
          
300       gblintSectionPlacement = CInt(theOptions.SectionPlacement)
          
310       gblintSectionIndicatorGap = CInt(theOptions.SectionIndicatorGap)
          
320       gblstrIndicatorColorCurrent = theOptions.CurrentDotColor.red & ";" & theOptions.CurrentDotColor.green & ";" & theOptions.CurrentDotColor.blue
330       gblstrFontBGColorCurrent = theOptions.CurrentTextColorBG.red & ";" & theOptions.CurrentTextColorBG.green & ";" & theOptions.CurrentTextColorBG.blue
340       gblstrFontColorCurrent = theOptions.CurrentTextColor.red & ";" & theOptions.CurrentTextColor.green & ";" & theOptions.CurrentTextColor.blue
          
350       gblShapeLineWeight = CInt(theOptions.ShapeLineWeight)
360       gblShapeLineStyle = CInt(theOptions.ShapeLineStyle)
370       gblShapeLineDashStyle = CInt(theOptions.ShapeLineDashStyle)
          
380       gblstrSlideIgnoreIndicator = theOptions.ignoreIndicator

390       gblTransparencyIndicatorSeen = CInt(theOptions.transparencyIndicatorSeen)
400       gblTransparencyIndicatorActive = CInt(theOptions.transparencyIndicatorActive)
410       gblTransparencyIndicatorUnseen = CInt(theOptions.transparencyIndicatorUnseen)
          
420       gblTransparencyFontSeen = CInt(theOptions.transparencyFontSeen)
430       gblTransparencyFontActive = CInt(theOptions.transparencyFontActive)
440       gblTransparencyFontUnseen = CInt(theOptions.transparencyFontUnseen)

450       gblEvenSpaceSecOrInd = theOptions.EvenSpaceSecOrInd
          gblSectionBand = theOptions.SectionBand
          gblSectionBandPadding = theOptions.SectionBandPadding
          gblSectionBandContrast = theOptions.SectionBandContrast
          gblSectionBandGradient = theOptions.SectionBandGradient
          gblSectionBandTransparencyCurrent = theOptions.SectionBandTransparencyCurrent
          gblSectionBandTransparencyLeft = theOptions.SectionBandTransparencyLeft
          gblSectionBandTransparencyRight = theOptions.SectionBandTransparencyRight
          gblSectionBandLeftNear = SerializeColor(theOptions.SectionBandLeftNear)
          gblSectionBandLeftFar = SerializeColor(theOptions.SectionBandLeftFar)
          gblSectionBandRightNear = SerializeColor(theOptions.SectionBandRightNear)
          gblSectionBandRightFar = SerializeColor(theOptions.SectionBandRightFar)
          gblSectionBandBackdrop = theOptions.SectionBandBackdrop
          gblSectionBandBackdropColor = SerializeColor(theOptions.SectionBandBackdropColor)
          gblSectionBandBackdropTransparency = theOptions.SectionBandBackdropTransparency
          gblSectionBandBackdropScale = theOptions.SectionBandBackdropScale
          gblFlatBar = theOptions.FlatBar
          gblFlatBarHeight = theOptions.FlatBarHeight
          gblFlatBarGap = theOptions.FlatBarGap
          gblFlatBarRoundness = theOptions.FlatBarRoundness
          gblFlatBarGlow = theOptions.FlatBarGlow
          gblFlatBarGlowRadius = theOptions.FlatBarGlowRadius
          gblFlatBarGlowTransparency = theOptions.FlatBarGlowTransparency
          gblLayeredNavigation = theOptions.LayeredNavigation
          gblNavigationWidth = theOptions.NavigationWidth
          gblSectionBlockHeight = theOptions.SectionBlockHeight
          gblRailWidth = theOptions.RailWidth
          gblRailThickness = theOptions.RailThickness
          gblLayerShadowBlur = theOptions.LayerShadowBlur
          gblLayerShadowTransparency = theOptions.LayerShadowTransparency
          gblLayerDepthStep = theOptions.LayerDepthStep
          gblLayerShadowDistance = theOptions.LayerShadowDistance
          gblAccentMode = theOptions.AccentMode
          gblAccentWidth = theOptions.AccentWidth
          gblAccentColor = SerializeColor(theOptions.AccentColor)
          gblSectionNumbering = theOptions.SectionNumbering
          gblCurrentTitleBold = theOptions.CurrentTitleBold
          gblFutureBarOutline = theOptions.FutureBarOutline
          gblSectionForegroundBars = theOptions.SectionForegroundBars
          gblDockEdge = theOptions.DockEdge
          gblSidebarWidth = theOptions.SidebarWidth
          gblDisplayBindingMode = theOptions.DisplayBindingMode
          gblDisplayBindingId = theOptions.DisplayBindingId
          gblUseThemeColors = theOptions.UseThemeColors
          gblSectionRoundness = theOptions.SectionRoundness
          gblReuseNavigation = theOptions.ReuseNavigation
          gblHardShadowOffset = theOptions.HardShadowOffset
          gblCounterMode = theOptions.CounterMode
          gblDensityMode = theOptions.DensityMode
          gblMinMarkerLength = theOptions.MinMarkerLength
          gblTitleOverflow = theOptions.TitleOverflow
          gblTitleAlignment = theOptions.TitleAlignment
          gblTitleCase = theOptions.TitleCase
          gblMarkerKind = theOptions.MarkerKind
          gblPanelShadow = theOptions.PanelShadow
          gblPresetSize = theOptions.PresetSize
          gblCrimsonType = theOptions.CrimsonType
          gblInactiveSectionBarTransparency = theOptions.InactiveSectionBarTransparency
          gblBandHeightMode = theOptions.BandHeightMode
          gblFlatRuleMode = theOptions.FlatRuleMode
          gblFlatRuleWidth = theOptions.FlatRuleWidth
          gblFlatRuleColor = SerializeColor(theOptions.FlatRuleColor)
          If Not gobjRibbon Is Nothing Then
              gobjRibbon.InvalidateControl "cbxSectionBand"
              gobjRibbon.InvalidateControl "cbxSectionBandContrast"
              gobjRibbon.InvalidateControl "cmbSectionBandPadding"
              gobjRibbon.InvalidateControl "cbxSectionBandGradient"
              gobjRibbon.InvalidateControl "edtBandTransparencyCurrent"
              gobjRibbon.InvalidateControl "edtBandTransparencyLeft"
              gobjRibbon.InvalidateControl "edtBandTransparencyRight"
              gobjRibbon.InvalidateControl "cbxBandBackdrop"
              gobjRibbon.InvalidateControl "edtBandBackdropTransparency"
              gobjRibbon.InvalidateControl "edtBandBackdropScale"
              gobjRibbon.InvalidateControl "cbxFlatBar"
              gobjRibbon.InvalidateControl "cbxFlatBarGlow"
              gobjRibbon.InvalidateControl "edtFlatBarHeight"
              gobjRibbon.InvalidateControl "edtFlatBarGap"
              gobjRibbon.InvalidateControl "edtFlatBarRoundness"
              gobjRibbon.InvalidateControl "edtFlatBarGlowRadius"
              gobjRibbon.InvalidateControl "edtFlatBarGlowTransparency"
              gobjRibbon.InvalidateControl "cbxLayeredNavigation"
              gobjRibbon.InvalidateControl "edtNavigationWidth"
              gobjRibbon.InvalidateControl "edtSectionBlockHeight"
              gobjRibbon.InvalidateControl "edtRailWidth"
              gobjRibbon.InvalidateControl "edtRailThickness"
              gobjRibbon.InvalidateControl "edtLayerShadowBlur"
              gobjRibbon.InvalidateControl "edtLayerShadowTransparency"
              gobjRibbon.InvalidateControl "edtLayerDepthStep"
              gobjRibbon.InvalidateControl "edtLayerShadowDistance"
              gobjRibbon.InvalidateControl "ddAccentMode"
              gobjRibbon.InvalidateControl "edtAccentWidth"
              gobjRibbon.InvalidateControl "cbxSectionNumbering"
              gobjRibbon.InvalidateControl "cbxCurrentTitleBold"
              gobjRibbon.InvalidateControl "cbxFutureBarOutline"
              gobjRibbon.InvalidateControl "cbxSectionForegroundBars"
              gobjRibbon.InvalidateControl "ddDockEdge"
              gobjRibbon.InvalidateControl "edtSidebarWidth"
              gobjRibbon.InvalidateControl "lblDisplayBinding"
              gobjRibbon.InvalidateControl "cbxUseThemeColors"
              gobjRibbon.InvalidateControl "edtSectionRoundness"
              gobjRibbon.InvalidateControl "cbxReuseNavigation"
              gobjRibbon.InvalidateControl "edtHardShadowOffset"
              gobjRibbon.InvalidateControl "ddCounterMode"
              gobjRibbon.InvalidateControl "ddDensityMode"
              gobjRibbon.InvalidateControl "edtMinMarkerLength"
              gobjRibbon.InvalidateControl "ddTitleOverflow"
              gobjRibbon.InvalidateControl "ddTitleAlignment"
              gobjRibbon.InvalidateControl "ddTitleCase"
              gobjRibbon.InvalidateControl "ddMarkerKind"
              gobjRibbon.InvalidateControl "cbxPanelShadow"
              gobjRibbon.InvalidateControl "ddFlatRuleMode"
              gobjRibbon.InvalidateControl "edtFlatRuleWidth"
              gobjRibbon.InvalidateControl "edtFlatRuleColor"
              gobjRibbon.InvalidateControl "mnuDocumentRules"
              gobjRibbon.InvalidateControl "mnuDocumentPresets"
          End If
          
exit_proc:
460       Exit Sub
          
ErrorHandler:
          If PI_IsSilentGeneration() Then PI_RaiseNativeFailure "optionsToRibbon", Erl, Err.Number, Err.Source, Err.Description
470       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
480       Select Case Err.Number
              Case -2147188160      'no active presentation
490               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "optionsToRibbon: Nothing to do!"
500               Resume exit_proc
510               Resume
520           Case Else
530               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "optionsToRibbon"
540               Resume Next
550               Resume
              
560       End Select


End Sub

' ----------------------------------------------------------------
' Procedure Name: SerializeProgressOptions
' Purpose: Old method, kept for reuse of old code.
'           New "SerializeProgressOptionsRibbon".
'           Same function: Combine different values to one string
'               (which will be saved into the presentation).
'               Values seperated by space.
'               Space in Fontname will be replaced by #
' Procedure Kind: Function
' Procedure Access: Public
' Parameter theOptions (ProgressDotsOptions): which option to be set
' Return Type: String
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function SerializeProgressOptions(theOptions As ProgressDotsOptions) As String
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct SerializeProgressOptions (theOptions As ProgressDotsOptions)"         'Debug: Write active procedure to log
              
30        With theOptions:
40            SerializeProgressOptions = gblIndicatorSettingsVersion & " " & _
                  .LeftMargin & " " & .RightMargin & " " & _
                  .yMargin & " " & .textHeight & " " & _
                  .dotSize & " " & .SectionGap & " " & _
                  SerializeColor(.seenDotColor) & " " & SerializeColor(.unseenDotColor) & " " & _
                  SerializeColor(.seenTextColor) & " " & SerializeColor(.unseenTextColor) & " " & _
                  Replace(.fontName, " ", "#") & " " & _
                  .dotShape & " " & .dotShapeID & " " & _
                  .ignoreSkipped & " " & .ignoreHidden & " " & _
                  Replace(.ignoreSections, " ", "#") & " " & .DirectUpdate & " " & _
                  .IndicatorSpaceAdjustment & " " & _
                  .useTextColorBG & " " & .useTextColorBG_Current & " " & .useTextColorBG_UnSeen & " " & _
                  .useShapeColorBG_Seen & " " & .useShapeColorBG_Current & " " & .useShapeColorBG_UnSeen & " " & _
                  SerializeColor(.seenTextColorBG) & " " & SerializeColor(.unseenTextColorBG) & " " & _
                  .IndicatorSqueeze & " " & .SectionPlacement & " " & .SectionIndicatorGap & " " & _
                  SerializeColor(.CurrentDotColor) & " " & SerializeColor(.CurrentTextColorBG) & " " & SerializeColor(.CurrentTextColor) & " " & _
                  .ShapeLineWeight & " " & .ShapeLineStyle & " " & .ShapeLineDashStyle & " " & _
                  Replace(.ignoreIndicator, " ", "#") & " " & _
                  .transparencyIndicatorSeen & " " & .transparencyIndicatorActive & " " & .transparencyIndicatorUnseen & " " & _
                  .transparencyFontSeen & " " & .transparencyFontActive & " " & .transparencyFontUnseen & " " & _
                  Replace(.EvenSpaceSecOrInd, " ", "#") & " " & _
                  Replace(.PIStyle, " ", "#") & " " & _
                  .SectionBand & " " & Trim$(Str$(.SectionBandPadding)) & " " & .SectionBandContrast
              SerializeProgressOptions = SerializeProgressOptions & " " & .SectionBandGradient & " " & _
                  Trim$(Str$(.SectionBandTransparencyCurrent)) & " " & Trim$(Str$(.SectionBandTransparencyLeft)) & " " & Trim$(Str$(.SectionBandTransparencyRight))
              SerializeProgressOptions = SerializeProgressOptions & " " & SerializeColor(.SectionBandLeftNear) & " " & _
                  SerializeColor(.SectionBandLeftFar) & " " & SerializeColor(.SectionBandRightNear) & " " & SerializeColor(.SectionBandRightFar)
              SerializeProgressOptions = SerializeProgressOptions & " " & .SectionBandBackdrop & " " & _
                  SerializeColor(.SectionBandBackdropColor) & " " & Trim$(Str$(.SectionBandBackdropTransparency)) & " " & Trim$(Str$(.SectionBandBackdropScale))
              SerializeProgressOptions = SerializeProgressOptions & " " & .FlatBar & " " & _
                  Trim$(Str$(.FlatBarHeight)) & " " & Trim$(Str$(.FlatBarGap)) & " " & Trim$(Str$(.FlatBarRoundness))
              SerializeProgressOptions = SerializeProgressOptions & " " & .FlatBarGlow & " " & _
                  Trim$(Str$(.FlatBarGlowRadius)) & " " & Trim$(Str$(.FlatBarGlowTransparency))
              SerializeProgressOptions = SerializeProgressOptions & " " & .LayeredNavigation & " " & _
                  Trim$(Str$(.NavigationWidth)) & " " & Trim$(Str$(.SectionBlockHeight)) & " " & Trim$(Str$(.RailWidth))
              SerializeProgressOptions = SerializeProgressOptions & " " & Trim$(Str$(.RailThickness)) & " " & _
                  Trim$(Str$(.LayerShadowBlur)) & " " & Trim$(Str$(.LayerShadowTransparency)) & " " & Trim$(Str$(.LayerDepthStep))
              SerializeProgressOptions = SerializeProgressOptions & " " & Trim$(Str$(.LayerShadowDistance))
              SerializeProgressOptions = SerializeProgressOptions & " " & .AccentMode & " " & _
                  Trim$(Str$(.AccentWidth)) & " " & SerializeColor(.AccentColor) & " " & _
                  .SectionNumbering & " " & .CurrentTitleBold & " " & .FutureBarOutline
              SerializeProgressOptions = SerializeProgressOptions & " " & .SectionForegroundBars
              SerializeProgressOptions = SerializeProgressOptions & " " & .DockEdge & " " & Trim$(Str$(.SidebarWidth)) & " " & _
                  .DisplayBindingMode & " " & PI_BindingToken(.DisplayBindingId) & " " & .UseThemeColors & " " & _
                  Trim$(Str$(.SectionRoundness)) & " " & .ReuseNavigation & " " & Trim$(Str$(.HardShadowOffset))
              SerializeProgressOptions = SerializeProgressOptions & " " & .CounterMode & " " & .DensityMode & " " & _
                  Trim$(Str$(.MinMarkerLength)) & " " & .TitleOverflow & " " & .TitleAlignment & " " & .TitleCase & " " & _
                  .FlatRuleMode & " " & Trim$(Str$(.FlatRuleWidth)) & " " & SerializeColor(.FlatRuleColor) & " " & .MarkerKind & " " & .PanelShadow
              SerializeProgressOptions = SerializeProgressOptions & " " & .PresetSize & " " & .CrimsonType & " " & .BandHeightMode & _
                      " " & .InactiveSectionBarTransparency & " " & PI_STYLE_CONFIG_REVISION

                                  
50        End With

exit_proc:
60        Exit Function

ErrorHandler:
70        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
80        Select Case Err.Number
              Case -2147188160      'no active presentation
90                MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "SerializeProgressOptions: Nothing to do!"
100               Resume exit_proc
110               Resume
120           Case Else
130               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "SerializeProgressOptions"
140               Resume Next
150               Resume

160       End Select

End Function

' ----------------------------------------------------------------
' Procedure Name: SerializeProgressOptionsRibbon
' Purpose: Combine different values to one string
'           (which will be saved into the presentation).
'           Values seperated by space.
'           Space in Fontname will be replaced by #
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: String
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function SerializeProgressOptionsRibbon() As String
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct SerializeProgressOptionsRibbon"         'Debug: Write active procedure to log
          
30        SerializeProgressOptionsRibbon = gblIndicatorSettingsVersion & " " & _
              gblintMarginLeft & " " & gblintMarginRight & " " & gblintMarginTop & _
              " " & gbldblFontSize & " " & gbldblIndicatorSize & " " & _
              gblintSectionSpacing & " " & gblstrIndicatorColorSeen & " " & _
              gblstrIndicatorColorUnSeen & " " & gblstrFontColorSeen & " " & _
              gblstrFontColorUnSeen & " " & Replace(gblstrFontname, " ", "#") & " " & _
              gblstrIndicatorType & " " & gblintIndicatorTypeID & " " & _
              gblintSlideIgnoreSkipped & " " & gblintSlideIgnoreHidden & " " & _
              Replace(gblstrSlideIgnoreSections, " ", "#") & " " & gblintGeneralProgressbarDirect & " " _
              & gblintIndicatorSpacingAdjustment & " " & _
              gblintUseTextBGCol & " " & gblintUseTextBGCol_Current & " " & gblintUseTextBGCol_UnSeen & " " & _
              gblintUseShapeBG_Seen & " " & gblintUseShapeBG_Current & " " & gblintUseShapeBG_UnSeen & " " & _
              gblstrFontBGColorSeen & " " & gblstrFontBGColorUnSeen & " " & _
              gblintIndicatorSqueeze & " " & gblintSectionPlacement & " " & _
              gblintSectionIndicatorGap & " " & gblstrIndicatorColorCurrent & " " & _
              gblstrFontBGColorCurrent & " " & gblstrFontColorCurrent & " " & _
              gblShapeLineWeight & " " & gblShapeLineStyle & " " & gblShapeLineDashStyle & " " & _
              Replace(gblstrSlideIgnoreIndicator, " ", "#") & " " & _
              gblTransparencyIndicatorSeen & " " & gblTransparencyIndicatorActive & " " & gblTransparencyIndicatorUnseen & " " & _
              gblTransparencyFontSeen & " " & gblTransparencyFontActive & " " & gblTransparencyFontUnseen & " " & _
              Replace(gblEvenSpaceSecOrInd, " ", "#") & _
              " " & Replace(gblstrIndicatorStyle, " ", "#") & " " & _
              gblSectionBand & " " & Trim$(Str$(gblSectionBandPadding)) & " " & gblSectionBandContrast
          SerializeProgressOptionsRibbon = SerializeProgressOptionsRibbon & " " & gblSectionBandGradient & " " & _
              Trim$(Str$(gblSectionBandTransparencyCurrent)) & " " & Trim$(Str$(gblSectionBandTransparencyLeft)) & " " & Trim$(Str$(gblSectionBandTransparencyRight))
          SerializeProgressOptionsRibbon = SerializeProgressOptionsRibbon & " " & gblSectionBandLeftNear & " " & _
              gblSectionBandLeftFar & " " & gblSectionBandRightNear & " " & gblSectionBandRightFar
          SerializeProgressOptionsRibbon = SerializeProgressOptionsRibbon & " " & gblSectionBandBackdrop & " " & _
              gblSectionBandBackdropColor & " " & Trim$(Str$(gblSectionBandBackdropTransparency)) & " " & Trim$(Str$(gblSectionBandBackdropScale))
          SerializeProgressOptionsRibbon = SerializeProgressOptionsRibbon & " " & gblFlatBar & " " & _
              Trim$(Str$(gblFlatBarHeight)) & " " & Trim$(Str$(gblFlatBarGap)) & " " & Trim$(Str$(gblFlatBarRoundness))
          SerializeProgressOptionsRibbon = SerializeProgressOptionsRibbon & " " & gblFlatBarGlow & " " & _
              Trim$(Str$(gblFlatBarGlowRadius)) & " " & Trim$(Str$(gblFlatBarGlowTransparency))
          SerializeProgressOptionsRibbon = SerializeProgressOptionsRibbon & " " & gblLayeredNavigation & " " & _
              Trim$(Str$(gblNavigationWidth)) & " " & Trim$(Str$(gblSectionBlockHeight)) & " " & Trim$(Str$(gblRailWidth))
          SerializeProgressOptionsRibbon = SerializeProgressOptionsRibbon & " " & Trim$(Str$(gblRailThickness)) & " " & _
              Trim$(Str$(gblLayerShadowBlur)) & " " & Trim$(Str$(gblLayerShadowTransparency)) & " " & Trim$(Str$(gblLayerDepthStep))
          SerializeProgressOptionsRibbon = SerializeProgressOptionsRibbon & " " & Trim$(Str$(gblLayerShadowDistance))
          SerializeProgressOptionsRibbon = SerializeProgressOptionsRibbon & " " & gblAccentMode & " " & _
              Trim$(Str$(gblAccentWidth)) & " " & gblAccentColor & " " & _
              gblSectionNumbering & " " & gblCurrentTitleBold & " " & gblFutureBarOutline
          SerializeProgressOptionsRibbon = SerializeProgressOptionsRibbon & " " & gblSectionForegroundBars
          SerializeProgressOptionsRibbon = SerializeProgressOptionsRibbon & " " & gblDockEdge & " " & Trim$(Str$(gblSidebarWidth)) & " " & _
              gblDisplayBindingMode & " " & PI_BindingToken(gblDisplayBindingId) & " " & gblUseThemeColors & " " & _
              Trim$(Str$(gblSectionRoundness)) & " " & gblReuseNavigation & " " & Trim$(Str$(gblHardShadowOffset))
          SerializeProgressOptionsRibbon = SerializeProgressOptionsRibbon & " " & gblCounterMode & " " & gblDensityMode & " " & _
              Trim$(Str$(gblMinMarkerLength)) & " " & gblTitleOverflow & " " & gblTitleAlignment & " " & gblTitleCase & " " & _
              gblFlatRuleMode & " " & Trim$(Str$(gblFlatRuleWidth)) & " " & gblFlatRuleColor & " " & gblMarkerKind & " " & gblPanelShadow
          SerializeProgressOptionsRibbon = SerializeProgressOptionsRibbon & " " & gblPresetSize & " " & gblCrimsonType & " " & gblBandHeightMode & _
              " " & gblInactiveSectionBarTransparency & " " & PI_STYLE_CONFIG_REVISION

exit_proc:
40        Exit Function

ErrorHandler:
50        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
60        Select Case Err.Number
              Case -2147188160      'no active presentation
70                MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "SerializeProgressOptionsRibbon: Nothing to do!"
80                Resume exit_proc
90                Resume
100           Case Else
110               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "SerializeProgressOptionsRibbon"
120               Resume Next
130               Resume

140       End Select

End Function

