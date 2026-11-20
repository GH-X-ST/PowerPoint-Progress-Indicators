Attribute VB_Name = "basGeneral"
Option Explicit

Public gblSectionBand As Integer
Public gblSectionBandPadding As Double
Public gblSectionBandContrast As Integer
Public gblSectionBandGradient As Integer
Public gblSectionBandTransparencyCurrent As Double
Public gblSectionBandTransparencyLeft As Double
Public gblSectionBandTransparencyRight As Double
Public gblSectionBandLeftNear As String
Public gblSectionBandLeftFar As String
Public gblSectionBandRightNear As String
Public gblSectionBandRightFar As String
Public gblSectionBandBackdrop As Integer
Public gblSectionBandBackdropColor As String
Public gblSectionBandBackdropTransparency As Double
Public gblSectionBandBackdropScale As Double
Public gblFlatBar As Integer
Public gblFlatBarHeight As Double
Public gblFlatBarGap As Double
Public gblFlatBarRoundness As Double
Public gblFlatBarGlow As Integer
Public gblFlatBarGlowRadius As Double
Public gblFlatBarGlowTransparency As Double
Public gblLayeredNavigation As Integer
Public gblNavigationWidth As Double
Public gblSectionBlockHeight As Double
Public gblRailWidth As Double
Public gblRailThickness As Double
Public gblLayerShadowBlur As Double
Public gblLayerShadowTransparency As Double
Public gblLayerDepthStep As Double
Public gblLayerShadowDistance As Double
Public gblAccentMode As Integer
Public gblAccentWidth As Double
Public gblAccentColor As String
Public gblSectionNumbering As Integer
Public gblCurrentTitleBold As Integer
Public gblFutureBarOutline As Integer
Public gblSectionForegroundBars As Integer
Public gblDockEdge As Integer
Public gblSidebarWidth As Double
Public gblDisplayBindingMode As Integer
Public gblDisplayBindingId As String
Public gblUseThemeColors As Integer
Public gblSectionRoundness As Double
Public gblReuseNavigation As Integer
Public gblHardShadowOffset As Double
Public gblCounterMode As Integer
Public gblDensityMode As Integer
Public gblMinMarkerLength As Double
Public gblTitleOverflow As Integer
Public gblTitleAlignment As Integer
Public gblTitleCase As Integer
Public gblMarkerKind As Integer
Public gblPanelShadow As Integer
Public gblPresetSize As Integer
Public gblCrimsonType As Integer
Public gblInactiveSectionBarTransparency As Integer
Public Const PI_STYLE_CONFIG_REVISION As Integer = 2
Public gblBandHeightMode As Integer
Public gblFlatRuleMode As Integer
Public gblFlatRuleWidth As Double
Public gblFlatRuleColor As String
Private mPIGenerationIds As String
Private Const PI_BINDING_REGISTRY As String = "PI_BINDING_REGISTRY"
Private Const PI_BINDING_NS As String = "urn:progress-indicator:binding:v1"
Private Const PI_DISPLAY_HIDDEN As String = "PI_DISPLAY_HIDDEN"
Private Const PI_RULES_NS As String = "urn:progress-indicator:display-rules:v1"
Private Const PI_PRESETS_NS As String = "urn:progress-indicator:user-presets:v1"
Private Const PI_DOCUMENT_LIMIT As Long = 32
Private Type PIStoredRule
    title As String
    bindingMode As Integer
    bindingId As String
    show As Boolean
    config As String
End Type
Private Type PIStoredPreset
    title As String
    config As String
End Type
Private mPIRules(1 To PI_DOCUMENT_LIMIT) As PIStoredRule
Private mPIRuleOptions(1 To PI_DOCUMENT_LIMIT) As ProgressDotsOptions
Private mPIRuleCount As Long
Private mPIRulesEnabled As Boolean
Private mPIRulesDefaultShow As Boolean
Private mPIRulesPresentation As Presentation
Private mPIRulesLoaded As Boolean
Private mPIRulesValidated As Boolean
Private mPIRuleLayouts() As CustomLayout
Private mPILayoutMatches() As Long
Private mPIRuleLayoutCount As Long
Private mPIPresets(1 To PI_DOCUMENT_LIMIT) As PIStoredPreset
Private mPIPresetCount As Long
Private mPIPresetMenuPresentation As Presentation
Private mPIRuleMenuPresentation As Presentation

'---------------------------------------------------
Public gblstrIndicatorStyle              As String

'For saving ribbon settings
Public gblstrFontname              As String
Public gbldblFontSize              As Double
Public gblstrFontColorSeen         As String
Public gblstrFontColorUnSeen       As String
Public gblstrFontColorCurrent       As String

Public gblstrIndicatorType         As String
Public gblintIndicatorTypeID       As Integer
Public gbldblIndicatorSize         As Double
Public gblstrIndicatorColorSeen    As String
Public gblstrIndicatorColorUnSeen  As String
Public gblstrIndicatorColorCurrent As String

Public gblintMarginLeft            As Integer
Public gblintMarginRight           As Integer
Public gblintMarginTop             As Integer

Public gblintSectionSpacing        As Integer
Public gblintIndicatorSpacingAdjustment        As Integer
Public gblintIndicatorSqueeze   As Integer

Public gblintSlideIgnoreSkipped    As Integer
Public gblintSlideIgnoreHidden     As Integer
Public gblstrSlideIgnoreSections   As String   'before 1.92 was integer. NONE = sections shown normally, TEXT = text invisible, but space left , FULL = no text, no space
Public gblstrSlideIgnoreIndicator   As String   'before 1.92 was INTEGER, but as we have a new dropdown it's string vrom 1.92 onwards

Public gblintGeneralProgressbarDirect   As Integer

Public gblintUseTextBGCol          As Integer
Public gblintUseTextBGCol_Current  As Integer
Public gblintUseTextBGCol_UnSeen   As Integer
Public gblintUseShapeBG_Seen       As Integer
Public gblintUseShapeBG_Current    As Integer
Public gblintUseShapeBG_UnSeen     As Integer


Public gblstrFontBGColorSeen       As String
Public gblstrFontBGColorUnSeen     As String
Public gblstrFontBGColorCurrent     As String

Public gblintSectionPlacement   As Integer

Public gblintSectionIndicatorGap    As Integer

Public gblShapeLineWeight As Integer
Public gblShapeLineStyle As Integer
Public gblShapeLineDashStyle As Integer


Public gblTransparencyIndicatorSeen As Integer
Public gblTransparencyIndicatorActive As Integer
Public gblTransparencyIndicatorUnseen As Integer
    
Public gblTransparencyFontSeen As Integer
Public gblTransparencyFontActive As Integer
Public gblTransparencyFontUnseen As Integer


'----------------------------------------------------
'for use in different modules
Public gblbIndicatorUpdateRunning As Boolean
Public gblCreatePressed As Boolean

Public gblvarShapePossible As Variant
Public gblvarShapeListIDs As Variant
Public gblvarShapeListNames  As Variant

Public gblstrMessageIgnore As String
Public gblstrMessageRepeat As String
Public gblstrMessageMore As String
Public gblstrMessageTagTitle As String

Public gblFontnameRibbonActivated As String

Public gblProgramVersionBase As String              ' gets value in mdlOptionsHandling
Public gblIndicatorSettingsVersion As String
Public gblstrSectionNameHidden  As String
Public gblstrPISectionHidden As String

Public gblProgramVersion As String

Public gblbolDebug As Boolean

Public gblShowPreview As Integer
Public gblTimer As Double

Public gblPIShapeCountTotalShape As Long
Public gblPIShapeCountTotalSection As Long

Public gblCancelCreate As Integer

' for REPEAT functionality
Public gblLastPI As Integer
Public gblLastPIslideNum As Integer

' for use of custom image
Public gblshowCustomImage As Integer
Public gblCustomImage_Seen As String
Public gblCustomImage_Active As String
Public gblCustomImage_UnSeen As String

' for preference: INDICATORS vs SECTIONS: What should be spaced evenly?
Public gblEvenSpaceSecOrInd As String

' set drawing start and drawing end page
Public gbliDrawStart As Integer
Public gbliDrawEnd As Integer

Function PickFile(Optional sFileName As String = "") As String
    Dim filePath As String
    'Dim isMac As Boolean
    
    ' Use the robust Mac detection method
    'isMac = IsMacPlatform()
    
    'If isMac Then
    #If Mac Then
        ' ==== macOS: Use AppleScript via MacScript ====
        On Error Resume Next
        filePath = MacScript("return POSIX path of (choose file with prompt """ & PI_Zh("8BF7900962E965874EF6FF1A") & """)")
        If Err.Number <> 0 Or filePath = "" Then
            On Error GoTo 0
            ' MacScript failed, use InputBox as fallback
            MsgBox PI_Zh("65E06CD562535F0065874EF6900962E97A9753E3FF0C8BF7624B52A88F9351655B8C657465874EF68DEF5F843002"), vbExclamation, PI_Zh("004D00610063002065874EF6900962E9")
            filePath = InputBox(PI_Zh("8BF78F9351655B8C657465874EF68DEF5F84FF0C4F8B59820020002F00550073006500720073002F0075007300650072006E0061006D0065002F0044006F00630075006D0065006E00740073002F006D" & _
                        "007900660069006C0065002E007400780074FF1A"), PI_Zh("65874EF68DEF5F84"), "/Users/" & Environ("USER") & "/Documents/" & sFileName)
        End If
        On Error GoTo 0
    #Else
        ' ==== Windows: Use FileDialog ====
        'Declare a variable as a FileDialog object.
        Dim fd As FileDialog
        
        'Create a FileDialog object as a File Picker dialog box.
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        'Declare a variable to contain the path
        'of each selected item. Even though the path is a String,
        'the variable must be a Variant because For Each...Next
        'routines only work with Variants and Objects.
        Dim vrtSelectedItem As Variant
        
        'Use a With...End With block to reference the FileDialog object.
        With fd
            'Allow only single file selection
            .AllowMultiSelect = False
            .Title = PI_Zh("900962E965874EF6")
            
            If sFileName <> "" Then
                .InitialFileName = sFileName
            End If
            
            'Use the Show method to display the File Picker dialog box and return the user's action.
            'The user pressed the button.
            If .Show = -1 Then
                'Get the first (and only) selected item
                filePath = .SelectedItems(1)
            Else
                'The user pressed Cancel.
                filePath = ""
            End If
        End With
        
        'Set the object variable to Nothing.
        Set fd = Nothing
    #End If
    
    ' Return the selected file path
    PickFile = filePath
    
End Function

Public Function PI_BindingToken(ByVal bindingId As String) As String
    If Len(bindingId) = 0 Then
        PI_BindingToken = "-"
    Else
        PI_BindingToken = bindingId
    End If
End Function

Public Function PI_RenderAxisLength(opts As ProgressDotsOptions, pres As Presentation) As Double
    If opts.DockEdge >= 2 Then
        PI_RenderAxisLength = pres.PageSetup.SlideHeight
    Else
        PI_RenderAxisLength = pres.PageSetup.SlideWidth
    End If
End Function

Public Sub PI_ClearGenerationScope()
    mPIGenerationIds = ""
End Sub

Public Function PI_IsGenerationTarget(targetSlide As Slide) As Boolean
    PI_IsGenerationTarget = (Len(mPIGenerationIds) = 0)
    If Not PI_IsGenerationTarget Then
        PI_IsGenerationTarget = (InStr(1, mPIGenerationIds, ";" & CStr(targetSlide.SlideID) & ";", vbBinaryCompare) > 0)
    End If
End Function

Private Function PI_CurrentSlide() As Slide
    On Error Resume Next
    If Application.SlideShowWindows.Count > 0 Then
        Set PI_CurrentSlide = Application.SlideShowWindows(1).View.Slide
    Else
        Set PI_CurrentSlide = ActiveWindow.Selection.SlideRange(1)
    End If
    On Error GoTo 0
End Function

Private Function PI_SelectGenerationScope(ByVal onlyCurrent As Boolean) As Boolean
    Dim target As Slide, selected As SlideRange, index As Long
    mPIGenerationIds = ";"
    If onlyCurrent Then
        Set target = PI_CurrentSlide()
        If Not target Is Nothing Then mPIGenerationIds = mPIGenerationIds & CStr(target.SlideID) & ";"
    Else
        On Error Resume Next
        Set selected = ActiveWindow.Selection.SlideRange
        On Error GoTo 0
        If Not selected Is Nothing Then
            For index = 1 To selected.Count
                mPIGenerationIds = mPIGenerationIds & CStr(selected(index).SlideID) & ";"
            Next index
        End If
    End If
    PI_SelectGenerationScope = (Len(mPIGenerationIds) > 1)
    If Not PI_SelectGenerationScope Then
        mPIGenerationIds = ""
        MsgBox PI_Zh("8BF75728666E901A89C656FE62165E7B706F72476D4F89C889C656FE4E2D90094E2D97008981590474067684987597623002"), vbInformation, PI_Zh("900962E998759762")
    End If
End Function

Public Sub PI_ScopeAction(control As IRibbonControl)
    Dim target As Slide, operation As String
    If Application.Presentations.Count = 0 Then Exit Sub
    If gblbIndicatorUpdateRunning Then Exit Sub
    On Error GoTo ScopeError
    operation = LCase$(control.Tag)
    If Not PI_SelectGenerationScope(operation = "generate_current") Then Exit Sub
    Select Case operation
        Case "remove_selected"
            DelProgressBar
        Case "hide_selected"
            For Each target In ActivePresentation.Slides
                If PI_IsGenerationTarget(target) Then
                    target.Tags.Add PI_DISPLAY_HIDDEN, "YES"
                    DelProgressBarOnSlide target.SlideIndex, "__progress-indicator__"
                End If
            Next target
        Case "generate_current", "generate_selected", "show_selected"
            If operation = "show_selected" Then
                For Each target In ActivePresentation.Slides
                    If PI_IsGenerationTarget(target) Then target.Tags.Delete PI_DISPLAY_HIDDEN
                Next target
            End If
            PersistPISettings SerializeProgressOptionsRibbon()
            gblShowPreview = 0
            gblCreatePressed = True
            RefreshProgressBar
    End Select
ScopeDone:
    PI_ClearGenerationScope
    Exit Sub
ScopeError:
    MsgBox Err.Description, vbExclamation, PI_Zh("9875976264CD4F5C")
    Resume ScopeDone
End Sub

Private Function PI_CustomerPart(ByVal owner As Object, ByVal bindingId As String) As Object
    If Len(bindingId) = 0 Then Exit Function
    On Error Resume Next
    Set PI_CustomerPart = owner.CustomerData.Item(bindingId)
    On Error GoTo 0
End Function

Private Function PI_OwnerHasBinding(ByVal owner As Object, ByVal bindingId As String) As Boolean
    Dim part As Object
    Set part = PI_CustomerPart(owner, bindingId)
    If part Is Nothing Then Exit Function
    PI_OwnerHasBinding = (StrComp(part.NamespaceURI, PI_BINDING_NS, vbBinaryCompare) = 0)
End Function

Public Function PI_ShouldDisplaySlide(targetSlide As Slide, opts As ProgressDotsOptions) As Boolean
    Dim ruleIndex As Long
    If UCase$(targetSlide.Tags(PI_DISPLAY_HIDDEN)) = "YES" Then Exit Function
    PI_EnsureRuleContext targetSlide
    If mPIRulesEnabled Then
        ruleIndex = PI_MatchingRule(targetSlide)
        PI_ShouldDisplaySlide = mPIRulesDefaultShow
        If ruleIndex > 0 Then PI_ShouldDisplaySlide = mPIRules(ruleIndex).show
        Exit Function
    End If
    Select Case opts.DisplayBindingMode
        Case 0
            PI_ShouldDisplaySlide = True
        Case 1
            PI_ShouldDisplaySlide = PI_OwnerHasBinding(targetSlide.Design.SlideMaster, opts.DisplayBindingId)
        Case 2
            PI_ShouldDisplaySlide = PI_OwnerHasBinding(targetSlide.CustomLayout, opts.DisplayBindingId)
    End Select
End Function

Public Function PI_ValidateDisplayBinding(opts As ProgressDotsOptions, pres As Presentation) As Boolean
    Dim design As Design, layout As CustomLayout
    On Error GoTo BindingError
    If Not PI_ValidateRules(pres) Then Exit Function
    If mPIRulesEnabled Then
        PI_ValidateDisplayBinding = True
        Exit Function
    End If
    If opts.DisplayBindingMode = 0 Then
        PI_ValidateDisplayBinding = True
        Exit Function
    End If
    For Each design In pres.Designs
        If opts.DisplayBindingMode = 1 Then
            If PI_OwnerHasBinding(design.SlideMaster, opts.DisplayBindingId) Then
                PI_ValidateDisplayBinding = True
                Exit Function
            End If
        Else
            For Each layout In design.SlideMaster.CustomLayouts
                If PI_OwnerHasBinding(layout, opts.DisplayBindingId) Then
                    PI_ValidateDisplayBinding = True
                    Exit Function
                End If
            Next layout
        End If
    Next design
    If PI_IsSilentGeneration() Then Err.Raise vbObjectError + 3081, "PI_ValidateDisplayBinding", "Bound master or layout does not exist."
    MsgBox PI_Zh("7ED15B9A76846BCD7248621672485F0F5DF24E0D5B585728FF0C62166B64914D7F6E676581EA53E64E004EFD65877A3F30028BF791CD65B07ED15B9A5F53524D6BCD7248002F72485F0FFF0C62166E059664663E793A7ED15B9A540E518D751F62103002"), vbExclamation, PI_Zh("663E793A7ED15B9A")
    Exit Function
BindingError:
    If PI_IsSilentGeneration() Then Err.Raise Err.Number, Err.Source, Err.Description
    MsgBox Err.Description, vbExclamation, PI_Zh("663E793A89C452199A8C8BC1")
End Function

Private Function PI_EnsureBinding(ByVal owner As Object, ByVal pres As Presentation) As String
    Dim ids As Variant, candidate As Variant, part As Object, registry As String
    registry = pres.Tags(PI_BINDING_REGISTRY)
    ids = Split(registry, ";")
    For Each candidate In ids
        If Len(CStr(candidate)) > 0 Then
            Set part = PI_CustomerPart(owner, CStr(candidate))
            If Not part Is Nothing Then
                If StrComp(part.NamespaceURI, PI_BINDING_NS, vbBinaryCompare) = 0 Then
                    PI_EnsureBinding = part.Id
                    Exit Function
                End If
            End If
        End If
    Next candidate
    Set part = owner.CustomerData.Add
    If Not part.LoadXML("<pi:binding xmlns:pi=""" & PI_BINDING_NS & """/>") Then
        Err.Raise vbObjectError + 3040, "PI_EnsureBinding", PI_Zh("65E06CD54FDD5B586BCD7248621672485F0F76847ED15B9A68078BC63002")
    End If
    PI_EnsureBinding = part.Id
    pres.Tags.Add PI_BINDING_REGISTRY, registry & ";" & part.Id
End Function

Private Function PI_CanPersistBinding(ByVal pres As Presentation) As Boolean
    Dim extension As String
    If Len(pres.Path) = 0 Then
        PI_CanPersistBinding = True
        Exit Function
    End If
    extension = LCase$(Mid$(pres.Name, InStrRev(pres.Name, ".") + 1))
    Select Case extension
        Case "pptx", "pptm", "potx", "potm", "ppsx", "ppsm"
            PI_CanPersistBinding = True
        Case Else
            MsgBox PI_Zh("8BF751485C0665877A3F53E65B584E3A0020002E0070007000740078002062160020002E007000700074006DFF0C518D8BBE7F6E6BCD7248621672485F0F7ED15B9A300265E772480020002E00700070007400204E0D4FDD75596B647ED15B9A51436570636E3002"), vbExclamation, PI_Zh("663E793A7ED15B9A")
    End Select
End Function

Public Sub PI_BindDisplay(control As IRibbonControl)
    Dim target As Slide, owner As Object, bindingId As String, mode As Integer
    If Application.Presentations.Count = 0 Then Exit Sub
    If gblbIndicatorUpdateRunning Then Exit Sub
    On Error GoTo BindingError
    PI_LoadRules ActivePresentation
    If mPIRulesEnabled Then
        MsgBox PI_Zh("8BF751485728591A89C4521983DC53554E2D505C752889C45219FF0C518D8C03657453554E00663E793A7ED15B9A3002"), vbInformation, PI_Zh("663E793A7ED15B9A")
        Exit Sub
    End If
    If LCase$(control.Tag) = "all" Then
        gblDisplayBindingMode = 0
        gblDisplayBindingId = ""
    Else
        Set target = PI_CurrentSlide()
        If target Is Nothing Then
            MsgBox PI_Zh("8BF751485728666E901A89C656FE4E2D90094E2D4E005F205E7B706F72473002"), vbInformation, PI_Zh("663E793A7ED15B9A")
            Exit Sub
        End If
        If Not PI_CanPersistBinding(ActivePresentation) Then Exit Sub
        If LCase$(control.Tag) = "master" Then
            mode = 1
            Set owner = target.Design.SlideMaster
        Else
            mode = 2
            Set owner = target.CustomLayout
        End If
        bindingId = PI_EnsureBinding(owner, ActivePresentation)
        gblDisplayBindingMode = mode
        gblDisplayBindingId = bindingId
    End If
    PI_CommitGeneralOptions
    Exit Sub
BindingError:
    MsgBox Err.Description, vbExclamation, PI_Zh("663E793A7ED15B9A")
End Sub

Public Sub PI_BindingGetLabel(control As IRibbonControl, ByRef label)
    On Error GoTo InvalidRules
    If Application.Presentations.Count > 0 Then
        PI_LoadRules ActivePresentation
        If mPIRulesEnabled Then
            label = PI_Zh("663E793AFF1A591A89C45219") & " (" & mPIRuleCount & ")"
            Exit Sub
        End If
    End If
    Select Case gblDisplayBindingMode
        Case 1: label = PI_Zh("663E793AFF1A7ED15B9A76846BCD7248")
        Case 2: label = PI_Zh("663E793AFF1A7ED15B9A768472485F0F")
        Case Else: label = PI_Zh("663E793AFF1A516890E8900275289875")
    End Select
    Exit Sub
InvalidRules:
    label = PI_Zh("663E793AFF1A89C452199700898168C067E5")
End Sub

Public Sub PI_DockGetIndex(control As IRibbonControl, ByRef index)
    index = gblDockEdge
End Sub

Public Sub PI_DockSetIndex(control As IRibbonControl, selectedID As String, selectedIndex As Integer)
    If Application.Presentations.Count = 0 Then Exit Sub
    If gblbIndicatorUpdateRunning Then Exit Sub
    gblDockEdge = CInt(PI_BarClamp(selectedIndex, 0, 3))
    PI_CommitGeneralOptions
End Sub

Public Sub PI_GeneralGetValue(control As IRibbonControl, ByRef value)
    Select Case control.Id
        Case "edtSidebarWidth": value = CStr(gblSidebarWidth)
        Case "edtSectionRoundness": value = CStr(gblSectionRoundness)
        Case "edtHardShadowOffset": value = CStr(gblHardShadowOffset)
        Case "edtMinMarkerLength": value = CStr(gblMinMarkerLength)
        Case "edtFlatRuleWidth": value = CStr(gblFlatRuleWidth)
        Case "edtFlatRuleColor": value = gblFlatRuleColor
    End Select
End Sub

Public Sub PI_GeneralSetValue(control As IRibbonControl, text As String)
    Dim value As Double, minimum As Double, maximum As Double
    If Application.Presentations.Count = 0 Then Exit Sub
    If gblbIndicatorUpdateRunning Then Exit Sub
    If control.Id = "edtFlatRuleColor" Then
        text = Replace(Trim$(text), " ", "")
        If Not PI_ValidColorToken(text) Then
            MsgBox PI_Zh("8BF78F9351654E094E2A4EE5520653F75206969476840020005200470042002065746570FF0C4F8B59820020003100300037003B003100310034003B00310032003830026BCF9879830356F44E3A00200030002081F300200032003500353002"), vbExclamation, PI_Zh("7EBF6761989C8272")
            Exit Sub
        End If
        gblFlatRuleColor = SerializeColor(DeserializeColor(text))
        PI_CommitGeneralOptions
        Exit Sub
    End If
    Select Case control.Id
        Case "edtSidebarWidth": minimum = 48: maximum = 240
        Case "edtSectionRoundness": minimum = 0: maximum = 50
        Case "edtHardShadowOffset": minimum = 0: maximum = 8
        Case "edtMinMarkerLength": minimum = 2: maximum = 20
        Case "edtFlatRuleWidth": minimum = 0.25: maximum = 3
        Case Else: Exit Sub
    End Select
    If Not IsNumeric(text) Then GoTo InvalidValue
    value = CDbl(text)
    If value < minimum Or value > maximum Then GoTo InvalidValue
    Select Case control.Id
        Case "edtSidebarWidth": gblSidebarWidth = value
        Case "edtSectionRoundness": gblSectionRoundness = value
        Case "edtHardShadowOffset": gblHardShadowOffset = value
        Case "edtMinMarkerLength": gblMinMarkerLength = value
        Case "edtFlatRuleWidth": gblFlatRuleWidth = value
    End Select
    PI_CommitGeneralOptions
    Exit Sub
InvalidValue:
    MsgBox PI_Zh("8BF78F9351650020") & CStr(minimum) & PI_Zh("002081F30020") & CStr(maximum) & PI_Zh("00204E4B95F476846570503C3002"), vbExclamation, PI_Zh("5BFC822A8BBE7F6E")
End Sub

Public Sub PI_GeneralGetPressed(control As IRibbonControl, ByRef pressed)
    Select Case control.Id
        Case "cbxUseThemeColors": pressed = (gblUseThemeColors = 1)
        Case "cbxReuseNavigation": pressed = (gblReuseNavigation = 1)
    End Select
End Sub

Public Sub PI_GeneralEnabled(control As IRibbonControl, ByRef enabled)
    enabled = (Application.Presentations.Count > 0 And Not gblbIndicatorUpdateRunning)
End Sub

Public Sub PI_LegacyBindingEnabled(control As IRibbonControl, ByRef enabled)
    On Error GoTo Disabled
    enabled = False
    If Application.Presentations.Count = 0 Or gblbIndicatorUpdateRunning Then Exit Sub
    PI_LoadRules ActivePresentation
    enabled = Not mPIRulesEnabled
Disabled:
End Sub

Public Sub PI_FoundationGetIndex(control As IRibbonControl, ByRef index)
    Select Case control.Id
        Case "ddCounterMode": index = gblCounterMode
        Case "ddDensityMode": index = gblDensityMode
        Case "ddTitleOverflow": index = gblTitleOverflow
        Case "ddTitleAlignment": index = gblTitleAlignment
        Case "ddTitleCase": index = gblTitleCase
        Case "ddMarkerKind": index = gblMarkerKind
        Case "ddFlatRuleMode": index = gblFlatRuleMode
    End Select
End Sub

Public Sub PI_FoundationSetIndex(control As IRibbonControl, selectedID As String, selectedIndex As Integer)
    If Application.Presentations.Count = 0 Or gblbIndicatorUpdateRunning Then Exit Sub
    Select Case control.Id
        Case "ddCounterMode": gblCounterMode = CInt(PI_BarClamp(selectedIndex, 0, 5))
        Case "ddDensityMode": gblDensityMode = CInt(PI_BarClamp(selectedIndex, 0, 2))
        Case "ddTitleOverflow": gblTitleOverflow = CInt(PI_BarClamp(selectedIndex, 0, 2))
        Case "ddTitleAlignment": gblTitleAlignment = CInt(PI_BarClamp(selectedIndex, 0, 2))
        Case "ddTitleCase": gblTitleCase = CInt(PI_BarClamp(selectedIndex, 0, 1))
        Case "ddMarkerKind"
            gblMarkerKind = selectedIndex
            gblFlatBar = 1
        Case "ddFlatRuleMode": gblFlatRuleMode = CInt(PI_BarClamp(selectedIndex, 0, 8))
        Case Else: Exit Sub
    End Select
    PI_CommitGeneralOptions
End Sub

Public Sub PI_TopBackdropEnabled(control As IRibbonControl, ByRef enabled)
    enabled = (gblDockEdge = 0 And Application.Presentations.Count > 0 And Not gblbIndicatorUpdateRunning)
End Sub

Public Sub PI_GeneralOnToggle(control As IRibbonControl, pressed As Boolean)
    If Application.Presentations.Count = 0 Then Exit Sub
    If gblbIndicatorUpdateRunning Then Exit Sub
    Select Case control.Id
        Case "cbxUseThemeColors": gblUseThemeColors = Abs(CInt(pressed))
        Case "cbxReuseNavigation": gblReuseNavigation = Abs(CInt(pressed))
    End Select
    PI_CommitGeneralOptions
End Sub

Private Sub PI_CommitGeneralOptions()
    PersistPISettings SerializeProgressOptionsRibbon()
    RefreshRibbon
    If Not gobjRibbon Is Nothing Then gobjRibbon.Invalidate
    RefreshProgressBar
End Sub

Public Sub PI_ClearRuleCache()
    mPIRulesLoaded = False
    mPIRulesValidated = False
    Set mPIRulesPresentation = Nothing
    Erase mPIRuleLayouts
    Erase mPILayoutMatches
    mPIRuleLayoutCount = 0
End Sub

Private Function PI_XmlEscape(ByVal value As String) As String
    If Not PI_IsXmlText(value) Then
        Err.Raise vbObjectError + 3060, "PI_XmlEscape", PI_Zh("540D79F06216914D7F6E542B67094E0D80FD4FDD5B58768463A752365B577B263002")
    End If
    value = Replace(value, "&", "&amp;")
    value = Replace(value, "<", "&lt;")
    value = Replace(value, ">", "&gt;")
    value = Replace(value, """", "&quot;")
    PI_XmlEscape = Replace(value, "'", "&apos;")
End Function

Private Function PI_IsXmlText(ByVal value As String) As Boolean
    Dim index As Long, character As Long, following As Long
    index = 1
    Do While index <= Len(value)
        character = AscW(Mid$(value, index, 1)) And &HFFFF&
        If character < 32 And character <> 9 And character <> 10 And character <> 13 Then Exit Function
        If character = &HFFFE& Or character = &HFFFF& Then Exit Function
        If character >= &HD800& And character <= &HDBFF& Then
            If index = Len(value) Then Exit Function
            following = AscW(Mid$(value, index + 1, 1)) And &HFFFF&
            If following < &HDC00& Or following > &HDFFF& Then Exit Function
            index = index + 1
        ElseIf character >= &HDC00& And character <= &HDFFF& Then
            Exit Function
        End If
        index = index + 1
    Loop
    PI_IsXmlText = True
End Function

Private Function PI_DocumentPart(ByVal pres As Presentation, ByVal namespaceURI As String, _
                                  ByVal rootName As String) As Object
    Dim parts As Object, part As Object, root As Object, mappedNamespace As String
    Set parts = pres.CustomXMLParts.SelectByNamespace(namespaceURI)
    If parts.Count = 0 Then Exit Function
    If parts.Count <> 1 Then
        Err.Raise vbObjectError + 3061, "PI_DocumentPart", PI_Zh("658768634E2D670991CD590D76848FDB5EA667616570636E90E84EF6FF0C8BF751486E0574065BF95E94768489C45219621698848BBE3002")
    End If
    Set part = parts.Item(1)
    If StrComp(part.NamespaceURI, namespaceURI, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 3062, "PI_DocumentPart", PI_Zh("8FDB5EA667616570636E547D540D7A7A95F465E065483002")
    End If
    On Error Resume Next
    mappedNamespace = part.NamespaceManager.LookupNamespace("pi")
    On Error GoTo 0
    If StrComp(mappedNamespace, namespaceURI, vbBinaryCompare) <> 0 Then part.NamespaceManager.AddNamespace "pi", namespaceURI
    Set root = part.SelectSingleNode("/pi:" & rootName)
    If root Is Nothing Then Err.Raise vbObjectError + 3063, "PI_DocumentPart", PI_Zh("8FDB5EA667616570636E6839828270B965E065483002")
    If PI_NodeText(root, "@version") <> "1" Then
        Err.Raise vbObjectError + 3064, "PI_DocumentPart", PI_Zh("6B64658768636570636E7248672C4E0D53D75F53524D63D24EF6652F63013002")
    End If
    Set PI_DocumentPart = part
End Function

Private Function PI_NodeText(ByVal node As Object, ByVal path As String) As String
    Dim found As Object
    Set found = node.SelectSingleNode(path)
    If found Is Nothing Then
        Err.Raise vbObjectError + 3065, "PI_NodeText", PI_Zh("8FDB5EA66761658768636570636E7F3A5C115FC589815B576BB5FF1A") & path
    End If
    PI_NodeText = found.Text
End Function

Private Sub PI_WriteDocumentPart(ByVal pres As Presentation, ByVal namespaceURI As String, _
                                  ByVal rootName As String, ByVal xml As String)
    Dim original As Object, replacement As Object, detail As String
    Set original = PI_DocumentPart(pres, namespaceURI, rootName)
    On Error GoTo WriteFailed
    Set replacement = pres.CustomXMLParts.Add(xml)
    If StrComp(replacement.NamespaceURI, namespaceURI, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 3066, "PI_WriteDocumentPart", PI_Zh("65E06CD59A8C8BC165B04FDD5B5876848FDB5EA667616570636E3002")
    End If
    If Not original Is Nothing Then original.Delete
    Exit Sub
WriteFailed:
    detail = Err.Description
    On Error Resume Next
    If Not replacement Is Nothing Then replacement.Delete
    On Error GoTo 0
    Err.Raise vbObjectError + 3067, "PI_WriteDocumentPart", detail
End Sub

Private Sub PI_DeleteDocumentParts(ByVal pres As Presentation, ByVal namespaceURI As String)
    Dim parts As Object, index As Long
    Set parts = pres.CustomXMLParts.SelectByNamespace(namespaceURI)
    For index = parts.Count To 1 Step -1
        If StrComp(parts.Item(index).NamespaceURI, namespaceURI, vbBinaryCompare) = 0 Then parts.Item(index).Delete
    Next index
End Sub

Private Function PI_ValidDocumentName(ByVal title As String) As Boolean
    Dim index As Long
    If Len(Trim$(title)) = 0 Or Len(title) > 80 Then Exit Function
    If Not PI_IsXmlText(title) Then Exit Function
    For index = 1 To Len(title)
        If (AscW(Mid$(title, index, 1)) And &HFFFF&) < 32 Then Exit Function
    Next index
    PI_ValidDocumentName = True
End Function

Private Function PI_AskDocumentName(ByVal prompt As String, ByVal initialName As String) As String
    Dim title As String
    title = Trim$(InputBox(prompt, PI_Zh("8FDB5EA66761658768638BBE7F6E"), initialName))
    If Len(title) = 0 Then Exit Function
    If Not PI_ValidDocumentName(title) Then
        MsgBox PI_Zh("540D79F0987B4E3A00200031002081F300200038003000204E2A5B577B26FF0C4E0D80FD542B6362884C621663A752365B577B263002"), vbExclamation, PI_Zh("540D79F065E06548")
        Exit Function
    End If
    PI_AskDocumentName = title
End Function

Private Function PI_ValidColorToken(ByVal value As String) As Boolean
    Dim channels As Variant, index As Long, number As Double
    channels = Split(value, ";")
    If UBound(channels) <> 2 Then Exit Function
    For index = 0 To 2
        If Not IsNumeric(channels(index)) Then Exit Function
        number = CDbl(channels(index))
        If number < 0 Or number > 255 Or number <> Fix(number) Then Exit Function
    Next index
    PI_ValidColorToken = True
End Function

Private Function PI_ValidStoredConfig(ByVal config As String) As Boolean
    Dim fields As Variant, index As Long, number As Double
    On Error GoTo InvalidConfig
    If Len(config) > 16384 Then Exit Function
    If InStr(config, vbCr) > 0 Or InStr(config, vbLf) > 0 Or InStr(config, vbTab) > 0 Then Exit Function
    fields = Split(config, " ")
    Select Case UBound(fields)
        Case 99, 101, 103, 104, 105, 106
        Case Else: Exit Function
    End Select
    If fields(0) <> gblIndicatorSettingsVersion Then Exit Function
    For index = 1 To UBound(fields)
        If Len(fields(index)) = 0 Then Exit Function
        Select Case index
            Case 7, 8, 9, 10, 25, 26, 30, 31, 32, 52, 53, 54, 55, 57, 78, 99
                If Not PI_ValidColorToken(CStr(fields(index))) Then Exit Function
            Case 11, 12, 86
                If Not PI_IsXmlText(CStr(fields(index))) Then Exit Function
            Case 16
                If fields(index) <> "Full" And fields(index) <> "Space#Only" And fields(index) <> "None" Then Exit Function
            Case 36
                If fields(index) <> "All" And fields(index) <> "Active#Section" And fields(index) <> "None" Then Exit Function
            Case 43
                If fields(index) <> "Shapes" And fields(index) <> "Sections" Then Exit Function
            Case 44
                If fields(index) <> "Full" And fields(index) <> "Indicator#only" And fields(index) <> "Sections#only" Then Exit Function
            Case Else
                If Not IsNumeric(fields(index)) Then Exit Function
                If index < 45 Then
                    number = CDbl(fields(index))
                Else
                    number = Val(fields(index))
                End If
                If number < -32768 Or number > 32767 Then Exit Function
                Select Case index
                    Case 1 To 3, 6, 13 To 15, 17 To 24, 27 To 29, 33 To 35, 37 To 42, 105
                        If number <> Fix(number) Then Exit Function
                End Select
                Select Case index
                    Case 1 To 3, 6
                        If number < 0 Then Exit Function
                    Case 4
                        If number < 1 Or number > 4000 Then Exit Function
                    Case 5
                        If number <= 0 Or number > 4032 Then Exit Function
                    Case 13
                        If Not IsArray(gblvarShapeListIDs) Then Exit Function
                        If number < 0 Or number >= UBound(gblvarShapeListIDs) Then Exit Function
                    Case 14, 15, 17, 19 To 24, 28, 45, 47, 48, 56, 60, 64, 67, 79 To 82, 87, 89, 96, 101, 103
                        If number <> 0 And number <> 1 Then Exit Function
                    Case 106
                        If number <> PI_STYLE_CONFIG_REVISION Then Exit Function
                    Case 102, 104
                        If number < 0 Or number > 2 Or number <> Fix(number) Then Exit Function
                    Case 33
                        If number < 0 Or number > 1584 Then Exit Function
                    Case 34
                        If number < 1 Or number > 5 Then Exit Function
                    Case 35
                        If number < 1 Or number > 8 Then Exit Function
                    Case 37 To 42, 49 To 51, 58, 66, 73, 105
                        If number < 0 Or number > 100 Then Exit Function
                    Case 46, 62, 65, 72, 75
                        If number < 0 Or number > 20 Then Exit Function
                    Case 59
                        If number < 110 Or number > 300 Then Exit Function
                    Case 61
                        If number < 0.5 Or number > 12 Then Exit Function
                    Case 63, 88
                        If number < 0 Or number > 50 Then Exit Function
                    Case 68, 70
                        If number < 0 Or number > 4032 Then Exit Function
                    Case 69
                        If number < 0 Or number > 200 Then Exit Function
                    Case 71
                        If number < 0.5 Or number > 30 Then Exit Function
                    Case 74
                        If number < 0 Or number > 5 Then Exit Function
                    Case 76
                        If number < 0 Or number > 4 Or number <> Fix(number) Then Exit Function
                    Case 77
                        If number < 0.5 Or number > 4 Then Exit Function
                    Case 83
                        If number < 0 Or number > 3 Or number <> Fix(number) Then Exit Function
                    Case 84
                        If number < 48 Or number > 240 Then Exit Function
                    Case 85, 92, 94, 95
                        If number < 0 Or number > 2 Or number <> Fix(number) Then Exit Function
                    Case 91, 100
                        If number < 0 Or number > 5 Or number <> Fix(number) Then Exit Function
                    Case 90
                        If number < 0 Or number > 8 Then Exit Function
                    Case 93
                        If number < 2 Or number > 20 Then Exit Function
                    Case 97
                        If number < 0 Or number > 8 Or number <> Fix(number) Then Exit Function
                    Case 98
                        If number < 0.25 Or number > 3 Then Exit Function
                End Select
        End Select
    Next index
    PI_ValidStoredConfig = True
InvalidConfig:
End Function

Private Sub PI_LoadPresets(ByVal pres As Presentation)
    Dim part As Object, nodes As Object, node As Object, index As Long, prior As Long
    mPIPresetCount = 0
    Set part = PI_DocumentPart(pres, PI_PRESETS_NS, "presets")
    If part Is Nothing Then Exit Sub
    Set nodes = part.SelectNodes("/pi:presets/pi:preset")
    If nodes.Count > PI_DOCUMENT_LIMIT Then
        Err.Raise vbObjectError + 3070, "PI_LoadPresets", PI_Zh("6B64658768636700591A652F630100200033003200204E2A7528623798848BBE3002")
    End If
    For index = 1 To nodes.Count
        Set node = nodes.Item(index)
        mPIPresets(index).title = PI_NodeText(node, "pi:name")
        mPIPresets(index).config = PI_NodeText(node, "pi:config")
        If Not PI_ValidDocumentName(mPIPresets(index).title) Then
            Err.Raise vbObjectError + 3071, "PI_LoadPresets", PI_Zh("658768634E2D670965E06548768498848BBE540D79F03002")
        End If
        For prior = 1 To index - 1
            If StrComp(mPIPresets(prior).title, mPIPresets(index).title, vbBinaryCompare) = 0 Then
                Err.Raise vbObjectError + 3072, "PI_LoadPresets", PI_Zh("658768634E2D670991CD590D768498848BBE540D79F03002")
            End If
        Next prior
    Next index
    mPIPresetCount = nodes.Count
End Sub

Private Sub PI_SavePresets(ByVal pres As Presentation)
    Dim xml As String, index As Long
    xml = "<pi:presets xmlns:pi=""" & PI_PRESETS_NS & """ version=""1"">"
    For index = 1 To mPIPresetCount
        xml = xml & "<pi:preset><pi:name>" & PI_XmlEscape(mPIPresets(index).title) & "</pi:name>"
        xml = xml & "<pi:config>" & PI_XmlEscape(mPIPresets(index).config) & "</pi:config></pi:preset>"
    Next index
    PI_WriteDocumentPart pres, PI_PRESETS_NS, "presets", xml & "</pi:presets>"
    If Not gobjRibbon Is Nothing Then gobjRibbon.InvalidateControl "mnuDocumentPresets"
End Sub

Private Function PI_MenuButton(ByVal identifier As String, ByVal title As String, _
                                 ByVal action As String, ByVal tag As String) As String
    PI_MenuButton = "<button id=""" & identifier & """ label=""" & PI_XmlEscape(title) & _
                   """ onAction=""" & action & """ tag=""" & PI_XmlEscape(tag) & """/>"
End Function

Private Function PI_MenuStart() As String
    PI_MenuStart = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">"
End Function

Public Sub PI_PresetsGetContent(control As IRibbonControl, ByRef content)
    Dim xml As String, index As Long
    On Error GoTo MenuError
    xml = PI_MenuStart()
    If Application.Presentations.Count = 0 Then
        content = xml & "<button id=""piPresetNoDocument"" label=""" & PI_Zh("8BF7514862535F006F14793A65877A3F") & """ enabled=""false""/></menu>"
        Exit Sub
    End If
    Set mPIPresetMenuPresentation = ActivePresentation
    PI_LoadPresets ActivePresentation
    xml = xml & PI_MenuButton("piPresetSave", PI_Zh("4FDD5B585F53524D914D7F6E4E3A6587686398848BBE"), "PI_PresetAction", "save")
    xml = xml & "<menuSeparator id=""piPresetSeparator"" title=""" & PI_Zh("6B6465876863768498848BBE") & """/>"
    For index = 1 To mPIPresetCount
        xml = xml & "<menu id=""piPreset" & index & """ label=""" & PI_XmlEscape(mPIPresets(index).title) & """>"
        xml = xml & PI_MenuButton("piPresetApply" & index, PI_Zh("5E947528"), "PI_PresetAction", "apply:" & index)
        xml = xml & PI_MenuButton("piPresetRename" & index, PI_Zh("91CD547D540D"), "PI_PresetAction", "rename:" & index)
        xml = xml & PI_MenuButton("piPresetDelete" & index, PI_Zh("52209664"), "PI_PresetAction", "delete:" & index)
        xml = xml & "</menu>"
    Next index
    If mPIPresetCount = 0 Then
        xml = xml & "<button id=""piPresetEmpty"" label=""" & PI_Zh("5C1A65E098848BBEFF1B4FDD5B5857285F53524D658768634E2D") & """ enabled=""false""/>"
    End If
    content = xml & "</menu>"
    Exit Sub
MenuError:
    content = PI_MenuStart() & "<button id=""piPresetReadError"" label=""" & _
              PI_Zh("98848BBE6570636E9700898168C067E5") & """ enabled=""false""/>" & _
              PI_MenuButton("piPresetClearInvalid", PI_Zh("6E0596646B6465876863768498848BBE6570636E"), "PI_PresetAction", "clear") & "</menu>"
End Sub

Public Sub PI_PresetAction(control As IRibbonControl)
    Dim action As String, index As Long, other As Long, title As String, config As String
    Dim options As ProgressDotsOptions, globals As ProgressDotsOptions
    If Application.Presentations.Count = 0 Or gblbIndicatorUpdateRunning Then Exit Sub
    On Error GoTo PresetError
    If Not mPIPresetMenuPresentation Is ActivePresentation Then
        Err.Raise vbObjectError + 3077, "PI_PresetAction", PI_Zh("5F53524D658768635DF2653953D8FF0C8BF791CD65B062535F006587686398848BBE83DC53553002")
    End If
    If Not PI_CanPersistBinding(ActivePresentation) Then Exit Sub
    action = LCase$(Split(control.Tag, ":")(0))
    If action = "clear" Then
        If MsgBox(PI_Zh("6E0596646B64658768637684516890E87528623798848BBE6570636EFF1F663E793A89C452194E0D53D75F7154CD3002"), vbQuestion + vbYesNo, PI_Zh("6E0596646587686398848BBE")) <> vbYes Then Exit Sub
        PI_DeleteDocumentParts ActivePresentation, PI_PRESETS_NS
        If Not gobjRibbon Is Nothing Then gobjRibbon.InvalidateControl "mnuDocumentPresets"
        Exit Sub
    End If
    PI_LoadPresets ActivePresentation
    If action = "save" Then
        title = PI_AskDocumentName(PI_Zh("4E3A5F53524D914D7F6E547D540D300298848BBE4EC54FDD5B5857285F53524D6F14793A65877A3F4E2D3002"), PI_Zh("6211768498848BBE"))
        If Len(title) = 0 Then Exit Sub
        For other = 1 To mPIPresetCount
            If StrComp(mPIPresets(other).title, title, vbBinaryCompare) = 0 Then
                index = other
                Exit For
            End If
        Next other
        If index > 0 Then
            If MsgBox(PI_Zh("66FF6362540C540D98848BBEFF1A") & title & "?", vbQuestion + vbYesNo, PI_Zh("4FDD5B5898848BBE")) <> vbYes Then Exit Sub
        Else
            If mPIPresetCount >= PI_DOCUMENT_LIMIT Then
                Err.Raise vbObjectError + 3073, "PI_PresetAction", PI_Zh("6700591A4FDD5B5800200033003200204E2A98848BBEFF0C8BF75148522096644E0D97008981768498848BBE3002")
            End If
            index = mPIPresetCount + 1
        End If
        config = SerializeProgressOptionsRibbon()
        If Not PI_ValidStoredConfig(config) Then
            Err.Raise vbObjectError + 3074, "PI_PresetAction", PI_Zh("5F53524D914D7F6E4E0D5B8C6574FF0C65E06CD54FDD5B584E3A98848BBE3002")
        End If
        mPIPresets(index).title = title
        mPIPresets(index).config = config
        If index > mPIPresetCount Then mPIPresetCount = index
        PI_SavePresets ActivePresentation
        Exit Sub
    End If
    index = CLng(Val(Mid$(control.Tag, InStr(control.Tag, ":") + 1)))
    If index < 1 Or index > mPIPresetCount Then Exit Sub
    Select Case action
        Case "apply"
            If Not PI_ValidStoredConfig(mPIPresets(index).config) Then
                Err.Raise vbObjectError + 3075, "PI_PresetAction", PI_Zh("8BE598848BBE914D7F6E65E06548FF0C8BF791CD65B04FDD5B58621652209664FF1A") & mPIPresets(index).title
            End If
            options = DeserializeProgressOptions(mPIPresets(index).config)
            globals = getDefaultOptions()
            options = PI_RuleVisualOptions(options, globals)
            options.DisplayBindingMode = globals.DisplayBindingMode
            options.DisplayBindingId = globals.DisplayBindingId
            optionsToRibbon options
            PI_CommitGeneralOptions
        Case "rename"
            title = PI_AskDocumentName(PI_Zh("8F93516565B0768498848BBE540D79F03002"), mPIPresets(index).title)
            If Len(title) = 0 Then Exit Sub
            For other = 1 To mPIPresetCount
                If other <> index Then
                    If StrComp(mPIPresets(other).title, title, vbBinaryCompare) = 0 Then
                        Err.Raise vbObjectError + 3076, "PI_PresetAction", PI_Zh("6B64540D79F05DF288AB51764ED698848BBE4F7F75283002")
                    End If
                End If
            Next other
            mPIPresets(index).title = title
            PI_SavePresets ActivePresentation
        Case "delete"
            If MsgBox(PI_Zh("522096646587686398848BBEFF1A") & mPIPresets(index).title & "?", vbQuestion + vbYesNo, PI_Zh("5220966498848BBE")) <> vbYes Then Exit Sub
            For other = index To mPIPresetCount - 1
                mPIPresets(other) = mPIPresets(other + 1)
            Next other
            mPIPresetCount = mPIPresetCount - 1
            PI_SavePresets ActivePresentation
    End Select
    Exit Sub
PresetError:
    MsgBox Err.Description, vbExclamation, PI_Zh("6587686398848BBE")
End Sub

Private Sub PI_LoadRules(ByVal pres As Presentation, Optional ByVal force As Boolean = False)
    Dim part As Object, root As Object, nodes As Object, node As Object
    Dim index As Long, prior As Long, value As String
    If Not force And mPIRulesLoaded Then
        If mPIRulesPresentation Is pres Then Exit Sub
    End If
    PI_ClearRuleCache
    mPIRuleCount = 0
    mPIRulesEnabled = False
    mPIRulesDefaultShow = True
    Set part = PI_DocumentPart(pres, PI_RULES_NS, "rules")
    If Not part Is Nothing Then
        Set root = part.SelectSingleNode("/pi:rules")
        value = PI_NodeText(root, "@enabled")
        If value <> "0" And value <> "1" Then GoTo InvalidRules
        mPIRulesEnabled = (value = "1")
        value = PI_NodeText(root, "@default")
        If value <> "show" And value <> "hide" Then GoTo InvalidRules
        mPIRulesDefaultShow = (value = "show")
        Set nodes = part.SelectNodes("/pi:rules/pi:rule")
        If nodes.Count > PI_DOCUMENT_LIMIT Then GoTo InvalidRules
        For index = 1 To nodes.Count
            Set node = nodes.Item(index)
            mPIRules(index).title = PI_NodeText(node, "pi:name")
            If Not PI_ValidDocumentName(mPIRules(index).title) Then GoTo InvalidRules
            value = PI_NodeText(node, "@kind")
            Select Case value
                Case "master": mPIRules(index).bindingMode = 1
                Case "layout": mPIRules(index).bindingMode = 2
                Case Else: GoTo InvalidRules
            End Select
            mPIRules(index).bindingId = PI_NodeText(node, "@target")
            If Len(mPIRules(index).bindingId) = 0 Or Len(mPIRules(index).bindingId) > 80 Then GoTo InvalidRules
            value = PI_NodeText(node, "@action")
            If value <> "show" And value <> "hide" Then GoTo InvalidRules
            mPIRules(index).show = (value = "show")
            mPIRules(index).config = PI_NodeText(node, "pi:config")
            For prior = 1 To index - 1
                If mPIRules(prior).bindingMode = mPIRules(index).bindingMode Then
                    If StrComp(mPIRules(prior).bindingId, mPIRules(index).bindingId, vbTextCompare) = 0 Then GoTo InvalidRules
                End If
            Next prior
        Next index
        mPIRuleCount = nodes.Count
    End If
    Set mPIRulesPresentation = pres
    mPIRulesLoaded = True
    Exit Sub
InvalidRules:
    Err.Raise vbObjectError + 3080, "PI_LoadRules", PI_Zh("591A89C452196570636E65E0654830028BF768C067E589C45219FF0C6216572889C4521983DC53554E2D6E057A7A540E91CD65B08BBE7F6E3002")
End Sub

Private Sub PI_SaveRules(ByVal pres As Presentation)
    Dim xml As String, index As Long, kind As String, action As String, fallback As String
    fallback = "hide"
    If mPIRulesDefaultShow Then fallback = "show"
    xml = "<pi:rules xmlns:pi=""" & PI_RULES_NS & """ version=""1"" enabled=""" & _
          Abs(CInt(mPIRulesEnabled)) & """ default=""" & fallback & """>"
    For index = 1 To mPIRuleCount
        kind = "master"
        If mPIRules(index).bindingMode = 2 Then kind = "layout"
        action = "hide"
        If mPIRules(index).show Then action = "show"
        xml = xml & "<pi:rule kind=""" & kind & """ target=""" & PI_XmlEscape(mPIRules(index).bindingId) & _
              """ action=""" & action & """><pi:name>" & PI_XmlEscape(mPIRules(index).title) & "</pi:name>"
        xml = xml & "<pi:config>" & PI_XmlEscape(mPIRules(index).config) & "</pi:config></pi:rule>"
    Next index
    PI_WriteDocumentPart pres, PI_RULES_NS, "rules", xml & "</pi:rules>"
    PI_ClearRuleCache
    If Not gobjRibbon Is Nothing Then
        gobjRibbon.InvalidateControl "mnuDocumentRules"
        gobjRibbon.InvalidateControl "lblDisplayBinding"
    End If
End Sub

Private Function PI_BindingExists(ByVal pres As Presentation, ByVal mode As Integer, ByVal bindingId As String) As Boolean
    Dim design As Design, layout As CustomLayout
    For Each design In pres.Designs
        If mode = 1 Then
            If PI_OwnerHasBinding(design.SlideMaster, bindingId) Then
                PI_BindingExists = True
                Exit Function
            End If
        ElseIf mode = 2 Then
            For Each layout In design.SlideMaster.CustomLayouts
                If PI_OwnerHasBinding(layout, bindingId) Then
                    PI_BindingExists = True
                    Exit Function
                End If
            Next layout
        End If
    Next design
End Function

Private Function PI_ValidateRules(ByVal pres As Presentation) As Boolean
    Dim index As Long
    PI_LoadRules pres, True
    If mPIRulesEnabled Then
        For index = 1 To mPIRuleCount
            If Not PI_BindingExists(pres, mPIRules(index).bindingMode, mPIRules(index).bindingId) Then
                Err.Raise vbObjectError + 3081, "PI_ValidateRules", _
                    PI_Zh("89C4521962407ED15B9A76846BCD7248621672485F0F5DF24E0D5B585728FF1A") & mPIRules(index).title & vbCrLf & _
                    PI_Zh("8BF752209664621691CD65B04FDD5B588BE589C45219FF0C7136540E518D751F6210300273B067098FDB5EA66761672A4F5C66F465393002")
            End If
            If Not PI_ValidStoredConfig(mPIRules(index).config) Then
                Err.Raise vbObjectError + 3082, "PI_ValidateRules", _
                    PI_Zh("89C45219914D7F6E65E06548FF1A") & mPIRules(index).title & vbCrLf & _
                    PI_Zh("8BF791CD65B04FDD5B588BE589C4521962165C06517652209664300273B067098FDB5EA66761672A4F5C66F465393002")
            End If
            mPIRuleOptions(index) = DeserializeProgressOptions(mPIRules(index).config)
        Next index
    End If
    mPIRulesValidated = True
    PI_ValidateRules = True
End Function

Private Function PI_MatchingRule(ByVal targetSlide As Slide) As Long
    Dim index As Long, pass As Long, result As Long, targetLayout As CustomLayout
    If Not mPIRulesEnabled Then Exit Function
    Set targetLayout = targetSlide.CustomLayout
    For index = 1 To mPIRuleLayoutCount
        If mPIRuleLayouts(index) Is targetLayout Then
            PI_MatchingRule = mPILayoutMatches(index)
            Exit Function
        End If
    Next index
    For pass = 2 To 1 Step -1
        For index = 1 To mPIRuleCount
            If mPIRules(index).bindingMode = pass Then
                If pass = 2 Then
                    If PI_OwnerHasBinding(targetLayout, mPIRules(index).bindingId) Then
                        result = index
                        GoTo CacheResult
                    End If
                Else
                    If PI_OwnerHasBinding(targetSlide.Design.SlideMaster, mPIRules(index).bindingId) Then
                        result = index
                        GoTo CacheResult
                    End If
                End If
            End If
        Next index
    Next pass
CacheResult:
    mPIRuleLayoutCount = mPIRuleLayoutCount + 1
    ReDim Preserve mPIRuleLayouts(1 To mPIRuleLayoutCount)
    ReDim Preserve mPILayoutMatches(1 To mPIRuleLayoutCount)
    Set mPIRuleLayouts(mPIRuleLayoutCount) = targetLayout
    mPILayoutMatches(mPIRuleLayoutCount) = result
    PI_MatchingRule = result
End Function

Private Sub PI_EnsureRuleContext(ByVal targetSlide As Slide)
    Dim pres As Presentation
    Set pres = targetSlide.Parent
    PI_LoadRules pres
    If Not mPIRulesValidated Then
        If Not PI_ValidateRules(pres) Then Err.Raise vbObjectError + 3083, "PI_EnsureRuleContext", PI_Zh("89C45219672A901A8FC79A8C8BC13002")
    End If
End Sub

Private Function PI_RuleVisualOptions(saved As ProgressDotsOptions, globalOptions As ProgressDotsOptions) As ProgressDotsOptions
    Dim merged As ProgressDotsOptions
    merged = saved
    'One run has one denominator and one execution policy, regardless of matched style.
    merged.ignoreHidden = globalOptions.ignoreHidden
    merged.ignoreSkipped = globalOptions.ignoreSkipped
    merged.ignoreSections = globalOptions.ignoreSections
    merged.ignoreIndicator = globalOptions.ignoreIndicator
    merged.DirectUpdate = globalOptions.DirectUpdate
    merged.ReuseNavigation = globalOptions.ReuseNavigation
    merged.DisplayBindingMode = 0
    merged.DisplayBindingId = ""
    PI_RuleVisualOptions = merged
End Function

Private Function PI_RuleDescription(ByVal index As Long) As String
    Dim kind As String, action As String
    kind = PI_Zh("6BCD7248")
    If mPIRules(index).bindingMode = 2 Then kind = PI_Zh("72485F0F")
    action = PI_Zh("969085CF")
    If mPIRules(index).show Then action = PI_Zh("663E793A")
    PI_RuleDescription = kind & " / " & action & " / " & mPIRules(index).title
End Function

Public Sub PI_RulesGetContent(control As IRibbonControl, ByRef content)
    Dim xml As String, index As Long, status As String
    On Error GoTo MenuError
    xml = PI_MenuStart()
    If Application.Presentations.Count = 0 Then
        content = xml & "<button id=""piRuleNoDocument"" label=""" & PI_Zh("8BF7514862535F006F14793A65877A3F") & """ enabled=""false""/></menu>"
        Exit Sub
    End If
    Set mPIRuleMenuPresentation = ActivePresentation
    PI_LoadRules ActivePresentation, True
    status = PI_Zh("542F7528591A89C45219")
    If mPIRulesEnabled Then status = PI_Zh("505C7528591A89C45219FF0C6062590D53554E007ED15B9A")
    xml = xml & PI_MenuButton("piRulesToggle", status, "PI_RuleAction", "toggle")
    xml = xml & "<menu id=""piRulesAdd"" label=""" & PI_Zh("4E3A5F53524D987597624FDD5B5889C45219") & """>"
    xml = xml & PI_MenuButton("piRuleMasterShow", PI_Zh("6BCD7248FF1A663E793A5E764F7F75285F53524D68375F0F"), "PI_RuleAction", "master_show")
    xml = xml & PI_MenuButton("piRuleLayoutShow", PI_Zh("72485F0FFF1A663E793A5E764F7F75285F53524D68375F0F"), "PI_RuleAction", "layout_show")
    xml = xml & PI_MenuButton("piRuleMasterHide", PI_Zh("6BCD7248FF1A969085CF8FDB5EA66761"), "PI_RuleAction", "master_hide")
    xml = xml & PI_MenuButton("piRuleLayoutHide", PI_Zh("72485F0FFF1A969085CF8FDB5EA66761"), "PI_RuleAction", "layout_hide") & "</menu>"
    status = PI_Zh("672A5339914D98759762FF1A663E793A")
    If Not mPIRulesDefaultShow Then status = PI_Zh("672A5339914D98759762FF1A969085CF")
    xml = xml & "<menu id=""piRulesDefault"" label=""" & status & """>"
    xml = xml & PI_MenuButton("piRulesDefaultShow", PI_Zh("663E793AFF0C4F7F752851685C40914D7F6E"), "PI_RuleAction", "default_show")
    xml = xml & PI_MenuButton("piRulesDefaultHide", PI_Zh("969085CF"), "PI_RuleAction", "default_hide") & "</menu>"
    xml = xml & "<menuSeparator id=""piRulesSeparator"" title=""" & PI_Zh("72485F0F4F1851484E8E6BCD7248FF1B624B52A8969085CF4F185148") & """/>"
    For index = 1 To mPIRuleCount
        xml = xml & "<menu id=""piRule" & index & """ label=""" & PI_XmlEscape(PI_RuleDescription(index)) & """>"
        xml = xml & PI_MenuButton("piRuleView" & index, PI_Zh("67E5770B89C45219"), "PI_RuleAction", "view:" & index)
        xml = xml & PI_MenuButton("piRuleLoad" & index, PI_Zh("8F7D516568375F0F5230529F80FD533A"), "PI_RuleAction", "load:" & index)
        xml = xml & PI_MenuButton("piRuleRemove" & index, PI_Zh("5220966489C45219"), "PI_RuleAction", "remove:" & index) & "</menu>"
    Next index
    xml = xml & PI_MenuButton("piRulesClear", PI_Zh("6E057A7A89C452195E76505C7528"), "PI_RuleAction", "clear")
    content = xml & "</menu>"
    Exit Sub
MenuError:
    content = PI_MenuStart() & "<button id=""piRulesReadError"" label=""" & _
              PI_XmlEscape(PI_Zh("89C452196570636E65E06548FF0C8BF76E057A7A540E91CD5EFA")) & """ enabled=""false""/>" & _
              PI_MenuButton("piRulesClearInvalid", PI_Zh("6E057A7A89C452195E76505C7528"), "PI_RuleAction", "clear") & "</menu>"
End Sub

Public Sub PI_RuleAction(control As IRibbonControl)
    Dim action As String, index As Long, other As Long, title As String, bindingId As String, detail As String
    Dim mode As Integer, target As Slide, owner As Object, options As ProgressDotsOptions, globals As ProgressDotsOptions
    If Application.Presentations.Count = 0 Or gblbIndicatorUpdateRunning Then Exit Sub
    On Error GoTo RuleError
    If Not mPIRuleMenuPresentation Is ActivePresentation Then
        Err.Raise vbObjectError + 3088, "PI_RuleAction", PI_Zh("5F53524D658768635DF2653953D8FF0C8BF791CD65B062535F00663E793A89C4521983DC53553002")
    End If
    action = LCase$(Split(control.Tag, ":")(0))
    If action = "clear" Then
        If MsgBox(PI_Zh("522096646B6465876863516890E8663E793A89C452195E766062590D53554E007ED15B9AFF1F7528623798848BBE4E0D53D75F7154CD3002"), _
                  vbQuestion + vbYesNo, PI_Zh("6E057A7A663E793A89C45219")) <> vbYes Then Exit Sub
        PI_DeleteDocumentParts ActivePresentation, PI_RULES_NS
        PI_ClearRuleCache
        PI_CommitGeneralOptions
        Exit Sub
    End If
    PI_LoadRules ActivePresentation, True
    If action = "view" Or action = "load" Or action = "remove" Then
        index = CLng(Val(Mid$(control.Tag, InStr(control.Tag, ":") + 1)))
        If index < 1 Or index > mPIRuleCount Then Exit Sub
        Select Case action
            Case "view"
                detail = PI_RuleDescription(index) & vbCrLf & PI_Zh("76EE680768078BC6FF1A") & mPIRules(index).bindingId
                If Not PI_BindingExists(ActivePresentation, mPIRules(index).bindingMode, mPIRules(index).bindingId) Then
                    detail = detail & vbCrLf & PI_Zh("72B66001FF1A76EE68074E0D5B585728FF0C8BF791CD65B04FDD5B586216522096646B6489C452193002")
                ElseIf Not PI_ValidStoredConfig(mPIRules(index).config) Then
                    detail = detail & vbCrLf & PI_Zh("72B66001FF1A4FDD5B587684914D7F6E65E065483002")
                Else
                    options = DeserializeProgressOptions(mPIRules(index).config)
                    detail = detail & vbCrLf & PI_Zh("505C97608FB9FF1A") & Choose(options.DockEdge + 1, PI_Zh("987690E8"), PI_Zh("5E9590E8"), PI_Zh("5DE64FA7"), PI_Zh("53F34FA7"))
                    detail = detail & vbCrLf & PI_Zh("5B574F53FF1A") & options.fontName & " / " & options.textHeight & " pt"
                    detail = detail & vbCrLf & PI_Zh("52066BCD59CB7EC84F7F752851685C407684969085CF30018DF38FC74E0E7AE082827EDF8BA18BBE7F6E3002")
                End If
                MsgBox detail, vbInformation, PI_Zh("663E793A89C45219")
                Exit Sub
            Case "load"
                If Not PI_ValidStoredConfig(mPIRules(index).config) Then
                    Err.Raise vbObjectError + 3084, "PI_RuleAction", PI_Zh("4FDD5B58768489C45219914D7F6E65E065483002")
                End If
                options = DeserializeProgressOptions(mPIRules(index).config)
                globals = getDefaultOptions()
                options = PI_RuleVisualOptions(options, globals)
                options.DisplayBindingMode = globals.DisplayBindingMode
                options.DisplayBindingId = globals.DisplayBindingId
                optionsToRibbon options
                PI_CommitGeneralOptions
                Exit Sub
            Case "remove"
                If MsgBox(PI_Zh("5220966489C45219FF1A") & mPIRules(index).title & "?", vbQuestion + vbYesNo, PI_Zh("5220966489C45219")) <> vbYes Then Exit Sub
                For other = index To mPIRuleCount - 1
                    mPIRules(other) = mPIRules(other + 1)
                Next other
                mPIRuleCount = mPIRuleCount - 1
        End Select
    Else
        If Not PI_CanPersistBinding(ActivePresentation) Then Exit Sub
        Select Case action
            Case "toggle"
                mPIRulesEnabled = Not mPIRulesEnabled
            Case "default_show"
                mPIRulesDefaultShow = True
            Case "default_hide"
                mPIRulesDefaultShow = False
            Case "master_show", "layout_show", "master_hide", "layout_hide"
                Set target = PI_CurrentSlide()
                If target Is Nothing Then
                    Err.Raise vbObjectError + 3085, "PI_RuleAction", PI_Zh("8BF751485728666E901A89C656FE4E2D90094E2D4E005F205E7B706F72473002")
                End If
                mode = 1
                Set owner = target.Design.SlideMaster
                If Left$(action, 6) = "layout" Then
                    mode = 2
                    Set owner = target.CustomLayout
                End If
                title = PI_AskDocumentName(PI_Zh("4E3A6B64663E793A89C45219547D540D30025DF26709540C4E0076EE680789C4521965F65C0666FF63623002"), owner.Name)
                If Len(title) = 0 Then Exit Sub
                For other = 1 To mPIRuleCount
                    If mPIRules(other).bindingMode = mode Then
                        If PI_OwnerHasBinding(owner, mPIRules(other).bindingId) Then
                            index = other
                            Exit For
                        End If
                    End If
                Next other
                If index = 0 Then
                    If mPIRuleCount >= PI_DOCUMENT_LIMIT Then
                        Err.Raise vbObjectError + 3086, "PI_RuleAction", PI_Zh("6700591A652F63010020003300320020676189C45219FF0C8BF75148522096644E0D97008981768489C452193002")
                    End If
                    index = mPIRuleCount + 1
                Else
                    If MsgBox(PI_Zh("66FF63626B6476EE68075DF2670989C45219FF1A") & mPIRules(index).title & "?", _
                              vbQuestion + vbYesNo, PI_Zh("4FDD5B58663E793A89C45219")) <> vbYes Then Exit Sub
                End If
                detail = SerializeProgressOptionsRibbon()
                If Not PI_ValidStoredConfig(detail) Then
                    Err.Raise vbObjectError + 3087, "PI_RuleAction", PI_Zh("5F53524D914D7F6E4E0D5B8C6574FF0C65E06CD54FDD5B584E3A89C452193002")
                End If
                bindingId = PI_EnsureBinding(owner, ActivePresentation)
                mPIRules(index).title = title
                mPIRules(index).bindingMode = mode
                mPIRules(index).bindingId = bindingId
                mPIRules(index).show = (Right$(action, 4) = "show")
                mPIRules(index).config = detail
                If index > mPIRuleCount Then mPIRuleCount = index
                mPIRulesEnabled = True
            Case Else
                Exit Sub
        End Select
    End If
    PI_SaveRules ActivePresentation
    PI_CommitGeneralOptions
    Exit Sub
RuleError:
    PI_ClearRuleCache
    MsgBox Err.Description, vbExclamation, PI_Zh("663E793A89C45219")
End Sub

Private Function PI_ColorFromLong(ByVal value As Long) As ColorType
    PI_ColorFromLong.red = value And &HFF&
    PI_ColorFromLong.green = (value \ 256) And &HFF&
    PI_ColorFromLong.blue = (value \ 65536) And &HFF&
End Function

Private Function PI_MixRGB(ByVal foreground As Long, ByVal background As Long, ByVal amount As Double) As Long
    Dim red As Long, green As Long, blue As Long
    red = CLng((foreground And &HFF&) * amount + (background And &HFF&) * (1 - amount))
    green = CLng(((foreground \ 256) And &HFF&) * amount + ((background \ 256) And &HFF&) * (1 - amount))
    blue = CLng(((foreground \ 65536) And &HFF&) * amount + ((background \ 65536) And &HFF&) * (1 - amount))
    PI_MixRGB = RGB(red, green, blue)
End Function

Public Sub PI_ResolveThemeOptions(targetSlide As Slide, opts As ProgressDotsOptions)
    Dim scheme As ThemeColorScheme
    Dim accent As Long, light As Long, dark As Long, background As Long, surface As Long, foreground As Long
    If opts.UseThemeColors <> 1 Then Exit Sub
    Set scheme = targetSlide.Design.SlideMaster.Theme.ThemeColorScheme
    accent = scheme.Colors(msoThemeAccent1).RGB
    light = scheme.Colors(msoThemeLight1).RGB
    dark = scheme.Colors(msoThemeDark1).RGB
    background = light
    If targetSlide.Background.Fill.Type = msoFillSolid Then background = targetSlide.Background.Fill.ForeColor.RGB
    If PI_ContrastColor(background) = RGB(255, 255, 255) Then
        surface = PI_MixRGB(light, dark, 0.12)
    Else
        surface = PI_MixRGB(dark, light, 0.06)
    End If
    foreground = PI_ContrastColor(surface)
    If ColorToRGB(opts.seenTextColorBG) = ColorToRGB(opts.unseenTextColorBG) Then
        opts.seenTextColorBG = PI_ColorFromLong(surface)
    Else
        opts.seenTextColorBG = PI_ColorFromLong(PI_MixRGB(foreground, surface, 0.11))
    End If
    opts.unseenTextColorBG = PI_ColorFromLong(surface)
    opts.CurrentTextColorBG = PI_ColorFromLong(accent)
    opts.seenTextColor = PI_ColorFromLong(foreground)
    opts.unseenTextColor = PI_ColorFromLong(foreground)
    opts.CurrentTextColor = PI_ColorFromLong(PI_ContrastColor(accent))
    opts.seenDotColor = PI_ColorFromLong(foreground)
    opts.unseenDotColor = PI_ColorFromLong(foreground)
    If PI_UsesSectionBand(opts) Then
        opts.CurrentDotColor = PI_ColorFromLong(PI_ContrastColor(accent))
    Else
        opts.CurrentDotColor = PI_ColorFromLong(accent)
    End If
    opts.AccentColor = PI_ColorFromLong(accent)
End Sub

Public Function PI_OptionsForSlide(opts As ProgressDotsOptions, targetSlide As Slide) As ProgressDotsOptions
    Dim resolved As ProgressDotsOptions, ruleIndex As Long
    resolved = opts
    PI_EnsureRuleContext targetSlide
    If mPIRulesEnabled Then
        ruleIndex = PI_MatchingRule(targetSlide)
        If ruleIndex > 0 Then
            If mPIRules(ruleIndex).show Then resolved = PI_RuleVisualOptions(mPIRuleOptions(ruleIndex), opts)
        End If
    End If
    PI_ResolveThemeOptions targetSlide, resolved
    If resolved.DockEdge <> 0 Then resolved.SectionBandBackdrop = 0
    PI_OptionsForSlide = resolved
End Function

Public Sub PI_ApplyDock(ByRef navigation As Shape, targetSlide As Slide, opts As ProgressDotsOptions)
    Dim members As ShapeRange, part As Shape
    Dim tagNames() As String, tagValues() As String
    Dim tagCount As Long, index As Long, originalZ As Long
    Dim currentZ As Long, previousZ As Long, zSteps As Long, zLimit As Long
    Dim originalName As String, role As String
    Dim titleTop As Double, dotsTop As Double, bandHeight As Double
    Dim pageWidth As Double, pageHeight As Double, candidateHeight As Double
    Dim failureDescription As String
    If navigation Is Nothing Then Exit Sub
    If opts.DockEdge = 0 Then Exit Sub
    On Error GoTo DockError
    pageWidth = targetSlide.Parent.PageSetup.SlideWidth
    pageHeight = targetSlide.Parent.PageSetup.SlideHeight
    PI_BandGeometry opts, titleTop, dotsTop, bandHeight
    originalName = navigation.Name
    originalZ = navigation.ZOrderPosition
    tagCount = navigation.Tags.Count
    If tagCount > 0 Then
        ReDim tagNames(1 To tagCount)
        ReDim tagValues(1 To tagCount)
        For index = 1 To tagCount
            tagNames(index) = navigation.Tags.Name(index)
            tagValues(index) = navigation.Tags.Value(index)
        Next index
    End If
    If navigation.Type = msoGroup Then
        Set members = navigation.Ungroup
    Else
        Set members = targetSlide.Shapes.Range(Array(navigation.Name))
    End If
    If Not PI_UsesSectionBand(opts) Then
        bandHeight = 0
        For index = 1 To members.Count
            role = LCase$(members(index).Tags("PI_ROLE"))
            If role = "section-title" Or role = "section-counter" Or role = "indicator" Or role = "indicator-summary" Then
                candidateHeight = members(index).Top + members(index).Height - opts.yMargin
                If candidateHeight > bandHeight Then bandHeight = candidateHeight
            End If
        Next index
    End If
    If opts.DockEdge >= 2 And bandHeight < opts.SidebarWidth Then bandHeight = opts.SidebarWidth
    For index = 1 To members.Count
        Set part = members(index)
        role = LCase$(part.Tags("PI_ROLE"))
        If opts.DockEdge = 1 Then
            If role = "section-title" Or role = "section-counter" Or role = "indicator" Or role = "indicator-summary" Then
                If opts.PresetSize > 0 And PI_UsesSectionBand(opts) Then
                    part.Top = pageHeight - part.Top - part.Height
                Else
                    part.Top = part.Top + pageHeight - 2 * opts.yMargin - bandHeight
                End If
            Else
                PI_MapDockSurface part, opts.DockEdge, pageWidth, pageHeight
            End If
        ElseIf role = "section-title" Then
            PI_DockTitle part, opts, pageWidth, bandHeight
        ElseIf role = "section-counter" Then
            PI_DockCounter part, opts, pageWidth, bandHeight
        ElseIf role = "indicator" Or role = "indicator-summary" Then
            PI_DockMarker part, opts, pageWidth, bandHeight
        Else
            PI_MapDockSurface part, opts.DockEdge, pageWidth, pageHeight
        End If
    Next index
    If members.Count > 1 Then
        Set navigation = members.Group
    Else
        Set navigation = members(1)
    End If
    navigation.Tags.Add "__progress-indicator__", "yes"
    navigation.Name = originalName
    For index = 1 To tagCount
        navigation.Tags.Add tagNames(index), tagValues(index)
    Next index
    navigation.Tags.Add "PI_DOCK", CStr(opts.DockEdge)
    navigation.Shadow.Visible = msoFalse
    ' Each move must approach the target within one traversal of the collection.
    zLimit = targetSlide.Shapes.Count
    currentZ = navigation.ZOrderPosition
    Do While currentZ > originalZ
        If zSteps >= zLimit Then Err.Raise vbObjectError + 3023, "PI_ApplyDock", _
            PI_Zh("5BFC822A5C427EA756DE4F4D59318D25FF1A8C0365746B6565708D858FC75F6272B6657091CF3002")
        previousZ = currentZ
        navigation.ZOrder msoSendBackward
        zSteps = zSteps + 1
        currentZ = navigation.ZOrderPosition
        If currentZ >= previousZ Then Err.Raise vbObjectError + 3023, "PI_ApplyDock", _
            PI_Zh("5BFC822A5C427EA756DE4F4D59318D25FF1A8C036574540E5C427EA76CA16709541176EE680779FB52A83002")
    Loop
    Do While currentZ < originalZ
        If zSteps >= zLimit Then Err.Raise vbObjectError + 3023, "PI_ApplyDock", _
            PI_Zh("5BFC822A5C427EA756DE4F4D59318D25FF1A8C0365746B6565708D858FC75F6272B6657091CF3002")
        previousZ = currentZ
        navigation.ZOrder msoBringForward
        zSteps = zSteps + 1
        currentZ = navigation.ZOrderPosition
        If currentZ <= previousZ Then Err.Raise vbObjectError + 3023, "PI_ApplyDock", _
            PI_Zh("5BFC822A5C427EA756DE4F4D59318D25FF1A8C036574540E5C427EA76CA16709541176EE680779FB52A83002")
    Loop
    Exit Sub
DockError:
    failureDescription = Err.Description
    Err.Raise vbObjectError + 3023, "PI_ApplyDock", failureDescription
End Sub

Private Sub PI_MapDockSurface(ByVal surface As Shape, ByVal edge As Integer, _
                              ByVal pageWidth As Double, ByVal pageHeight As Double)
    Dim width As Double, height As Double, centerX As Double, centerY As Double
    Dim mappedX As Double, mappedY As Double
    width = surface.Width
    height = surface.Height
    centerX = surface.Left + width / 2
    centerY = surface.Top + height / 2
    Select Case edge
        Case 1
            mappedX = centerX
            mappedY = pageHeight - centerY
            surface.Flip msoFlipVertical
        Case 2
            mappedX = centerY
            mappedY = centerX
            surface.Flip msoFlipVertical
            surface.Rotation = 90
        Case 3
            mappedX = pageWidth - centerY
            mappedY = centerX
            surface.Rotation = 90
        Case Else
            Exit Sub
    End Select
    surface.Left = mappedX - width / 2
    surface.Top = mappedY - height / 2
End Sub

Private Sub PI_DockTitle(ByVal title As Shape, opts As ProgressDotsOptions, _
                         ByVal pageWidth As Double, ByVal bandHeight As Double)
    Dim alongStart As Double, alongLength As Double, padding As Double, alongPadding As Double
    Dim markerWidth As Double, textWidth As Double, gap As Double
    alongStart = title.Left
    alongLength = title.Width
    padding = PI_BarClamp(opts.SectionBandPadding, 0, bandHeight / 4)
    alongPadding = PI_BarClamp(padding, 0, alongLength / 4)
    markerWidth = PI_BarClamp(PI_MarkerHeight(opts), 0, bandHeight / 3)
    gap = PI_BarClamp(opts.SectionIndicatorGap, 0, bandHeight / 3)
    If LCase$(opts.PIStyle) = "sections only" Then
        markerWidth = 0
        gap = 0
    End If
    textWidth = bandHeight - 2 * padding - markerWidth - gap
    If textWidth < 12 Then textWidth = 12
    title.Rotation = 0
    title.LockAspectRatio = msoFalse
    title.Width = textWidth
    title.Height = alongLength - PI_BarClamp(PI_CounterRowHeight(opts), 0, alongLength / 3)
    title.Top = alongStart
    If opts.DockEdge = 2 Then
        title.Left = opts.yMargin + padding
    Else
        title.Left = pageWidth - opts.yMargin - bandHeight + padding + markerWidth + gap
    End If
    With title.TextFrame2
        .Orientation = msoTextOrientationHorizontal
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = alongPadding
        .MarginBottom = alongPadding
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoTrue
        .AutoSize = msoAutoSizeTextToFitShape
        .TextRange.Font.Size = opts.textHeight
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub

Private Sub PI_DockMarker(ByVal marker As Shape, opts As ProgressDotsOptions, _
                          ByVal pageWidth As Double, ByVal bandHeight As Double)
    Dim alongStart As Double, originalWidth As Double, originalHeight As Double
    Dim padding As Double, inset As Double, markerWidth As Double, markerHeight As Double
    Dim left As Double, top As Double, markerScale As Double
    alongStart = marker.Left
    originalWidth = marker.Width
    originalHeight = marker.Height
    If LCase$(marker.Tags("PI_OUTLINED")) = "yes" Then
        inset = marker.Line.Weight / 2
        alongStart = alongStart - inset
        originalWidth = originalWidth + 2 * inset
        originalHeight = originalHeight + 2 * inset
    End If
    padding = PI_BarClamp(opts.SectionBandPadding, 0, bandHeight / 4)
    If opts.FlatBar = 1 Or LCase$(marker.Tags("PI_MARKER")) = "oval" Then
        markerWidth = originalHeight
        markerHeight = originalWidth
        top = alongStart
    Else
        markerWidth = originalWidth
        markerHeight = originalHeight
        If markerWidth > bandHeight / 3 Then
            markerScale = (bandHeight / 3) / markerWidth
            markerWidth = markerWidth * markerScale
            markerHeight = markerHeight * markerScale
        End If
        top = alongStart + (originalWidth - markerHeight) / 2
    End If
    If opts.DockEdge = 2 Then
        left = opts.yMargin + bandHeight - padding - markerWidth
    Else
        left = pageWidth - opts.yMargin - bandHeight + padding
    End If
    marker.Rotation = 0
    marker.LockAspectRatio = msoFalse
    marker.Left = left + inset
    marker.Top = top + inset
    marker.Width = markerWidth - 2 * inset
    marker.Height = markerHeight - 2 * inset
End Sub



' ----------------------------------------------------------------
' Procedure Name: IsMacPlatform
' Purpose: Trying multiple way to see if we are working on a MAC
' Return Type: Boolean
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2025-06-23
' ----------------------------------------------------------------
Function IsMacPlatform() As Boolean
    On Error Resume Next
    
    ' First try the Environ method (most reliable for Add-Ins)
    If Environ("HOME") <> "" And Environ("USERPROFILE") = "" Then
        IsMacPlatform = True
        Exit Function
    End If
    
    ' Second try checking for Mac directory structure
    If Dir("/Users", vbDirectory) <> "" Then
        IsMacPlatform = True
        Exit Function
    End If
    
    ' Third try MacScript availability test
    MacScript ("return 1")
    If Err.Number = 0 Then
        IsMacPlatform = True
        Exit Function
    End If
    
    ' Detect if running on Mac
    If (Application.OperatingSystem Like "*Mac*") Then
        IsMacPlatform = True
        Exit Function
    End If
    
    ' If all else fails, assume Windows
    IsMacPlatform = False
    
    On Error GoTo 0
End Function


' Unicode UI text is stored as UTF-16 code units so VBA does not depend on
' the host system ANSI code page. Internal configuration keys stay unchanged.
Public Function PI_Zh(ByVal hexText As String) As String
    Dim i As Long
    Dim codeUnit As Long
    Dim result As String
    For i = 1 To Len(hexText) Step 4
        codeUnit = CLng("&H" & Mid$(hexText, i, 4))
        If codeUnit > 32767 Then codeUnit = codeUnit - 65536
        result = result & ChrW(codeUnit)
    Next i
    PI_Zh = result
End Function

Public Sub PI_SetFormFont(ByVal form As Object)
    Dim ctl As Object
    Dim family As String
    On Error Resume Next
    #If Mac Then
        family = "PingFang SC"
    #Else
        family = "Microsoft YaHei"
    #End If
    form.Font.Name = family
    For Each ctl In form.Controls
        Select Case TypeName(ctl)
            Case "Label", "CommandButton", "CheckBox", "TextBox", "ComboBox", "Frame"
                If ctl.Name <> "lblProgressbar" Then ctl.Font.Name = family
        End Select
    Next ctl
    On Error GoTo 0
End Sub

Public Function PI_ColorTitle(ByVal key As String) As String
    Select Case key
        Case "Color Font Seen"
            PI_ColorTitle = PI_Zh("5DF28BB27AE0828265875B57")
        Case "Color Font Unseen"
            PI_ColorTitle = PI_Zh("672A8BB27AE0828265875B57")
        Case "Color Font Active"
            PI_ColorTitle = PI_Zh("5F53524D7AE0828265875B57")
        Case "Color Font Seen Background"
            PI_ColorTitle = PI_Zh("5DF28BB27AE0828280CC666F")
        Case "Color Font Unseen Background"
            PI_ColorTitle = PI_Zh("672A8BB27AE0828280CC666F")
        Case "Color Font Active Background"
            PI_ColorTitle = PI_Zh("5F53524D7AE0828280CC666F")
        Case "Color Indicator Seen"
            PI_ColorTitle = PI_Zh("5DF28BB28FDB5EA656FE6807")
        Case "Color Indicator Unseen"
            PI_ColorTitle = PI_Zh("672A8BB28FDB5EA656FE6807")
        Case "Color Indicator Active"
            PI_ColorTitle = PI_Zh("5F53524D8FDB5EA656FE6807")
        Case Else
            PI_ColorTitle = key
    End Select
End Function
