Attribute VB_Name = "clsPresentationEvents"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Const COMPANY_NAME = "datenbank-projekt.de"

Public WithEvents PPTEvent As Application
Attribute PPTEvent.VB_VarHelpID = -1

'----------------
'src: http://skp.mvps.org/downloads/pptevent.zip

Private Sub Class_Initialize()
          'MsgBox "This demo traps some Presentation events:" & vbCrLf & _
          '       "NewPresentation, PresentationOpen, PresentationNewSlide, PresentationSave and PresentationPrint" & vbCrLf & vbCrLf & _
          '       "Some Slide Show Events like: " & vbCrLf & _
          '       "SlideShowBegin & SlideShowEnd" & vbCrLf & vbCrLf & _
          '       "And Slide View mode Window events like :" & vbCrLf & _
          '       "WindowSelectionChange", vbInformation + vbOKOnly, _
          '       "The EventHandler class has been initialized."
10        RefreshRibbon
End Sub

Private Sub Class_Terminate()
10        RefreshRibbon
          'MsgBox "EventHandler is now inactive.", vbInformation + vbOKOnly, "PowerPoint Event Handler Example"
End Sub

Private Sub PPTEvent_WindowActivate(ByVal Pres As Presentation, ByVal Wn As DocumentWindow)
10        RefreshRibbon
          'MsgBox "The WindowActivate event occurs after any application or document window is activated." & _
          '       vbCrLf & _
          '       Pres.Name & " is the name of the activated window", vbOKOnly + vbInformation, "PowerPoint Event Handler Example"
End Sub

Private Sub PPTEvent_NewPresentation(ByVal Pres As Presentation)
10        RefreshRibbon

          'If Pres.HasTitleMaster Then
          '        With Pres.TitleMaster.HeadersFooters
          '            With .Footer
          '                .Text = COMPANY_NAME
          '                .visible = msoTrue
          '            End With
          '        End With
          '    End If
          '    With Pres.SlideMaster.HeadersFooters
          '        With .Footer
          '            .Text = COMPANY_NAME
          '            .visible = msoTrue
          '        End With
          '    End With
          'MsgBox "The NewPresentation event is fired whenever you create a new presentation. In this example" & _
          '       " we make use of the particular event to automatically add the Company's name as a footer" & _
          '       " in the Slide/Title Master.", vbOKOnly + vbInformation, "PowerPoint Event Handler Example"
          
End Sub

Private Sub PPTEvent_PresentationClose(ByVal Pres As Presentation)
10        RefreshRibbon
          'MsgBox "The PresentationClose event occurs before the presentation " & _
          '       "has been closed and removed from the Presentations collection" & vbCrLf & _
          '       Pres.Name & " is the name of the opened presentation", vbOKOnly + vbInformation, "PowerPoint Event Handler Example"

End Sub

Private Sub PPTEvent_PresentationNewSlide(ByVal Sld As slide)
    'If ActiveWindow.ViewType <> ppViewSlide Then
    '    ActiveWindow.ViewType = ppViewSlide
    'End If
    'Sld.Shapes.AddComment.Select
    'With ActiveWindow.Selection.TextRange
    '    .Text = "Slide creation date: "
    '    .Characters(Start:=22, Length:=0).InsertDateTime DateTimeFormat:=ppDateTimeddddMMMMddyyyy, InsertAsField:=msoFalse
    'End With
    'MsgBox "PresentationNewSlide event is fired when a new slide is inserted into the presentation. " & _
    '       "In this example we make use of the particular event to add a comment which stores the slide creation date", _
    '       vbOKOnly + vbInformation, "PowerPoint Event Handler Example"

End Sub

Private Sub PPTEvent_PresentationOpen(ByVal Pres As Presentation)
10        RefreshRibbon
          'MsgBox "The PresentationOpen occurs after the presentation " & _
          '       "has been opened and added to the Presentations collection" & vbCrLf & _
          '       Pres.Name & " is the name of the opened presentation", vbOKOnly + vbInformation, "PowerPoint Event Handler Example"
End Sub


Private Sub PPTEvent_PresentationPrint(ByVal Pres As Presentation)
    'Pres.PrintOptions.PrintHiddenSlides = True
    'MsgBox "The PresentationPrint occurs just before a print job is sent." & vbCrLf & _
    '       "In this example we ensure that the presentation will print the hidden slides too while printing.", vbOKOnly + vbInformation, "PowerPoint Event Handler Example"

End Sub

Private Sub PPTEvent_PresentationSave(ByVal Pres As Presentation)
    'Debug.Print Pres.Saved
    'MsgBox "The PresentationSave occured. Useful to perform some task(s) just prior to saving the presentation.", vbOKOnly + vbInformation, "PowerPoint Event Handler Example"

End Sub

Private Sub PPTEvent_SlideShowBegin(ByVal Wn As SlideShowWindow)
    'MsgBox "You've just begun the " & Wn.Presentation.Name & " show. The SlideShowBegin event was fired." & _
    '       "This event is NOT fired again if the current show branches into another show. It is fired only ONCE.", vbOKOnly + vbInformation, "PowerPoint Event Handler Example"
End Sub

Private Sub PPTEvent_SlideShowEnd(ByVal Pres As Presentation)
    'MsgBox "Show: " & Pres.Name & " just ended. Hope you had a happy viewing experience", vbOKOnly + vbInformation, "PowerPoint Event Handler Example"

End Sub

Private Sub PPTEvent_WindowBeforeRightClick(ByVal Sel As Selection, Cancel As Boolean)
    'Dim UserChoice As Integer
    'UserChoice = MsgBox("You have Right-clicked! If you want to display the context menu click YES else Click No", vbQuestion + vbYesNo, "PowerPoint Event Handler Example")
    'If UserChoice = vbYes Then
    '    Cancel = False
    'Else
    '    Cancel = True
    'End If
End Sub
