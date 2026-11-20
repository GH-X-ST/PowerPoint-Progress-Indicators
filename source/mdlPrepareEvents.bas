Attribute VB_Name = "mdlPrepareEvents"
Option Explicit

Dim cPPTObject As New clsPresentationEvents
Public gblTrapFlag As Boolean
'src: http://skp.mvps.org/downloads/pptevent.zip

Public Sub TrapEvents()
10        If gblTrapFlag = True Then
              'MsgBox "Relax, my friend, the EventHandler is already active.", vbInformation + vbOKOnly, "PowerPoint Event Handler Example"
20            Exit Sub
30        End If
40        Set cPPTObject.PPTEvent = Application
50        gblTrapFlag = True
End Sub

Public Sub ReleaseTrap()
10        If gblTrapFlag = True Then
20            Set cPPTObject.PPTEvent = Nothing
30            Set cPPTObject = Nothing
40            gblTrapFlag = False
50        End If
End Sub
