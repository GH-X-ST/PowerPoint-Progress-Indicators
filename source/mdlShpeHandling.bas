Attribute VB_Name = "mdlShpeHandling"
' ------------------------------------------------------
' Name: mdlShpeHandling
' Kind: Modul
' Purpose: Collectiion module for procs handling the shapes
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ------------------------------------------------------
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: retShapeID
' Purpose: return ppt internal shapeid for the selected shape
' Procedure Kind: Function
' Procedure Access: Public
' Parameter intPosition (Integer): dropdown index (0..196)
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function retShapeID(intPosition As Integer)
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct retShapeID: " & intPosition         'Debug: Write active procedure to log
30        Randomize Timer
40        retShapeID = "9"    'default in case of an error
50        If intPosition = 14 Then            'to do: if user chooses spacer a random symbol will be used
60            retShapeID = gblvarShapeListIDs(Round(Rnd() * 183, 0) + 14)
70        Else
80            retShapeID = gblvarShapeListIDs(intPosition + 1)
90        End If

exit_proc:
100       Exit Function
          
ErrorHandler:
110       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
120       Select Case Err.Number
              Case -2147188160      'no active presentation
130               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "retShapeID: Nothing to do!"
140               Resume exit_proc
150               Resume
160           Case 13 'maybe arrays not yet defined
170               Initialize
180               Resume
                  
190           Case Else
200               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "retShapeID"
210               Resume Next
220               Resume
              
230       End Select
End Function

' ----------------------------------------------------------------
' Procedure Name: retShapeName
' Purpose: return shapename from the position in the list
' Procedure Kind: Function
' Procedure Access: Public
' Parameter intPosition (Integer): drop down position (0...196)
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function retShapeName(intPosition As Integer)
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct retShapeName: " & intPosition         'Debug: Write active procedure to log
          
30        retShapeName = gblvarShapeListNames(intPosition + 1)

exit_proc:
40        Exit Function
          
ErrorHandler:
50        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
60        Select Case Err.Number
              Case -2147188160      'no active presentation
70                MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "retShapeName: Nothing to do!"
80                Resume exit_proc
90                Resume
100           Case 13 'maybe arrays not yet defined
110               Initialize
120               Resume
              
130           Case Else
140               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "retShapeName"
150               Resume Next
160               Resume
              
170       End Select
End Function

