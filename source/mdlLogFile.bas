Attribute VB_Name = "mdlLogFile"
' ------------------------------------------------------
' Name: mdlLogFile
' Kind: Module
' Purpose: Creates a Logfile in the current directory
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-22
' ------------------------------------------------------
Option Explicit

Public gblLogFileName As String

' ----------------------------------------------------------------
' Procedure Name: LogInfo
' Purpose: Adds the LogMessage Information to the logfile
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter LogMessage (String): Value / text to write to the logfile
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-22
' ----------------------------------------------------------------
Public Sub LogInfo(LogMessage As String)
10        On Error GoTo ErrorHandler
          
          'set path and name of the log file where you want to save
          'the log file
          Dim FileNum As Integer
              
20        LogMessage = Now() & ";" & LogMessage
30        If gblLogFileName & "" = "" Then
40            DefineLogFileName
50        End If
60        FileNum = FreeFile ' next file number
70        Open gblLogFileName For Append As #FileNum ' creates the file if it doesn't exist
80        Print #FileNum, LogMessage ' write information at the end of the text file
90        Close #FileNum ' close the file

          'Debug.Print LogMessage       'Attention: VBA on MAC seems to crash PPT when there is too much data in the direct window!
exit_proc:
100       Exit Sub

          
ErrorHandler:
110       Select Case Err.Number

              Case Else
120               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl & vbCrLf & PI_Zh("6B6495198BEF672A5199516565E55FD7FF0C8BF7622A56FE4FDD755995198BEF4FE1606F3002"), , "LogInfo"
130               Resume Next
140               Resume
              
150       End Select
End Sub

' ----------------------------------------------------------------
' Procedure Name: DefineLogFileName
' Purpose: Define the default logfile name including active path
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-22
' ----------------------------------------------------------------
Public Sub DefineLogFileName()
10        On Error GoTo ErrorHandler
          
          Dim sPath As String
          Dim pathSep             As String

20        pathSep = ":"
#If Mac Then
30        pathSep = ":"
#Else
40        pathSep = "\"
#End If
          
50        sPath = ActivePresentation.Path
60        gblLogFileName = sPath & pathSep & "LogFile_PI_" & Format(Now(), "YYYYmmddhhMMss") & ".log"    'Default setzen
          
exit_proc:
70        Exit Sub
          
ErrorHandler:
80        Select Case Err.Number

              Case Else
90                MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl & vbCrLf & PI_Zh("6B6495198BEF672A5199516565E55FD7FF0C8BF7622A56FE4FDD755995198BEF4FE1606F3002"), , "DefineLogFileName"
100               Resume Next
110               Resume
              
120       End Select
End Sub
