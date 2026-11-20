Attribute VB_Name = "mdlFileHandling"
' ------------------------------------------------------
' Name: mdlFileHandling
' Kind: Modul
' Purpose: Read and write settings from/to a file
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2025-06-19
' ------------------------------------------------------
Option Explicit

#If Mac Then
#Else
' OPENFILENAMEW uses native pointers on both supported Windows VBA7 platforms.
Private Type PIConfigSaveDialog
    byteSize As Long
#If Win64 Then
    alignOwner As Long
#End If
    ownerHwnd As LongPtr
    instanceHandle As LongPtr
    filterPointer As LongPtr
    customFilterPointer As LongPtr
    maxCustomFilterChars As Long
    filterIndex As Long
    filePointer As LongPtr
    maxFileChars As Long
#If Win64 Then
    alignFileTitle As Long
#End If
    fileTitlePointer As LongPtr
    maxFileTitleChars As Long
#If Win64 Then
    alignInitialDir As Long
#End If
    initialDirPointer As LongPtr
    titlePointer As LongPtr
    dialogFlags As Long
    fileOffset As Integer
    extensionOffset As Integer
    defaultExtPointer As LongPtr
    customData As LongPtr
    hookPointer As LongPtr
    templatePointer As LongPtr
    reservedPointer As LongPtr
    reservedValue As Long
    extendedFlags As Long
End Type

Private Declare PtrSafe Function PI_ActiveWindowHandle Lib "user32.dll" Alias "GetActiveWindow" () As LongPtr
Private Declare PtrSafe Function PI_GetSaveFileNameW Lib "comdlg32.dll" Alias "GetSaveFileNameW" (ByRef dialog As PIConfigSaveDialog) As Long
Private Declare PtrSafe Function PI_CommonDialogError Lib "comdlg32.dll" Alias "CommDlgExtendedError" () As Long

Private Function PI_PickConfigSavePath(ByVal suggestedFile As String) As String
    Const maxChars As Long = 32768
    Const overwritePrompt As Long = &H2
    Const keepDirectory As Long = &H8
    Const requireDirectory As Long = &H800
    Const explorerDialog As Long = &H80000
    Dim dialog As PIConfigSaveDialog, filenameBuffer As String, title As String, filter As String
    Dim defaultExt As String, directory As String, baseName As String, lastSlash As Long, dialogError As Long
    lastSlash = InStrRev(suggestedFile, "\")
    directory = Left$(suggestedFile, lastSlash)
    baseName = Mid$(suggestedFile, lastSlash + 1)
    filenameBuffer = String$(maxChars, vbNullChar)
    Mid$(filenameBuffer, 1, Len(baseName)) = baseName
    defaultExt = "txt"
    title = PI_Zh("4FDD5B588FDB5EA66761914D7F6E")
    filter = PI_Zh("6587672C65874EF6") & " (*.txt)" & vbNullChar & "*.txt" & vbNullChar & vbNullChar
    With dialog
        .byteSize = LenB(dialog)
        .ownerHwnd = PI_ActiveWindowHandle()
        .filterPointer = StrPtr(filter)
        .filterIndex = 1
        .filePointer = StrPtr(filenameBuffer)
        .maxFileChars = maxChars
        If Len(directory) > 0 Then .initialDirPointer = StrPtr(directory)
        .titlePointer = StrPtr(title)
        .defaultExtPointer = StrPtr(defaultExt)
        .dialogFlags = overwritePrompt Or keepDirectory Or requireDirectory Or explorerDialog
    End With
    If PI_GetSaveFileNameW(dialog) = 0 Then
        dialogError = PI_CommonDialogError()
        If dialogError <> 0 Then Err.Raise vbObjectError + 3090, "GetSaveFileNameW", "CommDlgExtendedError 0x" & Hex$(dialogError)
        Exit Function
    End If
    PI_PickConfigSavePath = Left$(filenameBuffer, InStr(filenameBuffer, vbNullChar) - 1)
End Function
#End If


' ----------------------------------------------------------------
' Procedure Name: SettingsExport
' Purpose: Export PI settings to a file for later use in a different file
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2025-06-20
' ----------------------------------------------------------------
Public Function SettingsExport(Optional ByVal suggestedFile As String = "")
10        If gblbolDebug = True Then LogInfo "Fct SettingsExport"         'Debug: Write active procedure to log
20        SaveTextToFile ReadPersistensPISettings(), suggestedFile
End Function

' ----------------------------------------------------------------
' Procedure Name: SettingsImport
' Purpose: Import (and apply) PI settings from a file
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2025-06-20
' ----------------------------------------------------------------
Public Function SettingsImport()
          Dim s As String
10        s = ReadSingleLineFromFile()
20        If s <> "" Then
30            PersistPISettings s
40            gblCreatePressed = True 'we force the update
50            RefreshProgressBar
60        Else
70            MsgBox PI_Zh("65874EF64E3A7A7A3002"), vbCritical, PI_Zh("65874EF64E3A7A7A")
80        End If

90        If gblbolDebug = True Then
100           LogInfo "Fct SettingsImport"         'Debug: Write active procedure to log
110           LogInfo "s: " & s
120       End If
          
End Function

' ----------------------------------------------------------------
' Procedure Name: PersistPISettings
' Purpose: Write settings to the current presentation
' Parameter s (String): prepared settings string use: s = SerializeProgressOptionsRibbon
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2025-06-20
' ----------------------------------------------------------------
Public Function PersistPISettings(s As String)
10        If gblbolDebug = True Then
20            LogInfo "Fct PersistPISettings"         'Debug: Write active procedure to log
30            LogInfo "s: " & s
40        End If

50        ActivePresentation.Tags.Add "__progress-indicator-config__", s
End Function

' ----------------------------------------------------------------
' Procedure Name: ReadPersistensPISettings
' Purpose: Read the PI settings from the current presentation
' Return Type: String
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2025-06-20
' ----------------------------------------------------------------
Public Function ReadPersistensPISettings(Optional ByVal presentation As Presentation = Nothing) As String
          Dim s As String
          If presentation Is Nothing Then Set presentation = ActivePresentation
10        s = presentation.Tags.Item("__progress-indicator-config__")
          
20        If Len(s) = 0 Then s = ""
          
30        If gblbolDebug = True Then
40            LogInfo "Fct ReadPersistensPISettings"         'Debug: Write active procedure to log
50            LogInfo "s: " & s
60        End If

70        ReadPersistensPISettings = s
End Function


' ----------------------------------------------------------------
' Procedure Name: SaveTextToFile
' Purpose: Save a string to a textfile. Note: An existing file will be overwritten. The text will NOT be appended!
' Parameter textToSave (String):
' Parameter sFileName (String):
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2025-06-20
' ----------------------------------------------------------------
Public Sub SaveTextToFile(textToSave As String, Optional sFileName As String = "")
    Dim filePath As String
    Dim fileNumber As Integer
    'Dim isMac As Boolean
    
    If sFileName = "" Then sFileName = "ProgressIndicatorSettings_" & Format(Now(), "YYYYMMDDHHmmSS") & ".txt"
    
    ' Detect if running on Mac
    'isMac = (Application.OperatingSystem Like "*Mac*")
    
    ' Prompt for file path
#If Mac Then
    'If isMac Then
    ' ==== macOS: Use AppleScript if available via MacScript ====
    On Error Resume Next
    '2025-07-02 commented out for testing
    '        filePath = MacScript("return POSIX path of (choose file name with prompt ""Save your file as:"" default name """ & sFileName & """)")
    '        If Err.Number <> 0 Or filePath = "" Then
    '            On Error GoTo 0
    '            MsgBox "MacScript does not seem to work. You will be presented with an inputbox to paste the path and filename where you want to save the file.", vbExclamation, "Sorry MAC user"
    '            filePath = InputBox("Enter full path and file name where the text should be saved (e.g. /Users/you/Documents/" & sFileName & "):", "Save As", "/Users/" & Environ("USER") & "/Documents/" & sFileName)
    '        End If

    '2025-07-02 inserted for testing
    filePath = InputBox(PI_Zh("8BF7590D52364E0B65B9914D7F6E5B577B264E32FF0C5E7681EA884C4FDD5B584E3A6587672C65874EF6FF0C4EE54FBF4E4B540E5BFC51653002"), PI_Zh("004D006100630020914D7F6E5BFC51654E0E5BFC51FA"), textToSave)
    GoTo proc_exit
    On Error GoTo 0
        
#Else
    On Error GoTo ErrorHandler
    filePath = PI_PickConfigSavePath(sFileName)
#End If
    
    ' Exit if no file selected
    If filePath = "" Then
        MsgBox PI_Zh("672A900962E965874EF6FF0C64CD4F5C5DF253D66D883002"), vbInformation
        GoTo proc_exit
    End If
    
#If Mac Then
    ' Make it a .txt file!
    If Len(filePath) > 4 Then
        If LCase(Right(filePath, 4)) <> ".txt" Then
            filePath = filePath & ".txt"
        End If
    Else
        MsgBox PI_Zh("65874EF6540D65E06548FF0C64CD4F5C5DF253D66D883002"), vbInformation
        GoTo proc_exit
    End If
    
#Else
    If LCase$(Right$(filePath, 4)) <> ".txt" Then
        MsgBox PI_Zh("8BF7900962E90020002E00740078007400206587672C65874EF63002"), vbExclamation
        GoTo proc_exit
    End If
#End If

    ' Write text to file (overwrite mode)
    On Error GoTo ErrorHandler
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    Print #fileNumber, textToSave
    Close #fileNumber
    MsgBox PI_Zh("914D7F6E5DF24FDD5B585230FF1A") & vbCrLf & filePath, vbInformation

proc_exit:
    Exit Sub

ErrorHandler:
    MsgBox PI_Zh("5199516565874EF659318D25FF1A") & Err.description, vbCritical, PI_Zh("4FDD5B58914D7F6E59318D25")
    If fileNumber <> 0 Then Close #fileNumber
    If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl
            Resume proc_exit
            Resume
    End Select
    
End Sub


' ----------------------------------------------------------------
' Procedure Name: ReadSingleLineFromFile
' Purpose: Read a single line (settings) from a file which the user selects
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2025-06-20
' ----------------------------------------------------------------
Public Function ReadSingleLineFromFile() As String
    Dim sResult As String
    Dim filePath As String
    Dim fileNumber As Integer
    Dim fileContent As String
    'Dim isMac As Boolean
    
    ' Detect if running on Mac
    'isMac = (Application.OperatingSystem Like "*Mac*")
    
    ' Prompt for file path
#If Mac Then
    'If isMac Then
    sResult = InputBox(PI_Zh("8BF757284E0B65B97C988D344E4B524D5BFC51FA76845B8C6574914D7F6E5B577B264E323002"), PI_Zh("004D006100630020914D7F6E5BFC51654E0E5BFC51FA"))
#Else
    filePath = PickFile()
        
    ' Exit if no file selected
    If filePath = "" Then
        MsgBox PI_Zh("672A900962E965874EF6FF0C64CD4F5C5DF253D66D883002"), vbInformation
        sResult = ""
        GoTo proc_exit
    End If
        
    ' Read the single line from the file
    On Error GoTo ErrorHandler
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
        
    ' Check if file is empty
    If Not EOF(fileNumber) Then
        Line Input #fileNumber, fileContent
        sResult = fileContent
    Else
        sResult = ""
        MsgBox PI_Zh("6240900965874EF64E3A7A7A3002"), vbInformation
    End If
        
    Close #fileNumber
#End If
    
proc_exit:
    ReadSingleLineFromFile = sResult
    Exit Function
    
ErrorHandler:
    MsgBox PI_Zh("8BFB53D665874EF659318D25FF1A") & Err.description, vbCritical, PI_Zh("8BFB53D6914D7F6E59318D25")
    If fileNumber <> 0 Then Close #fileNumber
    If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
    Select Case Err.Number
        Case 53 ' File not found
            MsgBox PI_Zh("627E4E0D523065874EF6FF0C8BF768C067E58DEF5F84540E91CD8BD53002"), vbExclamation
        Case 70 ' Permission denied
            MsgBox PI_Zh("6CA167098BBF95EE67439650FF0C8BF768C067E565874EF6674396503002"), vbExclamation
        Case Else
            MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl
    End Select
    sResult = ""
    Resume proc_exit
    
End Function

Sub test1()
 #If Mac Then
 Debug.Print "MAC"
 #Else
 Debug.Print "WIN"
 #End If
End Sub
