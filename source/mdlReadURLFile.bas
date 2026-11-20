Attribute VB_Name = "mdlReadURLFile"
'---------------------------------------------------------------------------------------
' Procedure : ReadURLFile
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Access and read a file on a webserver
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFullURLWFile :   Full URL and Filename with the extension
'
' Usage:
' ~~~~~~
' ReadURLFile("http://www.google.ca/index.html")
' ReadURLFile("http://www.SomeDomain.com/SomeFolder/SomeFIle.txt")
' ReadURLFile("http://www.SomeDomain.com/SomeFolder/SomeFIle.dat")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2012-Jul-07                 Initial Release
'---------------------------------------------------------------------------------------
Function ReadURLFile(sFullURLWFile As String) As String
          'Returns an "" (VbNullString) if an error is reported back from the server
          '        otherwise returns the Response Text from the server
          'References for further development
          '        http://msdn.microsoft.com/en-us/library/ms766431%28v=vs.85%29.aspx
          '        http://msdn.microsoft.com/en-us/library/ms763809%28v=vs.85%29.aspx
          '        http://en.wikipedia.org/wiki/List_of_HTTP_status_codes
       
10        On Error GoTo Error_Handler
          Dim oHttp           As Object
       
20        Set oHttp = CreateObject("MSXML2.ServerXMLHTTP")
       
30        Call oHttp.Open("GET", sFullURLWFile, False)
40        Call oHttp.Send
          'Check for any errors reported by the server
50        If oHttp.Status >= 400 And oHttp.Status <= 599 Then
60            ReadURLFile = ""
70            GoTo Error_Handler
80        Else
90            ReadURLFile = oHttp.ResponseText
100       End If
       
Error_Handler_Exit:
110       On Error Resume Next
120       Call oHttp.Close
130       Set oHttp = Nothing
140       Exit Function
       
Error_Handler:
150       If Err = -2147012889 Then
160           MsgBox PI_Zh("65E06CD58FDE63A57F517EDCFF0C8BF768C067E57F517EDC8FDE63A553CA9632706B58998BBE7F6E3002"), vbCritical, PI_Zh("8FDE63A559318D25")
170           Resume Error_Handler_Exit
180       End If
190       If oHttp.Status >= 400 And oHttp.Status <= 599 Then
200           MsgBox PI_Zh("53D1751F4EE54E0B95198BEFFF1A") & vbCrLf & vbCrLf & _
                  PI_Zh("95198BEF7F1653F7FF1A") & oHttp.Status & vbCrLf & _
                  PI_Zh("95198BEF4F4D7F6EFF1A005200650061006400550052004C00460069006C0065") & vbCrLf & _
                  PI_Zh("95198BEF8BF4660EFF1A") & oHttp.StatusText, _
                  vbCritical, PI_Zh("64CD4F5C59318D25")
210       Else
220           MsgBox PI_Zh("53D1751F4EE54E0B95198BEFFF1A") & vbCrLf & vbCrLf & _
                  PI_Zh("95198BEF7F1653F7FF1A") & Err.Number & vbCrLf & _
                  PI_Zh("95198BEF4F4D7F6EFF1A005200650061006400550052004C00460069006C0065") & vbCrLf & _
                  PI_Zh("95198BEF8BF4660EFF1A") & Err.description, _
                  vbCritical, PI_Zh("64CD4F5C59318D25")
230       End If
240       Resume Error_Handler_Exit
End Function

