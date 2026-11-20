Attribute VB_Name = "mdlClrHandling"
' ------------------------------------------------------
' Name: mdlClrHandling
' Kind: Modul
' Purpose: collective module for all color handling routines
' Author: Olaf Noehring ( https://datenbank-projekt.de ) and in a large part https://stackoverflow.com/a/26685704
' Date: 2017-06-21
' ------------------------------------------------------
Option Explicit


'--START Palette handling----------------------------------------------------------------------------------------------------
'Source: https://stackoverflow.com/a/26685704
Public Type HSL
    h As Double ' Range 0 - 1
    s As Double ' Range 0 - 1
    L As Double ' Range 0 - 1
End Type

Public Type RGB
    R As Byte
    G As Byte
    B As Byte
End Type

' ----------------------------------------------------------------
' Procedure Name: ApplyTintAndShade
' Purpose: Apply tint and shade to a color
' Procedure Kind: Function
' Procedure Access: Public
' Parameter hc (HSL):
' Parameter TintAndShade (Double):
' Return Type: HSL
' Author: https://stackoverflow.com/a/26685704
' Date: 2017-06-21
' ----------------------------------------------------------------
Public Function ApplyTintAndShade(hc As HSL, TintAndShade As Double) As HSL
10        If gblbolDebug = True Then LogInfo "Fct ApplyTintAndShade (hc As HSL, TintAndShade As Double)"         'Debug: Write active procedure to log
          
20        If TintAndShade > 0 Then
30            hc.L = hc.L + (1 - hc.L) * TintAndShade
40        Else
50            hc.L = hc.L + hc.L * TintAndShade
60        End If

70        ApplyTintAndShade = hc

End Function

' ----------------------------------------------------------------
' Procedure Name: CalcColor
' Purpose: Calcularte Colorvariations
' Procedure Kind: Sub
' Procedure Access: Public
' Author: https://stackoverflow.com/a/26685704
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub CalcColor()
10        If gblbolDebug = True Then LogInfo "Sub CalcColor"         'Debug: Write active procedure to log
          
          Dim intPosition As Integer
          Dim ii As Integer, jj As Integer
          Dim Pres As Presentation
          Dim schemeColors As ThemeColorScheme
          Dim ts As Double
          Dim c, c2 As Long
          Dim hc As HSL, hc2 As HSL

20        Set Pres = ActivePresentation
30        Set schemeColors = Pres.Designs(1).SlideMaster.Theme.ThemeColorScheme

40        intPosition = 0

          ' For all colors
50        For ii = 0 To 11
60            c = schemeColors(ii + 1).RGB

              ' Generate all the color variations
70            For jj = 0 To 5
80                hc = RGBtoHSL(c)
90                ts = SelectTintOrShade(hc, jj)
100               hc2 = ApplyTintAndShade(hc, ts)
110               c2 = HSLtoRGB(hc2)
                  'not needed anymore: create color swatch on active slide
                  '        Call CreateShape(pres.Slides(1), ii, jj, c2)
                  'new: colorize button on form
120               procFormatDirect intPosition, c2
130               intPosition = intPosition + 1

              
140           Next jj
            
150       Next ii

End Sub

' ----------------------------------------------------------------
' Procedure Name: ColorToRGB
' Purpose:
' Procedure Kind: Function
' Procedure Access: Public
' Parameter c (ColorType):
' Return Type: Long
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function ColorToRGB(c As ColorType) As Long
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct ColorToRGB (c As ColorType)"         'Debug: Write active procedure to log
              
30        ColorToRGB = RGB(c.red, c.green, c.blue)
          
exit_proc:
40        Exit Function
          
ErrorHandler:
50        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
60        Select Case Err.Number
              Case -2147188160      'no active presentation
70                MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "ColorToRGB: Nothing to do!"
80                Resume exit_proc
90                Resume
100           Case Else
110               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "ColorToRGB"
120               Resume Next
130               Resume
              
140       End Select
              
End Function

' ----------------------------------------------------------------
' Procedure Name: Color_to_RGB
' Purpose: returns R G B values for a given color
' Procedure Kind: Function
' Procedure Access: Public
' Parameter Color (Long): color in long format
' Parameter strColorReturn (String): color you want returned (r,g,b or nothing for a combination)
' Return Type: String
' Author:  http://www.garybeene.com/code/visual%20basic168.htm
' Date: 2017-06-21
' ----------------------------------------------------------------
Function Color_to_RGB(Color As Long, Optional strColorReturn As String) As String
          'Function Color_to_RGB(Color As Long, R As Integer, G As Integer, B As Integer) As Long
          'source: http://www.garybeene.com/code/visual%20basic168.htm
          '  R = Color And &HFF&
          '  G = (Color And &HFF00&) \ &H100&
          '  B =   (Color &    And &HFF0000) \ &H10000
          '    'or
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct Color_to_RGB: " & Color & " / Optional: " & strColorReturn        'Debug: Write active procedure to log
              
          Dim R As Integer
          Dim G As Integer
          Dim B As Integer
          Dim strReturn As String
          
30        R = Color Mod 256
40        G = (Color \ 256) Mod 256
50        B = (Color \ 256 \ 256) Mod 256
          
60        Select Case strColorReturn
              Case "R"
70                strReturn = R

80            Case "G"
90                strReturn = G

100           Case "B"
110               strReturn = B

120           Case Else
130               strReturn = R & ";" & G & ";" & B

140       End Select

150       Color_to_RGB = strReturn
          
exit_proc:
160       Exit Function
          
ErrorHandler:
170       If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
180       Select Case Err.Number
              Case -2147188160      'no active presentation
190               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "Color_to_RGB: Nothing to do!"
200               Resume exit_proc
210               Resume
220           Case Else
230               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "Color_to_RGB"
240               Resume Next
250               Resume
              
260       End Select
End Function

' ----------------------------------------------------------------
' Procedure Name: CreateShape
' Purpose: Add a shape to a slide. This demos a colorswatch on the slide
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter slide (slide):
' Parameter xIndex (Integer):
' Parameter yIndex (Integer):
' Parameter Color (Long):
' Author: https://stackoverflow.com/a/26685704
' Date: 2017-06-21
' ----------------------------------------------------------------
Sub CreateShape(slide As slide, xIndex As Integer, yIndex As Integer, Color As Long)
10        If gblbolDebug = True Then LogInfo "Sub CreateShape: slide / " & xIndex & " / " & yIndex & " / " & Color       'Debug: Write active procedure to log
          
          'This created a color swatch on the active slide - which we do not need anymore
          Dim newShape As Shape
          Dim xStart As Integer, yStart As Integer
          Dim xOffset As Integer, yOffset As Integer
          Dim xSize As Integer, ySize As Integer
20        xStart = 100
30        yStart = 100
40        xOffset = 30
50        yOffset = 30
60        xSize = 25
70        ySize = 25

80        Set newShape = slide.Shapes.AddShape(msoShapeRectangle, xStart + xOffset * xIndex, yStart + yOffset * yIndex, xSize, ySize)
90        newShape.Fill.BackColor.RGB = Color
100       newShape.Fill.ForeColor.RGB = Color
110       newShape.Line.ForeColor.RGB = 0
120       newShape.Line.BackColor.RGB = 0

End Sub

' ----------------------------------------------------------------
' Procedure Name: DeserializeColor
' Purpose:
' Procedure Kind: Function
' Procedure Access: Public
' Parameter s (String):
' Return Type: ColorType
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function DeserializeColor(s As String) As ColorType
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct DeserializeColor: " & s          'Debug: Write active procedure to log
              
          Dim words() As String
30        words() = Split(s, ";")
40        DeserializeColor.red = CInt(words(0))
50        DeserializeColor.green = CInt(words(1))
60        DeserializeColor.blue = CInt(words(2))
          
exit_proc:
70        Exit Function
          
ErrorHandler:
80        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
90        Select Case Err.Number
              Case -2147188160      'no active presentation
100               MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "DeserializeColor: Nothing to do!"
110               Resume exit_proc
120               Resume
130           Case Else
140               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "DeserializeColor"
150               Resume Next
160               Resume
              
170       End Select
              
End Function

' ----------------------------------------------------------------
' Procedure Name: H2C
' Purpose: Hue to RGB calculation
' Procedure Kind: Function
' Procedure Access: Public
' Parameter x (Double):
' Parameter Y (Double):
' Parameter hc (Double):
' Return Type: Double
' Author: https://stackoverflow.com/a/26685704
' Date: 2017-06-21
' ----------------------------------------------------------------
Function H2C(x As Double, Y As Double, hc As Double) As Double
10        If gblbolDebug = True Then LogInfo "Fct H2C: " & x & " / " & Y & " / " & hc         'Debug: Write active procedure to log
          
20        Select Case hc
              Case Is < 1 / 6: H2C = Y + ((x - Y) * 6 * hc)
30            Case Is < 1 / 2: H2C = x
40            Case Is < 2 / 3: H2C = Y + ((x - Y) * ((2 / 3) - hc) * 6)
50            Case Else:       H2C = Y
60        End Select

End Function
'--END Palette handling----------------------------------------------------------------------------------------------------

' ----------------------------------------------------------------
' Procedure Name: HSLtoRGB
' Purpose: HSL to RGB calculation
' Procedure Kind: Function
' Procedure Access: Public
' Parameter HSL (HSL):
' Return Type: Long
' Author: https://stackoverflow.com/a/26685704
' Date: 2017-06-21
' ----------------------------------------------------------------
Function HSLtoRGB(ByRef HSL As HSL) As Long
10        If gblbolDebug = True Then LogInfo "Fct HSLtoRGB (ByRef HSL As HSL)"         'Debug: Write active procedure to log
          
          Dim R As Double
          Dim G As Double
          Dim B As Double

          Dim x As Double
          Dim Y As Double

20        With HSL

30            If .s = 0 Then

40                R = .L
50                G = .L
60                B = .L

70            Else

80                Select Case .L
                      Case Is < 0.5: x = .L * (1 + .s)
90                    Case Else:     x = .L + .s - (.L * .s)
100               End Select

110               Y = 2 * .L - x

120               R = H2C(x, Y, IIf(.h > 2 / 3, .h - 2 / 3, .h + 1 / 3))
130               G = H2C(x, Y, .h)
140               B = H2C(x, Y, IIf(.h < 1 / 3, .h + 2 / 3, .h - 1 / 3))

150           End If

160       End With

170       HSLtoRGB = CLng("&H00" & _
              Right$("0" & Hex$(Round(B * 255)), 2) & _
              Right$("0" & Hex$(Round(G * 255)), 2) & _
              Right$("0" & Hex$(Round(R * 255)), 2))

180       If gblbolDebug = True Then LogInfo "R: " & Round(B * 255) & ", G: " & Round(G * 255) & ", B: " & Round(R * 255)

End Function

' From RGB to HSL

' ----------------------------------------------------------------
' Procedure Name: RGBtoHSL
' Purpose: RGB to HSL calculation
' Procedure Kind: Function
' Procedure Access: Public
' Parameter RGB (Long):
' Return Type: HSL
' Author: https://stackoverflow.com/a/26685704
' Date: 2017-06-21
' ----------------------------------------------------------------
Function RGBtoHSL(ByVal RGB As Long) As HSL
10        If gblbolDebug = True Then LogInfo "Fct RGBtoHSL (ByVal RGB As Long)"         'Debug: Write active procedure to log
          
          Dim R As Double ' Range 0 - 1
          Dim G As Double ' Range 0 - 1
          Dim B As Double ' Range 0 - 1

          Dim RGB_Max  As Double
          Dim RGB_Min  As Double
          Dim RGB_Diff As Double

          Dim HexString As String

20        HexString = Right$(String$(7, "0") & Hex$(RGB), 8)
30        R = CDbl("&H" & Mid$(HexString, 7, 2)) / 255
40        G = CDbl("&H" & Mid$(HexString, 5, 2)) / 255
50        B = CDbl("&H" & Mid$(HexString, 3, 2)) / 255

60        RGB_Max = R
70        If G > RGB_Max Then RGB_Max = G
80        If B > RGB_Max Then RGB_Max = B

90        RGB_Min = R
100       If G < RGB_Min Then RGB_Min = G
110       If B < RGB_Min Then RGB_Min = B

120       RGB_Diff = RGB_Max - RGB_Min

130       With RGBtoHSL

140           .L = (RGB_Max + RGB_Min) / 2

150           If RGB_Diff = 0 Then

160               .s = 0
170               .h = 0

180           Else

190               Select Case RGB_Max
                      Case R: .h = (1 / 6) * (G - B) / RGB_Diff - (B > G)
200                   Case G: .h = (1 / 6) * (B - R) / RGB_Diff + (1 / 3)
210                   Case B: .h = (1 / 6) * (R - G) / RGB_Diff + (2 / 3)
220               End Select

230               Select Case .L
                      Case Is < 0.5: .s = RGB_Diff / (2 * .L)
240                   Case Else:     .s = RGB_Diff / (2 - (2 * .L))
250               End Select

260           End If

270       End With

End Function

' The tint and shade value is a value between -1.0 and 1.0, where
' -1.0 means fully shading (black), and 1.0 means fully tinting (white)
' A tint/shade value of 0.0 will not change the color
'Source https://stackoverflow.com/a/26685704
Public Function SelectTintOrShade(hc As HSL, variationIndex As Integer) As Double
10        If gblbolDebug = True Then LogInfo "Pub Fct SelectTintOrShade (hc As HSL, variationIndex As Integer)"         'Debug: Write active procedure to log
          
          Dim shades(5) As Variant
20        shades(0) = Array(0#, 0.5, 0.35, 0.25, 0.15, 0.05)
30        shades(1) = Array(0#, 0.9, 0.75, 0.5, 0.25, 0.1)
40        shades(2) = Array(0#, 0.8, 0.6, 0.4, -0.25, -0.5)
50        shades(3) = Array(0#, -0.1, -0.25, -0.5, -0.75, -0.9)
60        shades(4) = Array(0#, -0.05, -0.15, -0.25, -0.35, -0.5)

70        Select Case hc.L
              Case Is < 0.001: SelectTintOrShade = shades(0)(variationIndex)
80            Case Is < 0.2:   SelectTintOrShade = shades(1)(variationIndex)
90            Case Is < 0.8:   SelectTintOrShade = shades(2)(variationIndex)
100           Case Is < 0.999: SelectTintOrShade = shades(3)(variationIndex)
110           Case Else:       SelectTintOrShade = shades(4)(variationIndex)
120       End Select
End Function

' ----------------------------------------------------------------
' Procedure Name: SerializeColor
' Purpose: Create a ; seperated string from RGB values
' Procedure Kind: Function
' Procedure Access: Public
' Parameter c (ColorType):
' Return Type: String
' Author: Olaf Noehring ( https://datenbank-projekt.de )
' Date: 2017-06-21
' ----------------------------------------------------------------
Function SerializeColor(c As ColorType) As String
10        On Error GoTo ErrorHandler
20        If gblbolDebug = True Then LogInfo "Fct SerializeColor (c As ColorType)"         'Debug: Write active procedure to log
          
30        SerializeColor = CStr(c.red) & ";" & CStr(c.green) & ";" & CStr(c.blue)
          
exit_proc:
40        Exit Function
          
ErrorHandler:
50        If gblbolDebug = True Then LogInfo "L: " & Erl & "; # " & Err.Number & "; " & Err.description     'Debug: Write active procedure to log
60        Select Case Err.Number
              Case -2147188160      'no active presentation
70                MsgBox PI_Zh("8BF7514862535F004E004E2A6F14793A65877A3F3002") & PI_Zh("4EE37801884CFF1A") & Erl, , "SerializeColor: Nothing to do!"
80                Resume exit_proc
90                Resume
100           Case Else
110               MsgBox Err.Number & vbCrLf & Err.description & vbCrLf & PI_Zh("4EE37801884CFF1A") & Erl, , "SerializeColor"
120               Resume Next
130               Resume
              
140       End Select
          
              
End Function

