Attribute VB_Name = "Module3"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type


Public Pau As Boolean, Mouse As POINTAPI, lngHDC  As Long, AB(8) As Long
Public N2(70) As Double, RR1(70) As Double
Public N3(70) As Double, RR2(70) As Double
Public N4(70) As Double, RR3(70) As Double
Public X As Long, Y As Long, XX As Long, YY As Long, I5 As Integer
Public I As Double, I1 As Integer, I2 As Double, I3 As Double, N1 As Double, RR0 As Double
Public A1 As Double, A2 As Double, A3 As Double, A4 As Double, AA1 As Double, AA2 As Double, AA3 As Double, AA4 As Double, A As Double


Public Function CIS(Num) As Double ' Math Combinations for the Pictures
Dim Rev As Double
On Error Resume Next
Rev = Sin(Num) * Cos(Num) * Sin(Cos(Num)) * Cos(Sin(Num)) ' -1=<Rev<=1
CIS = Atn(Rev) * Rev '>=1 or <=-1
If CIS = 0 Then Dec
End Function

Public Function SZ(Num) As Double  'Math Combinations for the Pictures
Dim Rev As Double
On Error Resume Next
Rev = Sin(Num) * Cos(Num) * Sin(Cos(Num)) * Cos(Sin(Num)) * Tan(Num) '- 1 >= Rev <= 1
SZ = Atn(Rev) * Log(Abs(Num) + 1) * Sqr(Abs(Num)) * Exp(Rev) * Rev '>=1 or <=-1
If SZ = 0 Then Dec
End Function

Public Function QW(Num) As Double 'Math Combinations for the Pictures
Dim Rev As Double
On Error Resume Next
Rev = Sin(Num) * Cos(Num) * Sin(Cos(Num)) * Cos(Sin(Num)) ' -1=<Rev<=1
QW = Atn(Rev) * Tan(Rev) * Log(Abs(Rev) + 1) * Sqr(Abs(Rev)) * Exp(Rev) * Sqr(Abs(Num)) * Log(Abs(Num) + 1) '>=1 or <=-1
If QW = 0 Then Dec
End Function


Public Sub SLEEP1() ' Short Pause function
Dim G
For G = 1 To 10 ^ 5.9
DoEvents
Next G
End Sub

Public Function RNDD(Num2 As Double, Dot As Boolean) As Single
'Function makes a random number greater then 1 and less then -1
Randomize Timer
Randomize Rnd
Dim Rndd1 As Integer
24 RNDD = 1 + (Rnd * Num2)
If Dot = False Then RNDD = Int(RNDD)
If Int(RNDD) = 0 Or RNDD > Num2 Then GoTo 24
End Function

Public Function Sgn1(ParamArray Num2() As Variant) As Variant
Dim Rev
'The function gets sign of Number by Multiplicating the signs of the wanted numbers
Sgn1 = 1
For Rev = LBound(Num2) To UBound(Num2)
If Sgn(Num2(Rev)) <> 0 Then Sgn1 = Sgn1 * Sgn(Num2(Rev))
Next
End Function

Public Function Sgn2(Num As Double, WhatFunc As Integer) As Double
If WhatFunc = 1 Then Sgn2 = Num
If WhatFunc = 2 Then Sgn2 = -Num
If WhatFunc = 3 Then Sgn2 = Not Num
If WhatFunc = 4 Then Sgn2 = Not -Num
If WhatFunc = 5 Then Sgn2 = -Not Num
If WhatFunc = 6 Then Sgn2 = Abs(Num)
End Function

Public Sub Dec()
'The sub that calles the drawing sub
Dim I3
I5 = RNDD(5, False)
I3 = RNDD(2, False)
If I3 = 1 Then DD1
If I3 = 2 Then DD2
End Sub




