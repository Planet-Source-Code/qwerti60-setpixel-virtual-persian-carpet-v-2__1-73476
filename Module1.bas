Attribute VB_Name = "Module1"
Public Cmd As String, lp As RECT, hWnd1

Private hThumb As Long, rctThumb As RECT, lStyle As Long
Private Const WS_CHILD = &H40000000
Private Const GWL_STYLE = (-16)
Private Const GWL_HWNDPARENT = (-8)
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1

Sub Main()
HVNum

Cmd$ = Left$(UCase$(Trim(Command$)), 2)
Select Case Cmd$
Case "", "/S"
ShowCursor False
Load Form3

Case "/P"
Form3.Timer1.Interval = 1250
Form3.Timer2.Interval = 1
Load Form3

Case "/C"
    ShellAbout GetHwnd(Cmd), "Roman Braverman's - Screen Saver", "Beautiful Picture, Enjoy", 1
End
End Select
End Sub

Public Function GetHwnd(Cmd As String) As Long
'-----------------------------------------------------------------
    Dim Str As String                           ' substring variable
    Dim lenStr As Long                          ' length of substring
    Dim Idx As Long                             ' Index variable
'-----------------------------------------------------------------
    Str = Trim$(Cmd)                            ' copy command line
    lenStr = Len(Str)                           ' get size of string
    
    For Idx = lenStr To 1 Step -1               ' for each char in string
        Str = Right$(Str, Idx)                  ' chop off the rightmost char
        If IsNumeric(Str) Then                  ' if substring is numeric then value is an hWnd
            GetHwnd = Val(Str)           ' return hWnd value
            Exit For                            ' exit for loop
        End If
    Next
'-----------------------------------------------------------------
End Function

Private Sub HVNum()
'for sub D2 , D6 , D8 , D10
RR1(1) = 63496250857.9481
RR1(2) = 63509108395.5719: N2(2) = 928655968002.868
RR1(3) = 63503995048.3465: N2(3) = 94906265.4346
RR1(4) = 63495547759.8836: N2(4) = 675.00009
RR1(5) = 63496611060.4421: N2(5) = 127.000000013
RR1(6) = 65853092877.1852: N2(6) = 801.05003
RR1(7) = 63701475402.3922: N2(7) = 60000000.6014

'for sub D1 , D5 , D7 , D9
RR1(8) = 563056.9133
RR1(9) = 563124.2847: N2(9) = 120.70006
RR1(10) = 563198.4226: N2(10) = 89999999.9999999
RR1(11) = 563034.6488: N2(11) = 9.24999E+15
RR1(12) = 563132.7: N2(12) = 3E+16
RR1(13) = 563075.5941: N2(13) = 3037000493.90842
RR1(14) = 563105.73905: N2(14) = 3036999981.9083

'for sub D3
RR2(1) = 563072.1938
RR2(2) = 563110.434: N3(2) = 119.293
RR2(3) = 563218.3508: N3(3) = 119.293
RR2(4) = 563124.52397: N3(4) = 1102.00097
RR2(5) = 563254.34024: N3(5) = 1.01E+16
RR2(6) = 563084.52701: N3(6) = 3037000493.90842
RR2(7) = 563072.73993: N3(7) = 3000003579.09526

'for sub D4
RR2(8) = 63501864768.9999
RR2(9) = 63520840435.1382: N3(9) = 918046539776.001
RR2(10) = 63484372339.4724: N3(10) = 94905936.5151
RR2(11) = 63566023039.6901: N3(11) = 9222012.00195
RR2(12) = 63544459113.7958: N3(12) = 127.1445
RR2(13) = 63523111912.4249: N3(13) = 3120000.3202
RR2(14) = 63489083320.2793: N3(14) = 211111140.88504
End Sub

Public Sub Preview1()
    hThumb = CLng(Right(Command, Len(Command) - 2))
    Thumbnail = True
    GetClientRect hThumb, rctThumb
    lStyle = GetWindowLong(Form3.hwnd, GWL_STYLE)
    lStyle = lStyle Or WS_CHILD
    SetWindowLong Form3.hwnd, GWL_STYLE, lStyle
    SetParent Form3.hwnd, hThumb
    SetWindowLong Form3.hwnd, GWL_HWNDPARENT, hThumb
    SetWindowPos Form3.hwnd, HWND_TOP, 0, 0, rctThumb.Right, rctThumb.Bottom, SWP_SHOWWINDOW

End Sub

