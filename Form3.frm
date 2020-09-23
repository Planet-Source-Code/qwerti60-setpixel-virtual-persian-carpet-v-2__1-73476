VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   4665
   ClientTop       =   3645
   ClientWidth     =   5040
   DrawWidth       =   10
   FillColor       =   &H00004000&
   ForeColor       =   &H00004040&
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form3.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Left            =   1680
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Left            =   2160
      Top             =   1080
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   2280
      Top             =   2400
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rev1


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 32 And KeyCode <> Asc("P") And KeyCode <> Asc("p") And Cmd <> "/P" And Pau = False And KeyCode <> 113 Then Form_Unload (0)
If KeyCode = 32 And Pau <> True Then Dec 'Press spacebar to start a new drawing
If KeyCode = Asc("P") Or KeyCode = Asc("p") Then Pau = Not Pau  'P/p to pause/play the drawing
'If KeyCode = 113 Then SavePicture Form3.Image, App.Path + "\Untitled.bmp"
End Sub


Private Sub Form_Load()
Dim hdc1
GetCursorPos Mouse 'for detecting mouse's motion
Form3.AutoRedraw = True
hdc1 = GetWindowDC(GetDesktopWindow)
BitBlt Form3.hdc, 0, 0, Form3.Width, Form3.Height, hdc1, 0, 0, &HCC0020
'copies the screen to the form
Form3.AutoRedraw = False

If Cmd$ <> "/P" Then Me.WindowState = 2
If Cmd$ = "/P" Then Preview1

BringWindowToTop Me.hwnd ' for the Screen Saver
Me.Show
Pau = False
Dec
End Sub

Private Sub Form_LostFocus()
Form_Unload (0)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Cmd <> "/P" And Pau <> True Then Form_Unload (0) 'when the mouse is pressed the programme ends
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (((Mouse.X - X) <> 0) Or ((Mouse.Y - Y) <> 0)) And Cmd <> "/P" And Pau <> True Then Form_Unload (0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowCursor True
Unload Form3
End
End Sub


Private Sub Timer1_Timer()
Dec
End Sub

Private Sub Timer2_Timer()
Dim G
For G = 1 To (10 ^ 5.9)
Next G
End Sub

Private Sub Timer3_Timer()
If Cmd = "/P" And GetHwnd(VBA.Command$) = 0 Then Form_Unload (0) 'a loop to pause the drawing
BringWindowToTop Me.hwnd ' for the Screen Saver
DoEvents
If Pau = True Then
Do
DoEvents
Loop Until Pau = False
End If
End Sub
