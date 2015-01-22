VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse clicker"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Mouse button click"
      Height          =   1575
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   3615
      Begin VB.OptionButton Option6 
         Caption         =   "Middle"
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Right"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Left"
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   3240
      Top             =   720
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   3615
      Begin VB.OptionButton Option3 
         Caption         =   "Num Lock"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Scroll Lock"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Caps Lock"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Text            =   "1000"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   996
      Left            =   3240
      Top             =   240
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Fact delay, ms, before click"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Delay, ms, before click"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Please, select lock, which activates mouse button clicker:"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'In this project was used part of code of
'KPD-Team 1998
'URL: http://www.allapi.net/
'E-Mail: KPDTeam@Allapi.net
    
'Before you start this program, I suggest you save everything that wasn't saved yet.
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4
Const MOUSEEVENTF_MIDDLEDOWN = &H20
Const MOUSEEVENTF_MIDDLEUP = &H40
Const MOUSEEVENTF_MOVE = &H1
Const MOUSEEVENTF_ABSOLUTE = &H8000
Const MOUSEEVENTF_RIGHTDOWN = &H8
Const MOUSEEVENTF_RIGHTUP = &H10
Dim rnd100 As Integer
Dim n As Integer 'n is time multiplicator
Dim timer2_counter 'which compares with n as cycle counter
Dim t2 As Boolean



Private Sub Form_Load()
Timer1.Interval = Text1.Text
Option4.Value = True
t2 = False
End Sub


Private Sub Text1_Change()
If Text1.Text <= 65000 Then
Timer1.Interval = Text1.Text
End If
If Text1.Text > 65000 Then
    Timer2.Interval = 65000
    n = Text1.Text \ 65000
    Timer1.Interval = Text1.Text - n * 65000
    Label2.Caption = Text1.Text - n * 65000
    Label5.Caption = Text1.Text - n * 65000
    Timer1.Enabled = False
    Timer2.Enabled = True
    If Timer1.Interval = 0 Then
    Time1.Interval = 1000
    End If
End If
End Sub

Private Sub Timer1_Timer()


If Option1.Value = True Then
    If GetKeyState(vbKeyCapital) = 1 Then
        If Option4.Value = True Then
        mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0&, 0&, cButt, dwEI
        End If
        If Option5.Value = True Then
        mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0&, 0&, cButt, dwEI
        End If
        If Option6.Value = True Then
        mouse_event MOUSEEVENTF_MIDDLEDOWN Or MOUSEEVENTF_MIDDLEUP, 0&, 0&, cButt, dwEI
        End If
        
        Randomize
        rnd100 = Rnd(10) * 100 Mod 100
        If rmd100 > 50 Then
            If Timer1.Interval + rnd100 / 4 < 65535 Then
            Timer1.Interval = Timer1.Interval + rnd100 / 4
            End If
        Else
        Timer1.Interval = Timer1.Interval - rnd100 / 4
        End If
        
        If Text1.Text <= 65000 Then
            If Timer1.Interval / Text1.Text > 1.1 Then
            Timer1.Interval = Text1.Text
            End If
            If Timer1.Interval / Text1.Text < 0.9 Then
            Timer1.Interval = Text1.Text
            End If
        Else
            If Timer1.Interval / Label5.Caption > 1.1 Then
            Timer1.Interval = Label5.Caption
            End If
            If Timer1.Interval / Label5.Caption < 0.9 Then
            Timer1.Interval = Label5.Caption
            End If
        End If
    End If
End If
If Option2.Value = True Then
    If GetKeyState(vbKeyScrollLock) = 1 Then
        If Option4.Value = True Then
        mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0&, 0&, cButt, dwEI
        End If
        If Option5.Value = True Then
        mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0&, 0&, cButt, dwEI
        End If
        If Option6.Value = True Then
        mouse_event MOUSEEVENTF_MIDDLEDOWN Or MOUSEEVENTF_MIDDLEUP, 0&, 0&, cButt, dwEI
        End If
        
        Randomize
        rnd100 = Rnd(10) * 100 Mod 100
        If rmd100 > 50 Then
            If Timer1.Interval + rnd100 / 4 < 65535 Then
            Timer1.Interval = Timer1.Interval + rnd100 / 4
            End If
        Else
        Timer1.Interval = Timer1.Interval - rnd100 / 4
        End If
        
        If Text1.Text <= 65000 Then
            If Timer1.Interval / Text1.Text > 1.1 Then
            Timer1.Interval = Text1.Text
            End If
            If Timer1.Interval / Text1.Text < 0.9 Then
            Timer1.Interval = Text1.Text
            End If
        Else
            If Timer1.Interval / Label5.Caption > 1.1 Then
            Timer1.Interval = Label5.Caption
            End If
            If Timer1.Interval / Label5.Caption < 0.9 Then
            Timer1.Interval = Label5.Caption
            End If
        End If
    End If
End If
If Option3.Value = True Then
    If GetKeyState(vbKeyNumlock) = 1 Then
        If Option4.Value = True Then
        mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0&, 0&, cButt, dwEI
        End If
        If Option5.Value = True Then
        mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0&, 0&, cButt, dwEI
        End If
        If Option6.Value = True Then
        mouse_event MOUSEEVENTF_MIDDLEDOWN Or MOUSEEVENTF_MIDDLEUP, 0&, 0&, cButt, dwEI
        End If
        
        Randomize
        rnd100 = Rnd(10) * 100 Mod 100
        If rmd100 > 50 Then
            If Timer1.Interval + rnd100 / 4 < 65535 Then
            Timer1.Interval = Timer1.Interval + rnd100 / 4
            End If
        Else
        Timer1.Interval = Timer1.Interval - rnd100 / 4
        End If
        
        If Text1.Text <= 65000 Then
            If Timer1.Interval / Text1.Text > 1.1 Then
            Timer1.Interval = Text1.Text
            End If
            If Timer1.Interval / Text1.Text < 0.9 Then
            Timer1.Interval = Text1.Text
            End If
        Else
            If Timer1.Interval / Label5.Caption > 1.1 Then
            Timer1.Interval = Label5.Caption
            End If
            If Timer1.Interval / Label5.Caption < 0.9 Then
            Timer1.Interval = Label5.Caption
            End If
        End If
    End If
End If
Label2.Caption = Timer1.Interval

If t2 = True Then
    t2 = False
    Timer2.Enabled = True
    Timer1.Enabled = False
    End If
End Sub

Private Sub Timer2_Timer()
timer2_counter = timer2_counter + 1

If timer2_counter = n Then
    timer2_counter = 0
    Timer1.Enabled = True
    t2 = True
    Timer2.Enabled = False
End If
End Sub
