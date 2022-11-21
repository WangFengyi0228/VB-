VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "弹球游戏"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12990
   FillColor       =   &H008080FF&
   ForeColor       =   &H008080FF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   12990
   Begin VB.Timer Timer8 
      Left            =   11520
      Top             =   6360
   End
   Begin VB.Timer Timer7 
      Left            =   10920
      Top             =   6360
   End
   Begin VB.Timer Timer6 
      Left            =   10320
      Top             =   6360
   End
   Begin VB.Timer Timer5 
      Left            =   9720
      Top             =   6360
   End
   Begin VB.Timer Timer4 
      Left            =   9120
      Top             =   6360
   End
   Begin VB.Timer Timer3 
      Left            =   8520
      Top             =   6360
   End
   Begin VB.Timer Timer2 
      Left            =   7920
      Top             =   6360
   End
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Left            =   10920
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   7320
      Top             =   6360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      MaskColor       =   &H00FF0000&
      TabIndex        =   2
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "开始"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      MaskColor       =   &H00FF0000&
      TabIndex        =   1
      Top             =   6240
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   6000
      Left            =   0
      ScaleHeight     =   5940
      ScaleMode       =   0  'User
      ScaleWidth      =   10000
      TabIndex        =   0
      Top             =   0
      Width           =   10000
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -120
         ScaleHeight     =   315
         ScaleWidth      =   10035
         TabIndex        =   8
         Top             =   0
         Width           =   10095
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "幼圆"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9480
            TabIndex        =   9
            Text            =   "100"
            Top             =   0
            Width           =   660
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            BorderWidth     =   30
            X1              =   0
            X2              =   9900
            Y1              =   120
            Y2              =   120
         End
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H000000FF&
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   9240
         Shape           =   3  'Circle
         Top             =   -120
         Width           =   735
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H0000C000&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   8040
         Shape           =   3  'Circle
         Top             =   -360
         Width           =   495
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H000000FF&
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   6720
         Shape           =   3  'Circle
         Top             =   -120
         Width           =   495
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H0000C000&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   5640
         Shape           =   3  'Circle
         Top             =   -240
         Width           =   495
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   4440
         Shape           =   3  'Circle
         Top             =   -120
         Width           =   375
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FF0000&
         BorderColor     =   &H0000C000&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   3120
         Shape           =   3  'Circle
         Top             =   -360
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   -240
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000D&
         BorderWidth     =   8
         X1              =   1930.618
         X2              =   4343.892
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         BorderColor     =   &H0000C000&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   480
         Shape           =   3  'Circle
         Top             =   -360
         Width           =   495
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   11400
      TabIndex        =   6
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      Caption         =   "Level :"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   11400
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "Score："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpszCommand As String) As Long

Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10
Const SND_NOWAIT = &H2000

Dim gamescore As Integer
Dim blood As Integer
Dim move_x As Integer
Dim level As Integer
Dim dh As Integer
Dim temp As Long
Dim temp1 As Long
Dim midi As Long
Dim flag As Integer
Dim x_step1 As Integer
Dim y_step1 As Integer
Dim x_step2 As Integer
Dim y_step2 As Integer
Dim x_step3 As Integer
Dim y_step3 As Integer
Dim x_step4 As Integer
Dim y_step4 As Integer
Dim x_step5 As Integer
Dim y_step5 As Integer
Dim x_step6 As Integer
Dim y_step6 As Integer
Dim x_step7 As Integer
Dim y_step7 As Integer
Dim x_step8 As Integer
Dim y_step8 As Integer

'界面初始化
Private Sub Form_Load()
    blood = 100
    move_x = 10
    Text1.Text = blood
    Text1.Left = 9480
    Line2.X2 = 9900
    Picture1.Picture = LoadPicture("img/p1.jpg")
    Picture2.Picture = LoadPicture("img/e1.jpg")
    defeat
    Form1.Left = (Screen.Width - Form1.Width) / 2
    Form1.Top = (Screen.Height - Form1.Height) / 2 - 600
    flag = 1
    dh = 50
    
    x_step1 = 30
    y_step1 = 30
    x_step2 = 30
    y_step2 = 30
    x_step3 = 30
    y_step3 = 30
    x_step4 = 30
    y_step4 = 30
    x_step5 = 30
    y_step5 = 30
    x_step6 = 30
    y_step6 = 30
    x_step7 = 30
    y_step7 = 30
    x_step8 = 30
    y_step8 = 30
End Sub

'重来函数
Public Function defeat()
    'timer初始化
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
    Timer6.Enabled = False
    Timer7.Enabled = False
    Timer8.Enabled = False
    Timer1.Interval = 50
    Timer2.Interval = 50
    Timer3.Interval = 50
    Timer4.Interval = 50
    Timer5.Interval = 50
    Timer6.Interval = 50
    Timer7.Interval = 50
    Timer8.Interval = 50
    '球的位置初始化
    Shape1.Top = Picture1.Top - 1000
    Shape2.Top = Picture1.Top - 1000
    Shape3.Top = Picture1.Top - 1000
    Shape4.Top = Picture1.Top - 1000
    Shape5.Top = Picture1.Top - 1000
    Shape6.Top = Picture1.Top - 1000
    Shape7.Top = Picture1.Top - 1000
    Shape8.Top = Picture1.Top - 1000
         
    '小球水平位置的随机
    Shape1.Left = Picture1.Left + (Int(Rnd * (250 - 100 + 1)) + 100)
    Shape2.Left = Picture1.Left + (Int(Rnd * (1500 - 1250 + 1)) + 1250)
    Shape3.Left = Picture1.Left + (Int(Rnd * (2550 - 2400 + 1)) + 2400)
    Shape4.Left = Picture1.Left + (Int(Rnd * (4000 - 3850 + 1)) + 3850)
    Shape5.Left = Picture1.Left + (Int(Rnd * (5250 - 5100 + 1)) + 5100)
    Shape6.Left = Picture1.Left + (Int(Rnd * (6500 - 6350 + 1)) + 6350)
    Shape7.Left = Picture1.Left + (Int(Rnd * (7750 - 7600 + 1)) + 7600)
    Shape8.Left = Picture1.Left + (Int(Rnd * (9000 - 8850 + 1)) + 8850)
                
    '等级初始化
    level = 1
                
    gamescore = 0
    Label4.Caption = level
    Label2.Caption = gamescore
    move_x = 10
    Command1.Enabled = True
    Form1.Line1.X1 = 1920
    Form1.Line1.X2 = 4320
    
    blood = 100
    Text1.Text = blood
    Text1.Left = 9480
    Line2.X2 = 9900
    
    flag = 0
End Function

Public Function level1()
    Picture2.PaintPicture LoadPicture("img/e1.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
    Picture1.Picture = LoadPicture("img/p1.jpg")
    Shape6.Top = Picture1.Top - (Int(Rnd * (1300 - 1200 + 1)) + 1200)
    Shape3.Top = Picture1.Top + (Int(Rnd * (100 - 500 + 1)) + 50)
    Shape4.Top = Picture1.Top + (Int(Rnd * (700 - 600 + 1)) + 600) * 2
    dh = 50
    x_step3 = dh
    y_step3 = dh
    x_step4 = -dh
    y_step4 = dh
    x_step6 = dh
    y_step6 = dh
    Timer3.Enabled = True
    Timer4.Enabled = True
    Timer6.Enabled = True
End Function

Public Function level2()
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer6.Enabled = False
    Shape3.Top = Picture1.Top - 1000
    Shape4.Top = Picture1.Top - 1000
    Shape6.Top = Picture1.Top - 1000
    
    Picture2.PaintPicture LoadPicture("img/e1.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
    Picture1.Picture = LoadPicture("img/p1.jpg")
    If Form2.Option3.Value = True Then
         Line1.X2 = Line1.X2 - 110
    End If
    temp = sndPlaySound("snd/s3.wav", SND_ASYNC)
    If Form2.Option1.Value = True Then
         move_x = move_x + 30
    Else
         move_x = move_x + 1
    End If
    Shape4.Top = Picture1.Top + (Int(Rnd * (1300 - 1200 + 1)) + 1200) * 2
    Shape3.Top = Picture1.Top - (Int(Rnd * (900 - 800)) + 800) * 2
    Shape6.Top = Picture1.Top + (Int(Rnd * (100 + 1))) * 2
    dh = 70
    x_step3 = -dh
    y_step3 = dh
    x_step4 = dh
    y_step4 = dh
    x_step6 = -dh
    y_step6 = dh
    Timer3.Enabled = True
    Timer4.Enabled = True
    Timer6.Enabled = True
End Function

Public Function level3()
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer6.Enabled = False
    Shape3.Top = Picture1.Top - 1000
    Shape4.Top = Picture1.Top - 1000
    Shape6.Top = Picture1.Top - 1000

    If Form2.Option3.Value = True Then
         Line1.X2 = Line1.X2 - 110
    End If
    If Form2.Option1.Value = True Then
         move_x = move_x + 30
    Else
         move_x = move_x + 1
    End If
    temp = sndPlaySound("snd/s2.wav", SND_ASYNC)
    Picture2.PaintPicture LoadPicture("img/e2.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
    Picture1.Picture = LoadPicture("img/p2.jpg")
    '小球竖直位置的随机
    Shape6.Top = Picture1.Top + (Int(Rnd * (100 - 50 + 1)) + 50) * 9
    Shape3.Top = Picture1.Top - (Int(Rnd * (700 - 600 + 1)) + 600)
    Shape5.Top = Picture1.Top - (Int(Rnd * (1800 - 1700 + 1)) + 1700)
    Shape4.Top = Picture1.Top - (Int(Rnd * (1400 - 1300 + 1)) + 1300)
    dh = 70
    x_step3 = dh
    y_step3 = dh
    x_step4 = -dh
    y_step4 = dh
    x_step5 = dh
    y_step5 = dh
    x_step6 = -dh
    y_step6 = dh
    Timer3.Enabled = True
    Timer6.Enabled = True
    Timer4.Enabled = True
    Timer5.Enabled = True
End Function


Public Function level4()
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer6.Enabled = False
    Timer5.Enabled = False
    Shape3.Top = Picture1.Top - 1000
    Shape4.Top = Picture1.Top - 1000
    Shape6.Top = Picture1.Top - 1000
    Shape5.Top = Picture1.Top - 1000

    If Form2.Option3.Value = True Then
         Line1.X2 = Line1.X2 - 110
    End If
    If Form2.Option1.Value = True Then
         move_x = move_x + 30
    Else
         move_x = move_x + 1
    End If
    temp = sndPlaySound("snd/s4.wav", SND_ASYNC)
    Picture2.PaintPicture LoadPicture("img/e2.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
    Picture1.Picture = LoadPicture("img/p2.jpg")
    Shape5.Top = Picture1.Top - (Int(Rnd * (1500 - 1400 + 1)) + 1400) * 2
    Shape4.Top = Picture1.Top - (Int(Rnd * (600 - 500 + 1)) + 500) * 2
    Shape6.Top = Picture1.Top + (Int(Rnd * (100 - 50 + 1)) + 50)
    Shape3.Top = Picture1.Top - (Int(Rnd * (1000 - 900 + 1)) + 900) * 5
    dh = 80
    x_step3 = dh
    y_step3 = dh
    x_step4 = dh
    y_step4 = dh
    x_step6 = -dh
    y_step6 = dh
    x_step5 = -dh
    y_step5 = dh
    Timer3.Enabled = True
    Timer6.Enabled = True
    Timer4.Enabled = True
    Timer5.Enabled = True
End Function

Public Function level5()
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer6.Enabled = False
    Timer5.Enabled = False
    Shape3.Top = Picture1.Top - 1000
    Shape4.Top = Picture1.Top - 1000
    Shape6.Top = Picture1.Top - 1000
    Shape5.Top = Picture1.Top - 1000


    If Form2.Option3.Value = True Then
         Line1.X2 = Line1.X2 - 110
    End If
    If Form2.Option1.Value = True Then
         move_x = move_x + 30
    Else
         move_x = move_x + 1
    End If
    temp = sndPlaySound("snd/s1.wav", SND_ASYNC)
    Picture2.PaintPicture LoadPicture("img/e3.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
    Picture1.Picture = LoadPicture("img/p3.jpg")
    Shape5.Top = Picture1.Top - (Int(Rnd * (1000 - 900 + 1)) + 900) * 6
    Shape3.Top = Picture1.Top + (Int(Rnd * (100 - 50 + 1)) + 50)
    Shape4.Top = Picture1.Top - (Int(Rnd * (200 - 150 + 1)) + 150) * 2
    Shape6.Top = Picture1.Top - (Int(Rnd * (1500 - 1400 + 1)) + 1400)
    Shape2.Top = Picture1.Top - (Int(Rnd * (1800 - 1700 + 1)) + 1700) * 2
    dh = 80
    x_step2 = -dh
    y_step2 = dh
    x_step3 = dh
    y_step3 = dh
    x_step4 = dh
    y_step4 = dh
    x_step6 = -dh
    y_step6 = dh
    x_step5 = -dh
    y_step5 = dh
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer6.Enabled = True
    Timer4.Enabled = True
    Timer5.Enabled = True
End Function

Public Function level6()
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer6.Enabled = False
    Timer5.Enabled = False
    Timer2.Enabled = False
    Shape3.Top = Picture1.Top - 1000
    Shape4.Top = Picture1.Top - 1000
    Shape6.Top = Picture1.Top - 1000
    Shape5.Top = Picture1.Top - 1000
    Shape2.Top = Picture1.Top - 1000

    
    If Form2.Option3.Value = True Then
         Line1.X2 = Line1.X2 - 110
    End If
    If Form2.Option1.Value = True Then
         move_x = move_x + 30
    Else
         move_x = move_x + 1
    End If
    temp = sndPlaySound("snd/s2.wav", SND_ASYNC)
    Picture2.PaintPicture LoadPicture("img/e3.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
    Picture1.Picture = LoadPicture("img/p3.jpg")
    Shape2.Top = Picture1.Top - (Int(Rnd * (1000 - 900 + 1)) + 900) * 5
    Shape3.Top = Picture1.Top + (Int(Rnd * (250 - 100 + 1)) + 100)
    Shape5.Top = Picture1.Top - (Int(Rnd * (800 - 700 + 1)) + 700)
    Shape6.Top = Picture1.Top - (Int(Rnd * (900 - 800 + 1)) + 800) * 2
    Shape4.Top = Picture1.Top - (Int(Rnd * (1600 - 1500 + 1)) + 1500) * 2
    dh = 90
    x_step2 = dh
    y_step2 = dh
    x_step3 = dh
    y_step3 = -dh
    x_step4 = dh
    y_step4 = dh
    x_step6 = -dh
    y_step6 = dh
    x_step5 = dh
    y_step5 = dh
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer6.Enabled = True
    Timer4.Enabled = True
    Timer5.Enabled = True
End Function

Public Function level7()
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer6.Enabled = False
    Timer5.Enabled = False
    Timer2.Enabled = False
    Shape3.Top = Picture1.Top - 1000
    Shape4.Top = Picture1.Top - 1000
    Shape6.Top = Picture1.Top - 1000
    Shape5.Top = Picture1.Top - 1000
    Shape2.Top = Picture1.Top - 1000

    If Form2.Option3.Value = True Then
         Line1.X2 = Line1.X2 - 110
    End If
    If Form2.Option1.Value = True Then
         move_x = move_x + 30
    Else
         move_x = move_x + 1
    End If
    temp = sndPlaySound("snd/s3.wav", SND_ASYNC)
    Picture2.PaintPicture LoadPicture("img/e4.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
    Picture1.Picture = LoadPicture("img/p4.jpg")
    Shape3.Top = Picture1.Top - (Int(Rnd * (600 - 500 + 1)) + 500)
    Shape5.Top = Picture1.Top + (Int(Rnd * (250 - 150 + 1)) + 150) * 3
    Shape2.Top = Picture1.Top - (Int(Rnd * (200 - 100 + 1)) + 100)
    Shape6.Top = Picture1.Top - (Int(Rnd * (1000 - 900 + 1)) + 900)
    Shape4.Top = Picture1.Top - (Int(Rnd * (2000 - 1900 + 1)) + 1900)
    Shape7.Top = Picture1.Top - (Int(Rnd * (1000 - 900 + 1)) + 900) * 4
    dh = 90
    x_step2 = dh
    y_step2 = dh
    x_step3 = -dh
    y_step3 = dh
    x_step4 = dh
    y_step4 = dh
    x_step6 = -dh
    y_step6 = dh
    x_step5 = dh
    y_step5 = dh
    x_step7 = -dh
    y_step7 = dh
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer6.Enabled = True
    Timer4.Enabled = True
    Timer5.Enabled = True
    Timer7.Enabled = True
End Function

Public Function level8()
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer6.Enabled = False
    Timer5.Enabled = False
    Timer2.Enabled = False
    Timer7.Enabled = False
    Shape3.Top = Picture1.Top - 1000
    Shape4.Top = Picture1.Top - 1000
    Shape6.Top = Picture1.Top - 1000
    Shape5.Top = Picture1.Top - 1000
    Shape2.Top = Picture1.Top - 1000
    Shape7.Top = Picture1.Top - 1000

    If Form2.Option3.Value = True Then
         Line1.X2 = Line1.X2 - 110
    End If
    If Form2.Option1.Value = True Then
         move_x = move_x + 30
    Else
         move_x = move_x + 1
    End If
    temp = sndPlaySound("snd/s4.wav", SND_ASYNC)
    Picture2.PaintPicture LoadPicture("img/e4.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
    Picture1.Picture = LoadPicture("img/p4.jpg")
    Shape2.Top = Picture1.Top - (Int(Rnd * (900 - 800 + 1)) + 800)
    Shape5.Top = Picture1.Top + (Int(Rnd * (150 - 100 + 1)) + 100)
    Shape3.Top = Picture1.Top - (Int(Rnd * (500 - 300 + 1)) + 300)
    Shape6.Top = Picture1.Top + (Int(Rnd * (900 - 800 + 1)) + 800)
    Shape4.Top = Picture1.Top - (Int(Rnd * (2000 - 1900 + 1)) + 1900)
    Shape7.Top = Picture1.Top - (Int(Rnd * (1500 - 1400 + 1)) + 1400) * 3
    dh = 100
    x_step2 = -dh
    y_step2 = dh
    x_step3 = dh
    y_step3 = dh
    x_step4 = -dh
    y_step4 = dh
    x_step6 = dh
    y_step6 = dh
    x_step5 = -dh
    y_step5 = dh
    x_step7 = dh
    y_step7 = dh
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer6.Enabled = True
    Timer4.Enabled = True
    Timer5.Enabled = True
    Timer7.Enabled = True
End Function

Public Function level9()
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer6.Enabled = False
    Timer5.Enabled = False
    Timer2.Enabled = False
    Timer7.Enabled = False
    Shape3.Top = Picture1.Top - 1000
    Shape4.Top = Picture1.Top - 1000
    Shape6.Top = Picture1.Top - 1000
    Shape5.Top = Picture1.Top - 1000
    Shape2.Top = Picture1.Top - 1000
    Shape7.Top = Picture1.Top - 1000

    If Form2.Option3.Value = True Then
         Line1.X2 = Line1.X2 - 110
    End If
    If Form2.Option1.Value = True Then
         move_x = move_x + 30
    Else
         move_x = move_x + 1
    End If
    temp = sndPlaySound("snd/s1.wav", SND_ASYNC)
    Picture2.PaintPicture LoadPicture("img/e5.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
    Picture1.Picture = LoadPicture("img/p5.jpg")
    Shape7.Top = Picture1.Top - (Int(Rnd * (1000 - 900 + 1)) + 900)
    Shape5.Top = Picture1.Top + (Int(Rnd * (150 - 100 + 1)) + 100)
    Shape2.Top = Picture1.Top - (Int(Rnd * (500 - 400 + 1)) + 400)
    Shape6.Top = Picture1.Top + (Int(Rnd * (1000 - 900 + 1)) + 900)
    Shape3.Top = Picture1.Top - (Int(Rnd * (1800 - 1700 + 1)) + 1700)
    Shape1.Top = Picture1.Top - (Int(Rnd * (1500 - 1400 + 1)) + 1400) * 4
    Shape4.Top = Picture1.Top - (Int(Rnd * (1500 - 1400 + 1)) + 1400) * 2
    dh = 100
    x_step2 = -dh
    y_step2 = dh
    x_step3 = dh
    y_step3 = dh
    x_step4 = dh
    y_step4 = dh
    x_step6 = -dh
    y_step6 = dh
    x_step5 = dh
    y_step5 = dh
    x_step7 = -dh
    y_step7 = dh
    x_step1 = dh
    y_step1 = dh
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer6.Enabled = True
    Timer4.Enabled = True
    Timer5.Enabled = True
    Timer7.Enabled = True
End Function

Public Function level10()
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer6.Enabled = False
    Timer5.Enabled = False
    Timer2.Enabled = False
    Timer7.Enabled = False
    Timer1.Enabled = False
    Shape3.Top = Picture1.Top - 1000
    Shape4.Top = Picture1.Top - 1000
    Shape6.Top = Picture1.Top - 1000
    Shape5.Top = Picture1.Top - 1000
    Shape2.Top = Picture1.Top - 1000
    Shape7.Top = Picture1.Top - 1000
    Shape1.Top = Picture1.Top - 1000

    '判断是否为困难难度，若是则缩小板子
    If Form2.Option3.Value = True Then
         Line1.X2 = Line1.X2 - 110
    End If
    '判断是否为简单难度，若不是则放慢板子移动速度
    If Form2.Option1.Value = True Then
         move_x = move_x + 30
    Else
         move_x = move_x + 1
    End If
    '通关音效播放
    temp = sndPlaySound("snd/s3.wav", SND_ASYNC)
    '通关图片显示
    Picture2.PaintPicture LoadPicture("img/e5.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
    Picture1.Picture = LoadPicture("img/p5.jpg")
    '小球竖直位置随机初始化
    Shape6.Top = Picture1.Top - (Int(Rnd * (1000 - 900 + 1)) + 900)
    Shape5.Top = Picture1.Top + (Int(Rnd * (150 - 100 + 1)) + 100)
    Shape4.Top = Picture1.Top - (Int(Rnd * (500 - 400 + 1)) + 400)
    Shape7.Top = Picture1.Top + (Int(Rnd * (800 - 700 + 1)) + 700)
    Shape8.Top = Picture1.Top - (Int(Rnd * (2000 - 1900 + 1)) + 1900)
    Shape2.Top = Picture1.Top - (Int(Rnd * (1500 - 1400 + 1)) + 1400) * 3
    Shape3.Top = Picture1.Top - (Int(Rnd * (2000 - 1900 + 1)) + 1900) * 3
    '小球
    dh = 110
    x_step2 = -dh
    y_step2 = dh
    x_step3 = dh
    y_step3 = dh
    x_step4 = dh
    y_step4 = dh
    x_step6 = -dh
    y_step6 = dh
    x_step5 = dh
    y_step5 = dh
    x_step7 = -dh
    y_step7 = dh
    x_step1 = dh
    y_step1 = dh
    x_step8 = -dh
    y_step8 = dh
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer6.Enabled = True
    Timer4.Enabled = True
    Timer5.Enabled = True
    Timer7.Enabled = True
    Timer8.Enabled = True
End Function

'开始按钮
Private Sub Command1_Click()
   '初始化血量和血条
   blood = 100
   Text1.Text = blood
   Text1.Left = 9480
   Line2.X2 = 9900
   defeat
   flag = 1
   Picture1.SetFocus
   Picture2.PaintPicture LoadPicture("img/e1.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
   '显示初始分数和等级
   Label2.Caption = 0
   Label4.Caption = level
   '根据选择的难度播放对应的背景音乐
   DoEvents
      If Form2.Option1.Value = True Then
         temp1 = mciExecute("open snd/y9.mid alias midi")
      ElseIf Form2.Option2.Value = True Then
         temp1 = mciExecute("open snd/y5.mid alias midi")
      ElseIf Form2.Option3.Value = True Then
         temp1 = mciExecute("open snd/y11.mid alias midi")
      End If
   temp1 = mciExecute("play midi")
   '开始第一关
   level1
   Command1.Enabled = False
End Sub

'退出按钮
Private Sub Command2_Click()
   '关闭背景音乐
   If flag = 1 Then
      temp = mciExecute("close midi")
   End If
   '初始化
   defeat

   gamescore = 0
   Form1.Label2.Caption = 0
   Form1.Label4.Caption = 0
   Form1.Line1.X1 = 1920
   Form1.Line1.X2 = 4320
   Form1.Hide
   '播放主菜单的背景音乐
   temp = sndPlaySound("snd/BGM.wav", SND_ASYNC Or SND_LOOP Or SND_NOWAIT)
   '显示主菜单界面
   Form2.Show
End Sub

'移动板子
Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 37 '如果按下左箭头,使板子向左移动
        If Line1.X1 <= Picture1.Left Then
           Line1.X1 = Picture1.Left
        Else
           Line1.X1 = Line1.X1 - (90 + move_x)
           Line1.X2 = Line1.X2 - (90 + move_x)
        End If
        
        Case 39 '如果按下右箭头,使板子向右移动
        If Line1.X2 >= Picture1.Left + Picture1.Width Then
           Line1.X2 = Picture1.Left + Picture1.Width
        Else
           Line1.X1 = Line1.X1 + (90 + move_x)
           Line1.X2 = Line1.X2 + (90 + move_x)
        End If
    End Select
End Sub
 
Private Sub Timer1_Timer()
    Label4.Caption = level
    Label2.Caption = gamescore
    '右壁弹回
    If Shape1.Left + Shape1.Width >= Picture1.Left + Picture1.Width Then
       Shape1.Left = Picture1.Left + Picture1.Width - Shape1.Width
       x_step1 = -x_step1
    End If
    
    '左壁弹回
    If Shape1.Left <= 0 Then
       Shape1.Left = 0
       x_step1 = -x_step1
    End If
    Label4.Caption = level
    
    '判断板子是否接住了小球
    If Shape1.Top + Shape1.Height >= Line1.Y1 Then
        '接住了
        If Shape1.Left + Shape1.Width / 2 >= Line1.X1 And Shape1.Left + Shape1.Width / 2 <= Line1.X2 Then
            temp = sndPlaySound("snd/tan1.wav", SND_ASYNC Or SND_NOSTOP)
            Shape1.Top = Picture1.Top - 1000
            gamescore = gamescore + 10
            
            '反弹
            Shape1.Top = Line1.Y1 - Shape1.Height
            y_step1 = -y_step1
        '没接住
        Else
            '扣血，困难难度扣10点血，否则扣5点血
            If Form2.Option3.Value = True Then
                blood = blood - 10
                Text1.Text = blood
                '血条变短
                Line2.X2 = Line2.X2 - 470 * 2
                Text1.Left = Text1.Left - 470 * 2
            Else
                blood = blood - 5
                gamescore = gamescore - 5
                Text1.Text = blood
                Line2.X2 = Line2.X2 - 470
                Text1.Left = Text1.Left - 470
            End If
            '扣分
            gamescore = gamescore - 5
            Picture2.PaintPicture LoadPicture("img/de.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
            '落到底部消失
            Timer1.Enabled = False
            Shape1.Top = Picture1.Top - 1000
            '判断是否失败
            If blood <= 0 Then
                temp1 = mciExecute("close midi")
                temp = sndPlaySound("snd/fail.wav", SND_ASYNC)
                MsgBox "你输了!!!!", 64
                defeat
            End If
        End If
    End If
    If level = 9 Then
        level = 10
        level10
    End If
    '使小球移动
    Shape1.Move Shape1.Left + x_step1, Shape1.Top + y_step1
End Sub
 
Private Sub Timer2_Timer()
    Label4.Caption = level
    Label2.Caption = gamescore
    '右壁弹回
    If Shape2.Left + Shape2.Width >= Picture1.Left + Picture1.Width Then
       Shape2.Left = Picture1.Left + Picture1.Width - Shape2.Width
       x_step2 = -x_step2
    End If
    
    '左壁弹回
    If Shape2.Left <= 0 Then
       Shape2.Left = 0
       x_step2 = -x_step2
    End If
    '判断板子是否接住
    If Shape2.Top + Shape2.Height >= Line1.Y1 Then
        If Shape2.Left + Shape2.Width / 2 >= Line1.X1 And Shape2.Left + Shape2.Width / 2 <= Line1.X2 Then
            temp = sndPlaySound("snd/tan1.wav", SND_ASYNC Or SND_NOSTOP)
            gamescore = gamescore + 10
            If blood < 100 Then
                blood = blood + 5
                Text1.Text = blood
                Line2.X2 = Line2.X2 + 470
                Text1.Left = Text1.Left + 470
            End If
            
            Shape2.Top = Line1.Y1 - Shape2.Height
            y_step2 = -y_step2
        Else
            '落到底部消失
            Timer2.Enabled = False
            Shape2.Top = Picture1.Top - 1000
            If blood <= 0 Then
                temp1 = mciExecute("close midi")
                temp = sndPlaySound("snd/fail.wav", SND_ASYNC)
                MsgBox "你输了!!!!", 64
                defeat
            End If
        End If
    End If
    
    If level = 6 Then
            level = 7
            level7
        End If
    '使小球移动
    Shape2.Move Shape2.Left + x_step2, Shape2.Top + y_step2
End Sub
 
Private Sub Timer3_Timer()
    Label4.Caption = level
    Label2.Caption = gamescore
    
    '右壁弹回
    If Shape3.Left + Shape3.Width >= Picture1.Left + Picture1.Width Then
       Shape3.Left = Picture1.Left + Picture1.Width - Shape3.Width
       x_step3 = -x_step3
    End If
    
    '左壁弹回
    If Shape3.Left <= 0 Then
       Shape3.Left = 0
       x_step3 = -x_step3
    End If
    
    '判断板子是否接住
    If Shape3.Top + Shape3.Height >= Line1.Y1 Then
        If Shape3.Left + Shape3.Width / 2 >= Line1.X1 And Shape3.Left + Shape3.Width / 2 <= Line1.X2 Then
            temp = sndPlaySound("snd/tan1.wav", SND_ASYNC Or SND_NOSTOP)
            Shape3.Top = Line1.Y1 - Shape3.Height
            y_step3 = -y_step3

            gamescore = gamescore + 5
            Label2.Caption = gamescore
        Else
            If Form2.Option3.Value = True Then
                blood = blood - 10
                Text1.Text = blood
                Line2.X2 = Line2.X2 - 470 * 2
                Text1.Left = Text1.Left - 470 * 2
            Else
                blood = blood - 5
                gamescore = gamescore - 5
                Text1.Text = blood
                Line2.X2 = Line2.X2 - 470
                Text1.Left = Text1.Left - 470
            End If
            gamescore = gamescore - 5
            Picture2.PaintPicture LoadPicture("img/de.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
            '落到底部消失
            Timer3.Enabled = False
            Shape3.Top = Picture1.Top - 1000
            If blood <= 0 Then
                temp1 = mciExecute("close midi")
                temp = sndPlaySound("snd/fail.wav", SND_ASYNC)
                MsgBox "你输了!!!!", 64
                defeat
            End If
        End If
        If level = 4 Then
            level = 5
            level5
        End If
        If level = 2 Then
            level = 3
            level3
        End If
        '判断胜利
        If level = 10 Then
            Picture2.PaintPicture LoadPicture("img/e3.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
            temp1 = mciExecute("close midi")
            '播放胜利音效
            temp = sndPlaySound("snd/vic.wav", SND_ASYNC)
            MsgBox "你赢了!!!!", 64
            defeat
        End If
    End If
    
     '使小球移动
    Shape3.Move Shape3.Left + x_step3, Shape3.Top + y_step3
End Sub


Private Sub Timer4_Timer()
    Label4.Caption = level
    Label2.Caption = gamescore
    
    '右壁弹回
    If Shape4.Left + Shape4.Width >= Picture1.Left + Picture1.Width Then
       Shape4.Left = Picture1.Left + Picture1.Width - Shape4.Width
       x_step4 = -x_step4
    End If
    
    '左壁弹回
    If Shape4.Left <= 0 Then
       Shape4.Left = 0
       x_step4 = -x_step4
    End If
    

    '判断板子是否接住
    If Shape4.Top + Shape4.Height >= Line1.Y1 Then
        If Shape4.Left + Shape4.Width / 2 >= Line1.X1 And Shape4.Left + Shape4.Width / 2 <= Line1.X2 Then
            temp = sndPlaySound("snd/tan1.wav", SND_ASYNC Or SND_NOSTOP)
            Shape4.Top = Line1.Y1 - Shape4.Height
            y_step4 = -y_step4
            Line1.X2 = Line1.X2 - 150
            blood = blood - 5
            gamescore = gamescore - 5
            Text1.Text = blood
            Line2.X2 = Line2.X2 - 470
            Text1.Left = Text1.Left - 470
        Else
            '落到底部消失
            Timer4.Enabled = False
            Shape4.Top = Picture1.Top - 1000
            If blood <= 0 Then
                temp1 = mciExecute("close midi")
                temp = sndPlaySound("snd/fail.wav", SND_ASYNC)
                MsgBox "你输了!!!!", 64
                defeat
            End If
        End If
    End If
    '使小球移动
    Shape4.Move Shape4.Left + x_step4, Shape4.Top + y_step4
End Sub


Private Sub Timer5_Timer()
    Label4.Caption = level
    Label2.Caption = gamescore
    '右壁弹回
    If Shape5.Left + Shape5.Width >= Picture1.Left + Picture1.Width Then
       Shape5.Left = Picture1.Left + Picture1.Width - Shape5.Width
       x_step5 = -x_step5
    End If
    
    '左壁弹回
    If Shape5.Left <= 0 Then
       Shape5.Left = 0
       x_step5 = -x_step5
    End If
    '判断板子是否接住
    If Shape5.Top + Shape5.Height >= Line1.Y1 Then
        If Shape5.Left + Shape5.Width / 2 >= Line1.X1 And Shape5.Left + Shape5.Width / 2 <= Line1.X2 Then
            temp = sndPlaySound("snd/tan1.wav", SND_ASYNC Or SND_NOSTOP)
            Shape5.Top = Line1.Y1 - Shape5.Height
            y_step5 = -y_step5
            gamescore = gamescore + 5
            Label2.Caption = gamescore
        Else
            If Form2.Option3.Value = True Then
                blood = blood - 10
                Text1.Text = blood
                Line2.X2 = Line2.X2 - 470 * 2
                Text1.Left = Text1.Left - 470 * 2
            Else
                blood = blood - 5
                gamescore = gamescore - 5
                Text1.Text = blood
                Line2.X2 = Line2.X2 - 470
                Text1.Left = Text1.Left - 470
            End If
            gamescore = gamescore - 5
            Picture2.PaintPicture LoadPicture("img/de.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
            '落到底部消失
            Timer5.Enabled = False
            Shape5.Top = Picture1.Top - 1000
            If blood <= 0 Then
                temp1 = mciExecute("close midi")
                temp = sndPlaySound("snd/fail.wav", SND_ASYNC)
                MsgBox "你输了!!!!", 64
                defeat
            End If
        End If
        If level = 5 Then
            level = 6
            level6
        End If
        If level = 3 Then
            level = 4
            level4
        End If
    End If
    
    '使小球移动
    Shape5.Move Shape5.Left + x_step5, Shape5.Top + y_step5
End Sub

Private Sub Timer6_Timer()
    Label4.Caption = level
    Label2.Caption = gamescore
    
    '右壁弹回
    If Shape6.Left + Shape6.Width >= Picture1.Left + Picture1.Width Then
       Shape6.Left = Picture1.Left + Picture1.Width - Shape6.Width
       x_step6 = -x_step6
    End If
    
    '左壁弹回
    If Shape6.Left <= 0 Then
       Shape6.Left = 0
       x_step6 = -x_step6
    End If

    '判断板子是否接住
    If Shape6.Top + Shape6.Height >= Line1.Y1 Then
        If Shape6.Left + Shape6.Width / 2 >= Line1.X1 And Shape6.Left + Shape6.Width / 2 <= Line1.X2 Then
            temp = sndPlaySound("snd/tan1.wav", SND_ASYNC Or SND_NOSTOP)
            Shape6.Top = Line1.Y1 - Shape6.Height
            y_step6 = -y_step6
            
            gamescore = gamescore + 10
            Label2.Caption = gamescore
        Else
            If Form2.Option3.Value = True Then
                blood = blood - 10
                Text1.Text = blood
                Line2.X2 = Line2.X2 - 470 * 2
                Text1.Left = Text1.Left - 470 * 2
            Else
                blood = blood - 5
                gamescore = gamescore - 5
                Text1.Text = blood
                Line2.X2 = Line2.X2 - 470
                Text1.Left = Text1.Left - 470
            End If
            gamescore = gamescore - 5
            Picture2.PaintPicture LoadPicture("img/de.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
            '落到底部消失
            Timer6.Enabled = False
            Shape6.Top = Picture1.Top - 1000
            If blood <= 0 Then
                temp1 = mciExecute("close midi")
                temp = sndPlaySound("snd/fail.wav", SND_ASYNC)
                MsgBox "你输了!!!!", 64
                defeat
            End If
        End If
        If level = 1 Then
            level = 2
            level2
        End If
    End If
    
    '使小球移动
    Shape6.Move Shape6.Left + x_step6, Shape6.Top + y_step6
End Sub

Private Sub Timer7_Timer()
    Label4.Caption = level
    Label2.Caption = gamescore
    '右壁弹回
    If Shape7.Left + Shape7.Width >= Picture1.Left + Picture1.Width Then
       Shape7.Left = Picture1.Left + Picture1.Width - Shape7.Width
       x_step7 = -x_step7
    End If
    
    '左壁弹回
    If Shape7.Left <= 0 Then
       Shape7.Left = 0
       x_step7 = -x_step7
    End If
    '判断板子是否接住
    If Shape7.Top + Shape7.Height >= Line1.Y1 Then
        If Shape7.Left + Shape7.Width / 2 >= Line1.X1 And Shape7.Left + Shape7.Width / 2 <= Line1.X2 Then
            temp = sndPlaySound("snd/tan1.wav", SND_ASYNC Or SND_NOSTOP)
            Shape7.Top = Line1.Y1 - Shape7.Height
            y_step7 = -y_step7
            gamescore = gamescore + 10
            Label2.Caption = gamescore
        Else
            If Form2.Option3.Value = True Then
                blood = blood - 10
                Text1.Text = blood
                Line2.X2 = Line2.X2 - 470 * 2
                Text1.Left = Text1.Left - 470 * 2
            Else
                blood = blood - 5
                gamescore = gamescore - 5
                Text1.Text = blood
                Line2.X2 = Line2.X2 - 470
                Text1.Left = Text1.Left - 470
            End If
            Picture2.PaintPicture LoadPicture("img/de.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
            '落到底部消失
            Timer7.Enabled = False
            Shape7.Top = Picture1.Top - 1000
            If blood <= 0 Then
                temp1 = mciExecute("close midi")
                temp = sndPlaySound("snd/fail.wav", SND_ASYNC)
                MsgBox "你输了!!!!", 64
                defeat
            End If
        End If
        If level = 8 Then
            level = 9
            level9
        End If
        If level = 7 Then
            level = 8
            level8
        End If
    End If
    
     '使小球移动
    Shape7.Move Shape7.Left + x_step7, Shape7.Top + y_step7
End Sub

Private Sub Timer8_Timer()
    Label4.Caption = level
    Label2.Caption = gamescore
    '右壁弹回
    If Shape8.Left + Shape8.Width >= Picture1.Left + Picture1.Width Then
       Shape8.Left = Picture1.Left + Picture1.Width - Shape8.Width
       x_step8 = -x_step8
    End If
    
    '左壁弹回
    If Shape8.Left <= 0 Then
       Shape8.Left = 0
       x_step8 = -x_step8
    End If
    '判断板子是否接住
    If Shape8.Top + Shape8.Height >= Line1.Y1 Then
        If Shape8.Left + Shape8.Width / 2 >= Line1.X1 And Shape8.Left + Shape8.Width / 2 <= Line1.X2 Then
            temp = sndPlaySound("snd/tan1.wav", SND_ASYNC Or SND_NOSTOP)
            Shape8.Top = Line1.Y1 - Shape8.Height
            y_step8 = -y_step8
            gamescore = gamescore + 10
            Label2.Caption = gamescore
        Else
            If Form2.Option3.Value = True Then
                blood = blood - 10
                Text1.Text = blood
                Line2.X2 = Line2.X2 - 470 * 2
                Text1.Left = Text1.Left - 470 * 2
            Else
                blood = blood - 5
                gamescore = gamescore - 5
                Text1.Text = blood
                Line2.X2 = Line2.X2 - 470
                Text1.Left = Text1.Left - 470
            End If
            Picture2.PaintPicture LoadPicture("img/de.jpg"), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
            '落到底部消失
            Timer8.Enabled = False
            Shape8.Top = Picture1.Top - 1000
            If blood <= 0 Then
                temp1 = mciExecute("close midi")
                temp = sndPlaySound("snd/fail.wav", SND_ASYNC)
                MsgBox "你输了!!!!", 64
                defeat
            End If
        End If
    End If
     '使小球移动
    Shape8.Move Shape8.Left + x_step8, Shape8.Top + y_step8
End Sub



