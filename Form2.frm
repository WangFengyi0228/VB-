VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "弹球游戏"
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   12960
   LinkTopic       =   "Form2"
   ScaleHeight     =   7140
   ScaleWidth      =   12960
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000005&
      Caption         =   "进入游戏"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   5160
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7515
      ScaleWidth      =   13635
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin VB.CommandButton Command2 
         Caption         =   "退出游戏"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3960
         TabIndex        =   6
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000014&
         Caption         =   "难度选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   3720
         TabIndex        =   2
         Top             =   3120
         Width           =   4935
         Begin VB.OptionButton Option3 
            BackColor       =   &H80000014&
            Caption         =   "困难（板子移动较慢且缩小，扣血较多）"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1320
            Width           =   4695
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000014&
            Caption         =   "一般（板子移动较慢）"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   2775
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000014&
            Caption         =   "简单"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_LOOP = &H8
Const SND_PURGE = &H40
Const SND_NOWAIT = &H2000

Option Explicit

Private Sub Command1_Click()
    '音频播放停止
    Dim temp As Long
    temp = sndPlaySound("snd/gap.wav", SND_PURGE)
    '切换界面
    Form2.Hide
    Form1.Show
End Sub

Private Sub Command2_Click()
    Dim temp As Long
    temp = sndPlaySound("snd/gap.wav", SND_PURGE)
    End
End Sub


Private Sub Form_Load()
    Dim temp As Long
    temp = sndPlaySound("snd/BGM.wav", SND_ASYNC Or SND_LOOP Or SND_NOWAIT)
    Form1.Hide
    Picture1.Picture = LoadPicture("img/bg.jpg")
    Option1.Value = True
    Option2.Value = False
    Option3.Value = False
    Form2.Left = (Screen.Width - Form2.Width) / 2
    Form2.Top = (Screen.Height - Form2.Height) / 2 - 600
End Sub

Private Sub Text1_Change()

End Sub

