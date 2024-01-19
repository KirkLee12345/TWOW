VERSION 5.00
Begin VB.Form Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "兵者-主菜单"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9975
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame F3 
      Caption         =   "可加入房间列表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   6720
      TabIndex        =   4
      Top             =   960
      Width           =   3015
      Begin VB.CommandButton Command1 
         Caption         =   "刷新"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton CJoin 
         Caption         =   "加入房间"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   8
         Top             =   3000
         Width           =   2535
      End
      Begin VB.ListBox List1 
         Height          =   1860
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label L2 
         Caption         =   "点击选择上面的房间后点按钮"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2760
         Width           =   2535
      End
   End
   Begin VB.Frame F2 
      Caption         =   "创建房间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   3480
      TabIndex        =   3
      Top             =   960
      Width           =   3015
      Begin VB.TextBox TRoomName 
         Height          =   270
         Left            =   1080
         TabIndex        =   11
         Top             =   440
         Width           =   1695
      End
      Begin VB.CommandButton CSet 
         Caption         =   "创建房间"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   9
         Top             =   3000
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "创建房间需要花费100金币"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label LRoomName 
         Caption         =   "房间名称："
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame F1 
      Caption         =   "主菜单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   3015
      Begin VB.CommandButton CNothing 
         Caption         =   "Command1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton CAddMoney 
         Caption         =   "(测试)给自己加1金币"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton CAbout 
         Caption         =   "关于"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   6
         Top             =   3000
         Width           =   2535
      End
      Begin VB.CommandButton CRule 
         Caption         =   "教程"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Timer Trf 
      Left            =   9480
      Top             =   120
   End
   Begin VB.CommandButton Crf 
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9960
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label L1 
      Caption         =   """欢迎，"" + Data(1) + "" 你的金币："" + Data(2) + "" 在线玩家数："" + Data(3)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8340
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public onlineplayerscount



Private Sub CAbout_Click()

    If Not Debugmode Then Menu.Hide
    About.Show

End Sub



Private Sub CAddMoney_Click()
    Login.Winsock1.SendData "test moneyadd1"
End Sub



Private Sub CJoin_Click()
    If List1.Text = "" Then
        MsgBox "你还没有选择房间!"
    Else
        Login.Winsock1.SendData "room join " + List1.Text
        CJoin.Enabled = False
    End If
End Sub



Private Sub Command1_Click()
    Login.Winsock1.SendData "room r"
    Command1.Enabled = False
    Crf.Enabled = False
    Trf.Interval = 1000
End Sub



Private Sub Crf_Click()

    Login.Winsock1.SendData "selfinfo"
    Crf.Enabled = False
    Command1.Enabled = False
    Trf.Interval = 1000
    
End Sub



Private Sub CRule_Click()
    If Not Debugmode Then Menu.Hide
    Rule.Show
End Sub



Private Sub CSet_Click()
    Login.Winsock1.SendData "room create " + TRoomName.Text
End Sub



Private Sub Form_Load()

    L1.Caption = "欢迎! " + zh
    Login.Winsock1.SendData "selfinfo"
    
End Sub



Private Sub Form_Unload(Cancel As Integer)

    End

End Sub



Private Sub Trf_Timer()
    Trf.Interval = 0
    Crf.Enabled = True
    Command1.Enabled = True
End Sub



Private Sub TRoomName_Change()

    If money >= 100 And Not TRoomName = "" Then
        CSet.Enabled = True
    Else
        CSet.Enabled = False
    End If
    
End Sub
