VERSION 5.00
Begin VB.Form Game 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "兵者-房间 "
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   11910
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Caption         =   "对战记录"
      Height          =   3135
      Left            =   1680
      TabIndex        =   15
      Top             =   1560
      Width           =   3855
      Begin VB.TextBox Text1 
         Height          =   2775
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "Game.frx":0000
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame F3 
      Caption         =   "对手手牌"
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   5895
      Begin VB.Image Card 
         Height          =   855
         Index           =   17
         Left            =   5160
         Picture         =   "Game.frx":0006
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   16
         Left            =   4440
         Picture         =   "Game.frx":0B21
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   15
         Left            =   3720
         Picture         =   "Game.frx":163C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   14
         Left            =   3000
         Picture         =   "Game.frx":81D8
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   13
         Left            =   2280
         Picture         =   "Game.frx":ED74
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   12
         Left            =   1560
         Picture         =   "Game.frx":15910
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   11
         Left            =   840
         Picture         =   "Game.frx":1642B
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   10
         Left            =   120
         Picture         =   "Game.frx":1CFC7
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame F4 
      Caption         =   "对手被动卡槽"
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      Begin VB.Image Card 
         Height          =   855
         Index           =   19
         Left            =   840
         Picture         =   "Game.frx":23B63
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   18
         Left            =   120
         Picture         =   "Game.frx":2A6FF
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton CEndL 
      Caption         =   "结束回合"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10920
      MaskColor       =   &H80000010&
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.Frame F2 
      Caption         =   "你的被动卡槽"
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   10200
      TabIndex        =   1
      Top             =   5520
      Width           =   1575
      Begin VB.Image Card 
         Height          =   855
         Index           =   9
         Left            =   840
         Picture         =   "Game.frx":2B21A
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   8
         Left            =   120
         Picture         =   "Game.frx":2E35C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame F1 
      Caption         =   "你的手牌"
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   1320
      TabIndex        =   0
      Top             =   4920
      Width           =   8775
      Begin VB.Image Card 
         Height          =   1455
         Index           =   7
         Left            =   7680
         Picture         =   "Game.frx":2FAE7
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   6
         Left            =   6600
         Picture         =   "Game.frx":3145D
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   5
         Left            =   5520
         Picture         =   "Game.frx":335EE
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   4
         Left            =   4440
         Picture         =   "Game.frx":35157
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   3
         Left            =   3360
         Picture         =   "Game.frx":3696D
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   2
         Left            =   2280
         Picture         =   "Game.frx":37CCD
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   1
         Left            =   1200
         Picture         =   "Game.frx":387E8
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   0
         Left            =   120
         Picture         =   "Game.frx":3B3D4
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label5 
      Caption         =   "/6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "/6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   960
      Width           =   615
   End
   Begin VB.Label LNameE 
      Caption         =   "KirkLee123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9360
      TabIndex        =   12
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label LNameS 
      Caption         =   "KirkLee123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   11
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Image Card 
      Height          =   1455
      Index           =   26
      Left            =   7800
      Picture         =   "Game.frx":3CD5D
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Card 
      Height          =   1455
      Index           =   25
      Left            =   6720
      Picture         =   "Game.frx":3D878
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Card 
      Height          =   1455
      Index           =   24
      Left            =   5640
      Picture         =   "Game.frx":3E393
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Card 
      Height          =   1455
      Index           =   23
      Left            =   7800
      Picture         =   "Game.frx":3EEAE
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label L4 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "对手能量："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "你的能量："
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
      Left            =   120
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label L3 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Image Card 
      Height          =   1455
      Index           =   22
      Left            =   6720
      Picture         =   "Game.frx":3F9C9
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   975
   End
   Begin VB.Image Card 
      Height          =   1455
      Index           =   21
      Left            =   5640
      Picture         =   "Game.frx":404E4
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   975
   End
   Begin VB.Image Card 
      Height          =   1455
      Index           =   20
      Left            =   360
      Picture         =   "Game.frx":40FFF
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label L2 
      Caption         =   "188"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label L1 
      Caption         =   "牌堆剩余："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11880
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   4800
      Y2              =   4800
   End
End
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
    Dim i As Integer
    For i = 0 To 26
        Card(i).Picture = LoadPicture("res/0.jpg")
    Next
    LNameS.Caption = zh
    LNameE.Caption = "等待对手加入..."
    L2.Caption = "0"
    L3.Caption = "0"
    L4.Caption = "0"
    Text1.Text = ""
    Log ("欢迎进入房间!")
    Log ("等待对手加入...")
    CEndL.Enabled = False
    Game.Caption = "兵者-房间 " + roomname
End Sub












Public Function Gamet(datat)
    
    Dim Data
    datat = Replace(datat, "  ", " ")
    Data = Split(datat, " ")
    
    If Data(0) = "game" Then
        
        If Data(1) = "nowinfo" Then
        
        
        
        End If
        
        
        
    End If

End Function



Public Function Log(mmm)

    Text1.Text = Text1.Text + "[" + Time$ + "]" + mmm + vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
End Function



Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
