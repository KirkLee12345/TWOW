VERSION 5.00
Begin VB.Form Rule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "兵者-规则 *开发测试版本，不代表最终品质"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19110
   Icon            =   "Rule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   19110
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame 
      Caption         =   "基础规则介绍"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   18855
      Begin VB.Label Label 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Index           =   7
         Left            =   2040
         TabIndex        =   15
         Top             =   360
         Width           =   16335
      End
      Begin VB.Image Card 
         Height          =   2175
         Index           =   7
         Left            =   240
         Picture         =   "Rule.frx":6988A
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "卡牌介绍 - 反转"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   6
      Left            =   9600
      TabIndex        =   12
      Top             =   7560
      Width           =   9375
      Begin VB.Image Card 
         Height          =   1575
         Index           =   6
         Left            =   240
         Picture         =   "Rule.frx":70426
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   $"Rule.frx":73568
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   6
         Left            =   1440
         TabIndex        =   13
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "卡牌介绍 - 扣能"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   4
      Left            =   9600
      TabIndex        =   8
      Top             =   5280
      Width           =   9375
      Begin VB.Frame Frame 
         Caption         =   "卡牌介绍 - 充能"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Index           =   5
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   9375
         Begin VB.Image Card 
            Height          =   1575
            Index           =   5
            Left            =   240
            Picture         =   "Rule.frx":736DA
            Stretch         =   -1  'True
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label 
            Caption         =   "充能：充能卡牌可以增加自己的能量，且使用充能卡牌可以无视能量上限，增加的能量数值为牌面点数。"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Index           =   5
            Left            =   1440
            TabIndex        =   11
            Top             =   360
            Width           =   7695
         End
      End
      Begin VB.Image Card 
         Height          =   1575
         Index           =   4
         Left            =   240
         Picture         =   "Rule.frx":75269
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "攻击：攻击卡牌可以对对方"
         Height          =   1575
         Index           =   4
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "卡牌介绍 - 扣能"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   3
      Left            =   9600
      TabIndex        =   6
      Top             =   3000
      Width           =   9375
      Begin VB.Image Card 
         Height          =   1575
         Index           =   3
         Left            =   240
         Picture         =   "Rule.frx":772C4
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "扣能：扣能卡牌可以扣减对方的能量，且使用扣能卡牌可以无视能量下限，扣除的能量数值为牌面点数。"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   3
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "卡牌介绍 - 回血"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   7560
      Width           =   9375
      Begin VB.Label Label 
         Caption         =   $"Rule.frx":7931F
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   2
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   7695
      End
      Begin VB.Image Card 
         Height          =   1575
         Index           =   2
         Left            =   240
         Picture         =   "Rule.frx":793D6
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "卡牌介绍 - 攻击"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   9375
      Begin VB.Image Card 
         Height          =   1575
         Index           =   1
         Left            =   240
         Picture         =   "Rule.frx":7AB61
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   $"Rule.frx":7C73C
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "卡牌介绍 - 护盾"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   9375
      Begin VB.Label Label 
         Caption         =   $"Rule.frx":7C816
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   7695
      End
      Begin VB.Image Card 
         Height          =   1575
         Index           =   0
         Left            =   240
         Picture         =   "Rule.frx":7C9EB
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Rule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Label(7).Caption = "兵者：兵者（The Way Of War）是一款以UNO牌底为基础的创新玩法的多人卡牌游戏。"
    Label(7).Caption = Label(7).Caption + vbCrLf + "对局开始时，每个玩家从牌堆随机抽取指定数量的手牌，由玩家轮流进行自己的回合。当玩家手动结束自己的回合或者打出攻击卡牌后，轮到下一个玩家进行回合。"
    Label(7).Caption = Label(7).Caption + vbCrLf + "结束回合时若玩家没有进行出牌操作（放置被动卡牌不属于出牌操作），则可以摸得一张手牌。"
    Label(7).Caption = Label(7).Caption + vbCrLf + "玩家在打出部分需要消耗能量的卡牌时，消耗后剩余能量不能超过能量下限，否则不能打出这张牌。"
    Label(7).Caption = Label(7).Caption + vbCrLf + "玩家的回合开始时可以获得2点能量，通过这种方式获得能量不能超过能量上限。"
    Label(7).Caption = Label(7).Caption + vbCrLf + "在兵者中，玩家的生命值为当前持有的手牌数量。当受到伤害后，将被随机扣除自己的手牌，扣除张数为受到的伤害点数。"
    Label(7).Caption = Label(7).Caption + vbCrLf + "当玩家的手牌持有量为0时，该玩家被淘汰。"
    Label(7).Caption = Label(7).Caption + vbCrLf + "请注意：玩家自己把仅剩的最后一张手牌打出，也算作手牌持有量为0。"
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Menu.Show
End Sub

