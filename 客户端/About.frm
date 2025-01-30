VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "兵者-关于 *开发测试版本，不代表最终品质"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4110
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4110
   StartUpPosition =   1  '所有者中心
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label15 
      Caption         =   "Label15"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "兵者The Way Of War"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "(C) 2022-2025 氕氘氚工作室 版权所有"
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
      TabIndex        =   0
      Top             =   4320
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   1605
      Left            =   0
      Picture         =   "About.frx":6988A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4125
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()

    Label3.Caption = "版本：" + version_str
    Label4.Caption = "协议版本：" + Str(version)
    Label5.Caption = "更新日期：" + last_date
    Label6.Caption = ""
    Label7.Caption = ""
    Label8.Caption = ""
    Label9.Caption = ""
    Label10.Caption = ""
    Label11.Caption = "程序：KirkLee123"
    Label12.Caption = "美术：KirkLee123"
    Label13.Caption = "设计：KirkLee123"
    Label14.Caption = "运营：KirkLee123"
    Label15.Caption = "策划："
    Label16.Caption = ""
    Label17.Caption = ""
    Label18.Caption = ""

End Sub



Private Sub Form_Unload(Cancel As Integer)

    Menu.Show

End Sub
