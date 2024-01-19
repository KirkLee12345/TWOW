VERSION 5.00
Begin VB.Form Rule 
   Caption         =   "±¯’ﬂ-ΩÃ≥Ã"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   10215
   StartUpPosition =   2  '∆¡ƒª÷––ƒ
   Begin VB.Frame Frame 
      Caption         =   "ø®≈∆ΩÈ…‹"
      BeginProperty Font 
         Name            =   "ÀŒÃÂ"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   5520
      Width           =   9375
      Begin VB.Label Label 
         Caption         =   $"Rule.frx":0000
         Height          =   1455
         Index           =   2
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   7695
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   2
         Left            =   240
         Picture         =   "Rule.frx":0198
         Stretch         =   -1  'True
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "ø®≈∆ΩÈ…‹"
      BeginProperty Font 
         Name            =   "ÀŒÃÂ"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   9375
      Begin VB.Image Card 
         Height          =   1455
         Index           =   1
         Left            =   240
         Picture         =   "Rule.frx":0CB3
         Stretch         =   -1  'True
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label 
         Caption         =   $"Rule.frx":17CE
         Height          =   1455
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "ø®≈∆ΩÈ…‹"
      BeginProperty Font 
         Name            =   "ÀŒÃÂ"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   9375
      Begin VB.Label Label 
         Caption         =   $"Rule.frx":1966
         Height          =   1455
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   7695
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   0
         Left            =   240
         Picture         =   "Rule.frx":1AFE
         Stretch         =   -1  'True
         Top             =   360
         Width           =   975
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

    Card(0).Picture = LoadPicture("res/b.jpg")

End Sub
