VERSION 5.00
Begin VB.Form Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����-���˵� *�������԰汾������������Ʒ��"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9975
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9975
   StartUpPosition =   1  '����������
   Begin VB.Timer TFuck 
      Left            =   9480
      Top             =   720
   End
   Begin VB.Frame F3 
      Caption         =   "�����б�"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "ˢ��"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton CJoin 
         Caption         =   "���뷿��"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "���ѡ������ķ����㰴ť"
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "��������"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
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
      Begin VB.Label LRT 
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "����������Ҫ����100���"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�������ƣ�"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame F1 
      Caption         =   "���˵�"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "�����ڴ�"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "(����)���Լ���10���"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "ˢ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   """��ӭ��"" + Data(1) + "" ��Ľ�ң�"" + Data(2) + "" �����������"" + Data(3)"
      BeginProperty Font 
         Name            =   "����"
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
    Login.Winsock1.SendData UTF8_Encode("test moneyadd1")
End Sub



Private Sub CJoin_Click()
    If List1.Text = "" Then
        MsgBox "�㻹û��ѡ�񷿼�!"
    Else
        Login.Winsock1.SendData UTF8_Encode("room join " + Split(List1.Text, "��")(0))
        CJoin.Enabled = False
    End If
End Sub



Private Sub Command1_Click()
    Login.Winsock1.SendData UTF8_Encode("room r")
    Command1.Enabled = False
    Crf.Enabled = False
    Trf.Interval = 1000
End Sub



Private Sub Crf_Click()

    Login.Winsock1.SendData UTF8_Encode("selfinfo")
    Crf.Enabled = False
    Command1.Enabled = False
    Trf.Interval = 1000
    
End Sub



Private Sub CRule_Click()
    If Not Debugmode Then Menu.Hide
    Rule.Show
End Sub



Private Sub CSet_Click()
    Login.Winsock1.SendData UTF8_Encode("room create " + TRoomName.Text)
End Sub



Private Sub Form_Load()
    
    net_fa = 0
    L1.Caption = "��ӭ! " + zh
    Login.Winsock1.SendData UTF8_Encode("selfinfo")
    TFuck.Interval = 1000
    
End Sub



Private Sub Form_Unload(Cancel As Integer)

    End

End Sub



Private Sub TFuck_Timer()
    net_fa = net_fa + 1
    Login.Winsock1.SendData UTF8_Encode("f**k " + Str(net_fa))
    If net_fa - net_shou > 10 Then
        Menu.Caption = "����-���˵� - �����쳣 *�������԰汾������������Ʒ��"
    Else
        Menu.Caption = "����-���˵� *�������԰汾������������Ʒ��"
    End If
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
    
    Dim data
    data = Split(TRoomName, " ")
    If UBound(data) - LBound(data) > 0 Then
        CSet.Enabled = False
        LRT.Caption = "���������ܰ����ո�!"
    Else
        LRT.Caption = ""
    End If
    
End Sub
