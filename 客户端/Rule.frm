VERSION 5.00
Begin VB.Form Rule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����-���� *�������԰汾������������Ʒ��"
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
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame 
      Caption         =   "�����������"
      BeginProperty Font 
         Name            =   "����"
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
            Name            =   "΢���ź�"
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
      Caption         =   "���ƽ��� - ��ת"
      BeginProperty Font 
         Name            =   "����"
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
            Name            =   "΢���ź�"
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
      Caption         =   "���ƽ��� - ����"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "���ƽ��� - ����"
         BeginProperty Font 
            Name            =   "����"
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
            Caption         =   "���ܣ����ܿ��ƿ��������Լ�����������ʹ�ó��ܿ��ƿ��������������ޣ����ӵ�������ֵΪ���������"
            BeginProperty Font 
               Name            =   "΢���ź�"
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
         Caption         =   "�������������ƿ��ԶԶԷ�"
         Height          =   1575
         Index           =   4
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "���ƽ��� - ����"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "���ܣ����ܿ��ƿ��Կۼ��Է�����������ʹ�ÿ��ܿ��ƿ��������������ޣ��۳���������ֵΪ���������"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Caption         =   "���ƽ��� - ��Ѫ"
      BeginProperty Font 
         Name            =   "����"
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
            Name            =   "΢���ź�"
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
      Caption         =   "���ƽ��� - ����"
      BeginProperty Font 
         Name            =   "����"
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
            Name            =   "΢���ź�"
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
      Caption         =   "���ƽ��� - ����"
      BeginProperty Font 
         Name            =   "����"
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
            Name            =   "΢���ź�"
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
    Label(7).Caption = "���ߣ����ߣ�The Way Of War����һ����UNO�Ƶ�Ϊ�����Ĵ����淨�Ķ��˿�����Ϸ��"
    Label(7).Caption = Label(7).Caption + vbCrLf + "�Ծֿ�ʼʱ��ÿ����Ҵ��ƶ������ȡָ�����������ƣ���������������Լ��Ļغϡ�������ֶ������Լ��Ļغϻ��ߴ���������ƺ��ֵ���һ����ҽ��лغϡ�"
    Label(7).Caption = Label(7).Caption + vbCrLf + "�����غ�ʱ�����û�н��г��Ʋ��������ñ������Ʋ����ڳ��Ʋ����������������һ�����ơ�"
    Label(7).Caption = Label(7).Caption + vbCrLf + "����ڴ��������Ҫ���������Ŀ���ʱ�����ĺ�ʣ���������ܳ����������ޣ������ܴ�������ơ�"
    Label(7).Caption = Label(7).Caption + vbCrLf + "��ҵĻغϿ�ʼʱ���Ի��2��������ͨ�����ַ�ʽ����������ܳ����������ޡ�"
    Label(7).Caption = Label(7).Caption + vbCrLf + "�ڱ����У���ҵ�����ֵΪ��ǰ���е��������������ܵ��˺��󣬽�������۳��Լ������ƣ��۳�����Ϊ�ܵ����˺�������"
    Label(7).Caption = Label(7).Caption + vbCrLf + "����ҵ����Ƴ�����Ϊ0ʱ������ұ���̭��"
    Label(7).Caption = Label(7).Caption + vbCrLf + "��ע�⣺����Լ��ѽ�ʣ�����һ�����ƴ����Ҳ�������Ƴ�����Ϊ0��"
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Menu.Show
End Sub

