VERSION 5.00
Begin VB.Form Game 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����-���� *�������԰汾������������Ʒ�� "
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11910
   Icon            =   "Game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   11910
   StartUpPosition =   1  '����������
   Begin VB.Timer Timer1 
      Left            =   11400
      Top             =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ˢ��"
      Height          =   495
      Left            =   10920
      TabIndex        =   17
      Top             =   3120
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ս��¼"
      Height          =   3135
      Left            =   1680
      TabIndex        =   15
      Top             =   1560
      Width           =   3855
      Begin VB.TextBox Tchat 
         Height          =   270
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   2535
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "Game.frx":6988A
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame F3 
      Caption         =   "��������"
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   5895
      Begin VB.Image Card 
         Height          =   855
         Index           =   17
         Left            =   5160
         Picture         =   "Game.frx":69890
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   16
         Left            =   4440
         Picture         =   "Game.frx":6A3AB
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   15
         Left            =   3720
         Picture         =   "Game.frx":6AEC6
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   14
         Left            =   3000
         Picture         =   "Game.frx":71A62
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   13
         Left            =   2280
         Picture         =   "Game.frx":785FE
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   12
         Left            =   1560
         Picture         =   "Game.frx":7F19A
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   11
         Left            =   840
         Picture         =   "Game.frx":7FCB5
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   10
         Left            =   120
         Picture         =   "Game.frx":86851
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame F4 
      Caption         =   "���ֱ�������"
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   7680
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      Begin VB.Image Card 
         Height          =   855
         Index           =   19
         Left            =   840
         Picture         =   "Game.frx":8D3ED
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   18
         Left            =   120
         Picture         =   "Game.frx":93F89
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton CEndL 
      Caption         =   "�����غ�"
      BeginProperty Font 
         Name            =   "����"
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
      Top             =   2280
      Width           =   855
   End
   Begin VB.Frame F2 
      Caption         =   "��ı�������"
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
         Picture         =   "Game.frx":94AA4
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Card 
         Height          =   855
         Index           =   8
         Left            =   120
         Picture         =   "Game.frx":97BE6
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame F1 
      Caption         =   "������ƣ�˫�����Ƴ��ƣ�"
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
         Picture         =   "Game.frx":99371
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   6
         Left            =   6600
         Picture         =   "Game.frx":9ACE7
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   5
         Left            =   5520
         Picture         =   "Game.frx":9CE78
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   4
         Left            =   4440
         Picture         =   "Game.frx":9E9E1
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   3
         Left            =   3360
         Picture         =   "Game.frx":A01F7
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   2
         Left            =   2280
         Picture         =   "Game.frx":A1557
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   1
         Left            =   1200
         Picture         =   "Game.frx":A2072
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Card 
         Height          =   1455
         Index           =   0
         Left            =   120
         Picture         =   "Game.frx":A4C5E
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      Caption         =   ">> ���ƴ� >> "
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   9000
      TabIndex        =   19
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image Card 
      Height          =   2055
      Index           =   27
      Left            =   9480
      Picture         =   "Game.frx":A65E7
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label LT 
      Caption         =   "LT"
      Height          =   975
      Left            =   8880
      TabIndex        =   18
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "/6"
      BeginProperty Font 
         Name            =   "����"
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
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "/6"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label LNameE 
      Caption         =   "KirkLee123"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Picture         =   "Game.frx":A7102
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Card 
      Height          =   1455
      Index           =   25
      Left            =   6720
      Picture         =   "Game.frx":A7C1D
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Card 
      Height          =   1455
      Index           =   24
      Left            =   5640
      Picture         =   "Game.frx":A8738
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Card 
      Height          =   1455
      Index           =   23
      Left            =   7800
      Picture         =   "Game.frx":A9253
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label L4 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���������"
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
      Left            =   120
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label L3 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "����"
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
      Picture         =   "Game.frx":A9D6E
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   975
   End
   Begin VB.Image Card 
      Height          =   1455
      Index           =   21
      Left            =   5640
      Picture         =   "Game.frx":AA889
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   975
   End
   Begin VB.Image Card 
      Height          =   1455
      Index           =   20
      Left            =   360
      Picture         =   "Game.frx":AB3A4
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label L2 
      Caption         =   "188"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�ƶ�ʣ�ࣺ"
      BeginProperty Font 
         Name            =   "����"
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
Public cards As Object



Private Sub Card_DblClick(Index As Integer)
    Dim c As String
    c = cards(Index)
    If c = "d0" Or c = "g0" Or c = "k0" Or c = "n0" Or c = "w2" Or c = "w4" Then
        Dim answer As VbMsgBoxResult
        answer = MsgBox("�ÿ��ƿ���ֱ�Ӵ�������Ϊ�������ƣ�ѡ���ǡ�ֱ�Ӵ����ѡ�񡮷񡯷���Ϊ������(��ʾ����תֱ�Ӵ����ʱ��Ч��)", vbYesNoCancel + vbExclamation, "ѡ����Ʒ�ʽ")
        Select Case answer
            Case vbYes
                Login.Winsock1.SendData UTF8_Encode("game use " + Str(Index))
            Case vbNo
                Login.Winsock1.SendData UTF8_Encode("game pass " + Str(Index))
        End Select
    Else
        Login.Winsock1.SendData UTF8_Encode("game use " + Str(Index))
    End If
End Sub

Private Sub Card_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim c As String
    c = cards(Index)
    
    If (Index >= 21 And Index <= 26) Or Index = 20 Or Index = 27 Then
        If Index >= 21 And Index <= 26 Then
            LT.Caption = "���ܿ��ۣ����ܿ��Ե����˺�����ʣ��1�㻤��ʱΪ�޵жܣ����Ե�������ʣ���˺���"
        End If
        If Index = 20 Then
            LT.Caption = "ʣ���ƶѡ��������κβ���ֱ�ӽ����غϿɴӴ���һ���ơ�"
        End If
        If Index = 27 Then
            LT.Caption = "��һ�ű��������"
        End If
    Else
        If Mid(c, 1, 1) = "0" Or Mid(c, 1, 1) = "b" Then
            If Mid(c, 1, 1) = "0" Then
                LT.Caption = "������ۿտ���Ҳ"
            End If
            If Mid(c, 1, 1) = "b" Then
                LT.Caption = "δ֪����"
            End If
        Else
            If Mid(c, 2, 1) = "0" Then
                LT.Caption = "��ת�������������Ч�������������Է������ͬ��ɫ�Ŀ���ʱ������������Ч�����ö���ת��"
            Else
                If Mid(c, 1, 1) = "d" Then
                    LT.Caption = "���ܣ�����" + Mid(c, 2, 1) + "�����������Լ�����" + Mid(c, 2, 1) + "�㻤�ܡ����ܿ��Ե����˺���"
                    If Mid(c, 2, 1) = "1" Then
                        LT.Caption = LT.Caption + "��ʣ��1�㻤��ʱΪ�޵жܣ����Ե�������ʣ���˺���"
                    End If
                End If
                If Mid(c, 1, 1) = "g" Then
                    LT.Caption = "����������" + Mid(c, 2, 1) + "�����������Է����" + Mid(c, 2, 1) + "���˺�����������߶Է�" + Mid(c, 2, 1) + "�����ơ��������ƺ����̽����غϡ�"
                End If
                If Mid(c, 1, 1) = "n" Then
                    LT.Caption = "���ܣ����Լ�����" + Mid(c, 2, 1) + "�����������Գ������ޡ�"
                End If
                If Mid(c, 1, 1) = "k" Then
                    LT.Caption = "���ܣ��öԷ���������" + Mid(c, 2, 1) + "�㣬���Գ������ޡ�"
                End If
                If Mid(c, 1, 1) = "w" Then
                    LT.Caption = "��Ѫ�����Լ�����" + Mid(c, 2, 1) + "�����ƣ����ɳ������ޡ����������Լ����Ʊ�������ʱ���������Լ�����" + Mid(c, 2, 1) + "�����Ƹ��"
                End If
                
            End If
        End If
    End If
    
    
    
End Sub

Private Sub CEndL_Click()
    Login.Winsock1.SendData UTF8_Encode("game next")
End Sub

Private Sub CEndL_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    LT.Caption = "�����غϣ�������ɺ�������תΪ�Է��Ļغϡ��������κβ���(���������ñ���)ֱ�ӵ����������һ�����ơ�"
End Sub

Private Sub Command1_Click()
    Login.Winsock1.SendData UTF8_Encode("game nowinfo")
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    LT.Caption = "ˢ�£��ӷ��������»�ȡ�Ծ���Ϣ��̫��û��Ӧ�볢�Ե��ˢ�¡�"
End Sub

Private Sub Form_Load()
    Set cards = CreateObject("Scripting.Dictionary")
    Dim i As Integer
    For i = 0 To 26
        Card(i).Picture = LoadPicture("res/0.jpg")
    Next
    LNameS.Caption = zh
    LNameE.Caption = "�ȴ����ּ���..."
    L2.Caption = "0"
    L3.Caption = "0"
    L4.Caption = "0"
    Text1.Text = ""
    Log ("��ӭ���뷿�� " + roomname + " !")
    Log ("�ȴ����ּ���...")
    CEndL.Enabled = False
    Command1.Enabled = False
    Game.Caption = "����-���� " + roomname + " *�������԰汾������������Ʒ��"
End Sub












Public Function Gamet(datat)
    
    Dim data
    datat = Replace(datat, "  ", " ")
    data = Split(datat, " ")
    
    If data(0) = "game" Then
        
        If data(1) = "start" Then
            LNameE.Caption = data(2)
            Log ("���� " + data(2) + " �����˷���!")
            Log ("��Ϸ������ʼ!")
            Login.Winsock1.SendData UTF8_Encode("game start")
            Timer1.Interval = 1000
        End If
        
        If data(1) = "exit" Then
            MsgBox "�Է��˳��˷��䣬�����ѹر�"
        End If
        
        If data(1) = "nowinfo" Then
            Dim i As Integer
            For i = 2 To 28
                Card(i - 2).Picture = LoadPicture("res/" + data(i) + ".jpg")
                cards(i - 2) = data(i)
            Next
            L3.Caption = Int(data(29))
            L4.Caption = Int(data(30))
            L2.Caption = Int(data(31))
            If data(32) = "1" Then
                Log ("��������Ļغ�!")
                CEndL.Enabled = True
            Else
                Log ("�����Ƕ��ֵĻغ�...")
                CEndL.Enabled = False
            End If
            Card(27).Picture = LoadPicture("res/" + data(33) + ".jpg")
        
        End If
        
        If data(1) = "log" Then
            Log (data(2))
        End If
        
        If data(1) = "end" Then
            If data(2) = "win" Then
                MsgBox "��ϲ��ȡ�öԾ�ʤ����"
            Else
                MsgBox "�����ˣ�����һ�Ѱɣ�"
            End If
        End If
        
        
    End If

End Function



Public Function Log(mmm)

    Text1.Text = Text1.Text + "[" + Time$ + "]" + mmm + vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
End Function



Private Sub Form_Unload(Cancel As Integer)
    Dim myexit
    myexit = MsgBox("��ȷ��Ҫ�˳�����Ϸ���˳���ֱ�ӹرշ��䣬�Է�Ҳ�ᱻ�߳��Ծ�", vbExclamation + vbYesNo + vbDefaultButton2, "�˳�ȷ��...")
    If myexit = vbNo Then
        Cancel = True
    End If
    If myexit <> vbNo Then
        Login.Winsock1.SendData UTF8_Encode("room exit")
        Menu.Show
    End If
End Sub

Private Sub L2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    LT.Caption = "ʣ���ƶѵ����������ƶѱ����껹δ�����Ծ֣���ƽ�֡�"
End Sub

Private Sub L3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    LT.Caption = "�㵱ǰӵ�е�����������ĻغϿ�ʼʱ������2�����������ɳ������ޡ�ʹ�ù���/���ܿ��ƻ�������Ӧ������ʹ���������ƿ��Ը��Լ��������������Գ������ޡ�"
End Sub

Private Sub L4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    LT.Caption = "�Է���ǰӵ�е����������Է��ĻغϿ�ʼʱ������2�����������ɳ������ޡ�ʹ�ÿ��ܿ��ƿ��Լ��ٶԷ����������Գ������ޡ�"
End Sub

Private Sub LNameE_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    LT.Caption = "���ֵ��û�����"
End Sub

Private Sub LNameS_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    LT.Caption = "����û�����"
End Sub

Private Sub Tchat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim msg As String
        msg = "game chat " + Tchat.Text
        Login.Winsock1.SendData UTF8_Encode(msg)
        Tchat.Text = ""
    End If
End Sub

Private Sub Timer1_Timer()
    Login.Winsock1.SendData UTF8_Encode("game nowinfo")
    Timer1.Interval = 0
    Command1.Enabled = True
End Sub
