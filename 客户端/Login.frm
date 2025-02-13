VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "兵者-登录 *开发测试版本，不代表最终品质"
   ClientHeight    =   2175
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4695
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1285.062
   ScaleMode       =   0  'User
   ScaleWidth      =   4408.351
   StartUpPosition =   1  '所有者中心
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2280
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "twowlogin.piedaochuan.top"
      RemotePort      =   17115
   End
   Begin VB.CheckBox Check1 
      Caption         =   "记住密码(请勿在公共电脑上勾选此项)"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   960
      Width           =   3735
   End
   Begin VB.TextBox txt1 
      Height          =   345
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   3765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   390
      Left            =   960
      TabIndex        =   4
      Top             =   1320
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Enabled         =   0   'False
      Height          =   390
      Left            =   2760
      TabIndex        =   5
      Top             =   1320
      Width           =   1260
   End
   Begin VB.TextBox txt2 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   3765
   End
   Begin VB.Label Label2 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "没有账号?点这里去注册!"
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
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   4695
   End
   Begin VB.Label lblLabels 
      Caption         =   "用户名:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "密码:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ll, mm, message, islogin, allrooms



Private Sub Form_Load()
    
    Debugmode = False
    
    version = 3
    version_str = "V1.2"
    last_date = "2025-02-13"
    ll = "C:\Windows\TWOWZHMM.txt"
    
    islogin = "False"
    
    If Debugmode Then
        Winsock1.RemoteHost = "127.0.0.1"
        Sign.Winsock1.RemoteHost = "127.0.0.1"
        Sign.Winsock2.RemoteHost = "127.0.0.1"
        Sign.Winsock3.RemoteHost = "127.0.0.1"
    End If
    If Debugmode Then Label1.Caption = "[" + Time$ + "]" + "Debug模式"
    
    If Dir(ll) <> "" Then
    
        Open ll For Input As #1
        
            Line Input #1, AAA
            xx = Split(AAA, "###")
            zh = xx(0)
            mm = xx(1)
            txt1.Text = zh
            txt2.Text = mm
            
            If txt1.Text = "" Or txt2.Text = "" Then
                cmdOK.Enabled = False
            Else
                cmdOK.Enabled = True
            End If
    
        Close
        
        If mm = "" Then
            Check1.Value = 0
        Else
            Check1.Value = 1
        End If
        
    End If
    
    Label2.Caption = version_str
    
End Sub



Private Sub cmdOK_Click()

    Savezhmm
    
    zh = txt1.Text
    mm = txt2.Text
    
    message = "login" + Str(version) + " " + zh + " " + mm
    
    Label1.Caption = "[" + Time$ + "]" + "正在连接至服务器..."
    cmdCancel.Enabled = True
    cmdOK.Enabled = False
    
    Winsock1.Connect

End Sub



Private Sub cmdCancel_Click()

    Label1.Caption = "[" + Time$ + "]" + "连接已关闭."
    Winsock1.Close
    cmdOK.Enabled = True
    cmdCancel.Enabled = False
    
End Sub



Private Sub Form_Unload(Cancel As Integer)

    Savezhmm
    

End Sub



Private Sub Savezhmm()

    If Check1.Value = 1 Then
        Open ll For Output As #2
            Print #2, txt1.Text + "###" + txt2.Text
        Close
    End If
    
End Sub



Private Sub Label1_Click()

    If Not Debugmode Then Login.Hide
    Sign.Show
    
End Sub



Private Sub txt1_Change()

    If txt1.Text = "" Or txt2.Text = "" Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If

End Sub



Private Sub txt2_Change()

    If txt1.Text = "" Or txt2.Text = "" Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If

End Sub



Private Sub Winsock1_Close()

    Label1.Caption = "[" + Time$ + "]" + "连接已关闭."
    Winsock1.Close
    cmdOK.Enabled = True
    cmdCancel.Enabled = False

End Sub



Private Sub Winsock1_Connect()

    Label1.Caption = "[" + Time$ + "]" + "正在登陆..."
    Winsock1.SendData UTF8_Encode(message)

End Sub



Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    
    Dim xx() As Byte
    
    If islogin = "True" Then
        
        Dim x As String
        Winsock1.GetData xx
        x = Utf8ToUnicode(xx)
        ToMenu (x)

        If Debugmode Then Check1.Caption = "[" + Time$ + "]" + x
    
    End If
    
    
    
    If islogin = "False" Then
        
        
        Dim datat As String
        Winsock1.GetData xx
        datat = Utf8ToUnicode(xx)
        datat = Replace(datat, "  ", " ")
        data = Split(datat, " ")
        
        
        If Debugmode Then Check1.Caption = "[" + Time$ + "]" + datat
        
        
        If data(0) = "loginfail" Then
            Winsock1.Close
            tv = Int(data(1))
            If version < tv Then
                Label1.Caption = "[" + Time$ + "]" + "登陆失败，版本过低，请及时更新"
            End If
            If version > tv Then
                Label1.Caption = "[" + Time$ + "]" + "登陆失败，版本过高，请联系作者"
            End If
            cmdOK.Enabled = True
            cmdCancel.Enabled = False
        End If
        
        If data(0) = "账号密码错误!" Then
            Winsock1.Close
            Label1.Caption = "[" + Time$ + "]" + "账号密码错误!"
            cmdOK.Enabled = True
            cmdCancel.Enabled = False
        End If
        
        If data(0) = "登陆成功!" Then
            islogin = "True"
            Menu.Show
            If Not Debugmode Then Login.Hide
        End If
        
        If data(0) = "重复登陆!" Then
            Label1.Caption = "[" + Time$ + "]" + "账号异地登陆!"
            MsgBox "账号异地登陆!本客户端即将关闭。如果你不知情，请立即联系作者修改密码！", 13, "注意"
            End
        End If
        
        If data(0) = "ServerClose" Then
            Winsock1.Close
            MsgBox "服务器已关闭，即将退出客户端！", 13, "服务器关闭"
            End
        End If
        
        If data(0) = "tip" Then
            MsgBox data(1)
        End If
        
    End If
        
End Sub



Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    Winsock1.Close
    Label1.Caption = "[" + Time$ + "]" + "出现未知错误,请重新登录."
    cmdOK.Enabled = True
    cmdCancel.Enabled = False
    
End Sub



Public Function ToMenu(datat)
    
    datat = Replace(datat, "  ", " ")
    data = Split(datat, " ")
    If Debugmode Then Check1.Caption = "[" + Time$ + "]" + datat
    
    If data(0) = "selfinfo" Then
        money = Int(data(2))
        Menu.L1.Caption = "欢迎，" + data(1) + " 你的金币：" + data(2) + " 在线玩家数：" + data(3)
        If money >= 100 And Not Menu.TRoomName = "" Then
            Menu.CSet.Enabled = True
        Else
            Menu.CSet.Enabled = False
        End If
    End If
    
    If data(0) = "ServerClose" Then
        Winsock1.Close
        MsgBox "服务器已关闭，即将退出客户端！", 13, "服务器关闭"
        End
    End If
    
    If data(0) = "game" Then
        Game.Gamet (datat)
    End If
    
    If data(0) = "tip" Then
        MsgBox data(1)
    End If
    
    If data(0) = "CreateRoomSucess" Then
        roomname = data(1)
        Menu.Hide
        Game.Show
    End If
    
    If data(0) = "JoinRoomSucess" Then
        roomname = data(1)
        Menu.Hide
        Game.Show
    End If
    
    If data(0) = "nowrooms" Then
        Dim num As Integer
        num = Int(data(1))
        If num = 0 Then
            Menu.CJoin.Enabled = False
        Else
            Menu.CJoin.Enabled = True
        End If
        
        allrooms = Split(data(2), "###")
        Menu.List1.Clear
        For i = 0 To num - 1
            Menu.List1.AddItem allrooms(i)
        Next
    End If
    
    If data(0) = "reflash" Then
        Winsock1.SendData UTF8_Encode("selfinfo")
    End If
    
    If data(0) = "f**k" Then
        net_shou = Int(data(1))
    End If


End Function








