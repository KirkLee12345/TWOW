VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Sign 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "兵者-注册 *开发测试版本，不代表最终品质"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "Sign.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4560
   StartUpPosition =   1  '所有者中心
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   4200
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "twowlogin.piedaochuan.top"
      RemotePort      =   17115
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "twowlogin.piedaochuan.top"
      RemotePort      =   17115
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   3480
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "twowlogin.piedaochuan.top"
      RemotePort      =   17115
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   2520
   End
   Begin VB.CommandButton Command2 
      Caption         =   "注册!"
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   960
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      TabIndex        =   3
      Top             =   1680
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "获取验证码"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   1370
      Width           =   3495
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   400
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "验证码:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "绑定邮箱:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "确认密码:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "密码:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "用户名:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Sign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ema, yzmt, zhmmyx



Private Sub Command1_Click()
    
    Command1.Enabled = False
    ema = Text1.Text
    Winsock3.Connect

End Sub



Private Sub Command2_Click()
    If Label6.Caption = "*用户名可用*" Then
        If Label7.Caption = "*没有问题*" Then
    
            zhmmyx = " " + Text2.Text + " " + Text3.Text + " " + ema
            
            If Winsock2.State <> sckClosed Then Winsock2.Close
            Winsock2.Connect

        Else
            Label8.Caption = "密码不一致!"
        End If
    Else
        If Label6.Caption = "" Then
            Label8.Caption = "无法连接至服务器!"
        Else
            Label8.Caption = "用户名不可用!"
        End If
    End If
End Sub



Private Sub Form_Load()

    Command1.Enabled = False
    Command2.Enabled = False
    
End Sub



Private Sub Form_Unload(Cancel As Integer)

    Login.Show
    
End Sub



Private Sub Text1_Change()
    
    sx
    
    Dim E_mail As String
    Dim rExp As New RegExp
    rExp.Pattern = "\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"
    E_mail = Text1.Text
    rExp.Test (E_mail)
    
    
    If yzmt < 1 And rExp.Test(E_mail) Then
        Command1.Enabled = True
    Else
        Command1.Enabled = False
    End If
    
End Sub



Private Sub Text2_Change()
    
    sx
    
    If Winsock1.State <> sckClosed Then Winsock1.Close
    Winsock1.Connect

End Sub



Private Sub Text3_Change()
    sx
    If Text3.Text = "" Or Text4.Text = "" Then
        Label7.Caption = ""
    Else
        If Text3.Text <> Text4.Text Then
            Label7.Caption = "两次输入密码不一致!"
        Else
            Label7.Caption = "*没有问题*"
        End If
    End If
End Sub



Private Sub Text4_Change()
    sx
    If Text3.Text = "" Or Text4.Text = "" Then
        Label7.Caption = ""
    Else
        If Text3.Text <> Text4.Text Then
            Label7.Caption = "两次输入密码不一致!"
        Else
            Label7.Caption = "*没有问题*"
        End If
    End If
End Sub



Private Sub Text5_Change()
    sx
End Sub



Private Sub Timer1_Timer()
    yzmt = yzmt - 1
    Command1.Caption = "获取验证码(" + Str(yzmt) + ")"
    If yzmt = 0 Then
        Command1.Enabled = True
        Command1.Caption = "获取验证码"
        Timer1.Interval = 0
    End If
End Sub



Private Sub Winsock1_Connect()

    Dim x As String
    x = "sign username " + Text2.Text
    Winsock1.SendData UTF8_Encode(x)
    
End Sub



Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim x As String
    Dim xx() As Byte
    Winsock1.GetData xx
    x = Utf8ToUnicode(xx)
    Label6.Caption = x
    Winsock1.Close
    
End Sub



Private Sub Winsock2_Connect()

    Winsock2.SendData UTF8_Encode("sign up " + zhmmyx + " " + Text5.Text)
    
End Sub



Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)

    Dim x As String
    Dim xx() As Byte
    Winsock2.GetData xx
    x = Utf8ToUnicode(xx)
    Label8.Caption = x
    Winsock2.Close
    
    If x = "注册成功!" Then
        Command2.Enabled = False
        Label8.Caption = "注册成功!可以关闭此窗口去登陆了!"
    End If

End Sub



Private Sub Winsock3_Connect()

    Winsock3.SendData UTF8_Encode("sign ema " + ema)

End Sub



Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)

    Dim x As String
    Dim xx() As Byte
    Winsock3.GetData xx
    x = Utf8ToUnicode(xx)
    If x = "sand yzm sucess" Then
        MsgBox "验证码发送成功，请及时查看！", 13, "验证码"
        yzmt = 60
        Timer1.Interval = 1000
        Command1.Enabled = False
    Else
        MsgBox x
    End If
    Winsock3.Close

End Sub



Public Function sx()
    
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
        Command2.Enabled = False
    Else
        Command2.Enabled = True
    End If
    
End Function

