VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Server 
   Caption         =   "����-�����"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   10110
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ListBox List2 
      Height          =   3480
      Left            =   8400
      TabIndex        =   9
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton CDebug 
      Caption         =   "DebugMode��False"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   120
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   3480
      Left            =   6720
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   7320
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox TPort 
      Height          =   270
      Left            =   720
      TabIndex        =   6
      Text            =   "17115"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CStop 
      Caption         =   "�رշ����"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton CStart 
      Caption         =   "���������"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton CSaveLog 
      Caption         =   "������־"
      Height          =   615
      Left            =   4080
      TabIndex        =   2
      Top             =   4560
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6600
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   9855
   End
   Begin VB.TextBox Text1 
      Height          =   3495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   6375
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   1
      Left            =   7800
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "�˿ڣ�"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public version, version_str, zh, mm, gb, txtname, infotxtname, server_port, Debugmode, tRoom

Public IndexName, IndexIfLogin, ConnectState, Emayzm, UserIsLogin, WsIsBusy, IndexRoom As Object

Public NameMoney, ZhuKe As Object



Private Sub Form_Load()
    
    Debugmode = False
    
    CStart.Enabled = True
    CStop.Enabled = False
    
    version_str = "V1.0"
    version = 1
    txtname = "\zhmm.txt"
    infotxtname = "\playersdata.txt"
    server_port = 17115
    
    Winsock1.LocalPort = server_port
    
    Set IndexName = CreateObject("Scripting.Dictionary")
    Set IndexIfLogin = CreateObject("Scripting.Dictionary")
    Set ConnectState = CreateObject("Scripting.Dictionary")
    Set Emayzm = CreateObject("Scripting.Dictionary")
    Set UserIsLogin = CreateObject("Scripting.Dictionary")
    Set WsIsBusy = CreateObject("Scripting.Dictionary")
    Set NameMoney = CreateObject("Scripting.Dictionary")
    Set IndexRoom = CreateObject("Scripting.Dictionary")
    Set RoomState = CreateObject("Scripting.Dictionary")
    Set ZhuKe = CreateObject("Scripting.Dictionary")


End Sub



Private Sub CStart_Click()
    
    CStart.Enabled = False
    CStop.Enabled = True
    
    If Debugmode Then Log ("������debugģʽ����...")
    If Debugmode Then Log ("�˺������ļ���" + txtname)
    If Debugmode Then Log ("��������ļ���" + infotxtname)
    
    Log ("����TWOW���������...")
    Log ("�˷���˵İ汾Ϊ " + version_str + " Э��汾Ϊ" + Str(version))
    Log ("�����˿�:" + Str(Winsock1.LocalPort))
    
    WsIsBusy.Item(0) = "False"
    WsIsBusy.Item(1) = "False"
    
    Winsock1.Listen
    
    '��ȡ�������    ����������
    If Dir(App.Path & infotxtname) <> "" Then
        If Debugmode Then Log ("���ڼ����������...")
        If Debugmode Then Log ("�����    ���")
        Open App.Path & infotxtname For Input As #1
            
            Dim now As String
            Do While Not EOF(1)
                
                Line Input #1, now
                info = Split(now, ",")
                
                If UBound(info) > 0 Then
                    
                    If Debugmode Then Log (info(0) + " " + info(1))
                    NameMoney.Item(info(0)) = Int(info(1))
                    
                    
                End If
                
            Loop
        
        Close #1
    End If
    
    
End Sub



Private Sub CStop_Click()
    
    Log "���ڹرշ����..."
    
    CStart.Enabled = True
    CStop.Enabled = False
    
    Winsock1.Close
    
    Dim i As Long
    For i = 0 To sckServer.Count - 1
        If sckServer(i).State <> sckClosed Then
            sckServer(i).SendData "ServerClose"
            'sckServer(i).Close
        End If
    Next
    

    IndexName.removeall
    IndexIfLogin.removeall
    WsIsBusy.removeall
    IndexName.removeall
    IndexIfLogin.removeall
    ConnectState.removeall
    Emayzm.removeall
    UserIsLogin.removeall
    WsIsBusy.removeall
    IndexRoom.removeall
    
    
    '�����������    ����������
    If Debugmode Then Log ("���ڱ����������...")
    If Debugmode Then Log ("�����    ���")
    Dim now As String
    now = ""
    
    Dim keys
    keys = NameMoney.keys
    For i = 0 To UBound(keys)
        
        now = now + keys(i) + "," + Str(NameMoney.Item(keys(i))) + vbCrLf
        
    Next
    
    Open App.Path & infotxtname For Output As #1
        If Debugmode Then Log (now)
        Print #1, now
    Close #1
    
    NameMoney.removeall
    
    
End Sub



Private Sub CSaveLog_Click()

    SaveLog
    
End Sub



Private Sub SaveLog()
    
    Dim stn As String
    stn = "\TWOWLOG\" + Mid(Date$, 1, 4) + "." + Mid(Date$, 6, 2) + "." + Mid(Date$, 9, 2) + "..." + Mid(Time$, 1, 2) + "." + Mid(Time$, 4, 2) + "." + Mid(Time$, 7, 2) + ".log" + ".txt"
    Open App.Path & stn For Output As #2
    Print #2, Text1.Text
    Close
    MsgBox "��־�ѱ���!", , "����-�����"

End Sub




Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    Dim datat As String
    sckServer(Index).GetData datat
    datat = Replace(datat, "  ", " ")
    Data = Split(datat, " ")
    
    If Debugmode Then Log ("<<<" + "(" + Str(Index) + ")" + datat)
    
    
    If Data(0) = "login" Then
    '��½
        
        Log ("(" + Str(Index) + ")" + Data(2) + " ���ڳ��Ե�½" + "...")
        
        If Int(Data(1)) = version Then
        '�汾��ȷ
        
            zh = Data(2)
            mm = Data(3)
            zhmm = zh + "," + mm
            
            Dim nowzhmm As String
            nowzhmm = ""
            
            Open App.Path & txtname For Input As #1
            
            Do While zhmm <> nowzhmm
                If EOF(1) = True Then
                    Log ("(" + Str(Index) + ")" + zh + " �˺��������!")
                    Dim cw As String
                    sckServer(Index).SendData "�˺��������!"
                    'sckServer(Index).Close
                    Exit Do
                Else
                    Line Input #1, savedata
                    AAAAA = Split(savedata, ",")
                    If UBound(AAAAA) >= 1 Then nowzhmm = AAAAA(0) + "," + AAAAA(1)
                End If
            Loop
            
            If zhmm = nowzhmm Then
            
                If UserIsLogin.Exists(zh) Then
            
                    sckServer(Index).SendData "�ظ���½!"
                    Log ("(" + Str(Index) + ")" + zh + " �ظ���½!")
            
                Else
                
                    sckServer(Index).SendData "��½�ɹ�! "
                    Log ("(" + Str(Index) + ")" + zh + " ��½�ɹ�!")
                    IndexName.Item(Index) = zh
                    UserIsLogin.Item(zh) = "True"
                    IndexIfLogin.Item(Index) = "True"
                    IndexRoom.Item(Index) = "0"
                    
                    sxOnPlays
                
                End If
                
                
                
            End If
            
            Close #1
            
        Else
        '�汾����ȷ
            Log ("(" + Str(Index) + ")" + Data(2) + " ��½ʧ�ܣ��汾��һ�£��ͻ��˰汾Ϊ" + Data(1))
            sckServer(Index).SendData "loginfail" + " " + Str(version)
            'sckServer(Index).Close
        End If
    
    End If
    
    
    
    If Data(0) = "sign" Then
    'ע��
        
        If Data(1) = "username" Then
        '����û����Ƿ����
        
            Open App.Path & txtname For Input As #1
                Dim aa As String
                aa = ""
                cw = ""
                
                Do While Data(2) <> aa
                    If EOF(1) = True Then
                        
                        cw = "*�û�������*"
                        
                        Exit Do
                    Else
                        Line Input #1, aa
                        xx = Split(aa, ",")
                        If UBound(xx) >= 1 Then aa = xx(0)
                    End If
                Loop
                
            Close #1
            
            If Data(2) = aa Then
                cw = "�û���������"
            End If
            
            sckServer(Index).SendData cw
            If Debugmode Then Log (">>>" + "(" + Str(Index) + ")" + cw)
            
        End If
        
        
        
        If Data(1) = "up" Then
        'д��ע����Ϣ
            
            
            inyzm = Data(3)
            ema = Split(Data(2), "###")(3)
            
            If Int(inyzm) = Emayzm.Item(ema) Then
                
                Emayzm.Remove ema
                
                
                x = Data(2)
                
                Open App.Path & txtname For Input As #1
        
                    aa = ""
                    xxx = Split(x, "###")
                    xa = xxx(3)
                    cw = ""
                    
                    Do While xa <> aa
                        If EOF(1) = True Then
                            
                            cw = " ע��ɹ�!"
                            
                            Exit Do
                        Else
                            Line Input #1, aa
                            xx = Split(aa, ",")
                            If UBound(xx) >= 1 Then aa = xx(2)
                            
                            If xa = aa Then
                                cw = "�������ѱ���!"
                                Exit Do
                            End If
                            
                        End If
                    Loop
                Close #1
                
                If cw = "ע��ɹ�!" Then
                        
                        Open App.Path & txtname For Input As #1
                        sxx = "0"
                        Do Until EOF(1)
                            Line Input #1, Ax
                            If sxx = "0" Then
                                Sx = Sx + Ax
                                sxx = "1"
                            Else
                                Sx = Sx + vbNewLine + Ax
                            End If
                            If EOF(1) = True Then
                                Exit Do
                            End If
                        Loop
                        Close
                        
                        xxxxxx = Split(x, "###")
                        x = xxxxxx(1) + "," + xxxxxx(2) + "," + xxxxxx(3)
                        Sx = Sx + vbNewLine + x
                        
                        Open App.Path & txtname For Output As #1
                        Print #1, Sx
                        Close
                        
                        NameMoney.Item(xxxxxx(1)) = 0
                        
                        Text1.Text = Text1.Text + "[" + Time$ + "]" + xxxxxx(1) + "/" + xxxxxx(3) + "ע��ɹ�!" + vbCrLf
                        Text1.SelStart = Len(Text1.Text)
                        
                End If
            
                sckServer(Index).SendData cw
                If Debugmode Then Log (">>>" + "(" + Str(Index) + ")" + cw)
                
            Else
            
                sckServer(Index).SendData "��֤�����!"
                If Debugmode Then Log ("(" + Str(Index) + ")" + "��֤�����!")

            End If
            
        End If
        
        
        
        If Data(1) = "ema" Then
            
            Dim i As Integer
            i = Len(Text1.Text) + Int(Mid(Time$, 7, 2))
            
            For l = 1 To i
                yzm = Int(1000000 * Rnd)
            Next
            
            Emayzm.Item(Data(2)) = yzm
            
            NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
    
            Set Email = CreateObject("CDO.Message")
            Email.From = "TDR_Group@foxmail.com" '��������ַ
            Email.To = Data(2) 'Ҫ�����ĵ�ַ
            Email.Subject = "�����˺�ע��-������֤" '�ʼ�����
            
            Dim tbody As String
            
            tbody = "��л��ע������˺�!������ע�����֤��Ϊ " + Str(yzm) + " ,�뾡�����ע��!"
            
            Email.Textbody = tbody '�ʼ�����
            
            'Email.AddAttachment "c:\1.txt" ?'������������Ӷ�� ?�����������������·�� ���� "c:\1.txt" ?�������Ҫ��������ɾ������
            
            With Email.Configuration.Fields
                .Item(NameSpace & "sendusing") = 2
                .Item(NameSpace & "smtpserver") = "smtp.qq.com" '�����ʼ�������
                .Item(NameSpace & "smtpserverport") = 25
                .Item(NameSpace & "smtpauthenticate") = 1
                .Item(NameSpace & "sendusername") = "TDR_Group" '�����û���
                .Item(NameSpace & "sendpassword") = "awcitnitvfgvbhci" '��Ȩ��
                .Update
            End With
            
            Email.Send
            
            If Debugmode Then Log ("(" + Str(Index) + ")" + "�����ʼ���" + Data(2) + " ��֤�룺" + Str(yzm))
            sckServer(Index).SendData "sand yzm sucess"
        
        End If
    
    End If
    
    
    
    If Data(0) = "selfinfo" And UserIsLogin(IndexName(Index)) = "True" Then
    '��ȡ��Ϣ
    
        zh = IndexName(Index)
        Money = NameMoney(zh)
        Dim onlineplayerscount As Integer
        onlineplayerscount = IndexName.Count
        
        Dim message As String
        message = "selfinfo " + zh + " " + Str(Money) + " " + Str(onlineplayerscount) + " "
        sckServer(Index).SendData message
        If Debugmode Then Log (">>>" + "(" + Str(Index) + ")" + message)
        
    End If



    If Data(0) = "test" And UserIsLogin(IndexName(Index)) = "True" Then
    '����
        
        If Data(1) = "moneyadd1" Then
            
            NameMoney.Item(IndexName(Index)) = NameMoney.Item(IndexName(Index)) + 1
            
            zh = IndexName(Index)
            Money = NameMoney(zh)
            Dim onlineplayerscount2 As Integer
            onlineplayerscount2 = IndexName.Count
            
            Dim message2 As String
            message2 = "selfinfo " + zh + " " + Str(Money) + " " + Str(onlineplayerscount2)
            sckServer(Index).SendData message2
            If Debugmode Then Log (">>>" + "(" + Str(Index) + ")" + message2)
            
        End If
        
    End If
    
    
    
    
    If Data(0) = "room" And UserIsLogin(IndexName(Index)) = "True" Then
    '����
        
        If Data(1) = "create" Then
        '��������
            If NameMoney.Item(IndexName(Index)) >= 100 Then
            
                NameMoney.Item(IndexName(Index)) = NameMoney.Item(IndexName(Index)) - 100
                Log ("(" + Str(Index) + ")" + "�������䣺" + Data(2))
                IndexRoom.Item(Index) = Data(2)
                sckServer(Index).SendData "CreateRoomSucess " + Data(2) + " "
                sxRooms
                
            Else
                sckServer(Index).SendData "tip ��Ҳ���"
                Log ("(" + Str(Index) + ")" + IndexName.Item(Index) + "�쳣����Ҳ���ȴ���󴴽�����")
            End If
            
            
            
        End If
        
        If Data(1) = "r" Then
        'ˢ�·���
            sxRooms
        End If
        
        If Data(1) = "join" Then
        '���뷿��
            tRoom = Data(2)                     ' 2023-10-20 �������������Ҫ��ɼ��뷿�䲿��
            
            
            
            
            
            
            
            
            
        End If
        
    End If









End Sub



Private Sub Form_Unload(Cancel As Integer)

    SaveLog

End Sub



Private Sub TPort_Change()

    server_port = TPort.Text

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim i As Long
    For i = 0 To WsIsBusy.Count - 1
        If sckServer(i).State <> sckClosed Then
            sckServer(i).Close
        End If
    Next
    
End Sub



Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)

    Dim SockIndex As Long
    Dim SockNum As Long
    
    On Error Resume Next
    'Log (requestID & " ��������")
    '�������ӵ��û���
    SockNum = WsIsBusy.Count

    '���ҿ��е�sock
    SockIndex = FindFreeSocket()
    '������е�sock��æ����̬���sock
    If SockIndex >= SockNum Then
        Load sckServer(SockIndex)
    End If
    WsIsBusy.Item(SockIndex) = "True"
    'sckServer(SockIndex).Tag = SockIndex

    '��������
    sckServer(SockIndex).Accept requestID
    'Log (SockIndex & " ��������")
    
End Sub



Private Sub sckServer_Close(Index As Integer)

    If sckServer(Index).State <> sckClosed Then
        sckServer(Index).Close
    End If
    
    WsIsBusy.Item(Index) = "False"
    
    If UserIsLogin.Exists(IndexName.Item(Index)) Then UserIsLogin.Remove (IndexName.Item(Index))
    
    If IndexName.Exists(Index) Then
        Log ("(" + Str(Index) + ")" + IndexName.Item(Index) + " �Ͽ�����")
        IndexName.Remove Index
    End If
    
    If IndexRoom.Exists(Index) Then IndexRoom.Remove (Index)
    
    If IndexIfLogin.Exists(Index) Then IndexIfLogin.Remove (Index)
    
    sxOnPlays
    sxRooms
    
End Sub



Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    Log "����δ֪����!����������!"
    Winsock1.Close
    CStart.Enabled = True
    CStop.Enabled = False

End Sub



Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    If sckServer(Index).State <> sckClosed Then
        sckServer(Index).Close
    End If
    
End Sub



Private Sub CDebug_Click()

    If Debugmode Then
        Debugmode = False
    Else
        Debugmode = True
    End If
    CDebug.Caption = "DebugMode:" + Str(Debugmode)

End Sub



'Ѱ�ҿ��е�socket
Public Function FindFreeSocket() As Long

    Dim SockCount As Long, i As Long
    SockCount = WsIsBusy.Count - 1
    
    For i = 0 To SockCount
        If WsIsBusy.Item(i) = "False" Then
            FindFreeSocket = i
            Exit Function
        End If
    Next

    FindFreeSocket = WsIsBusy.Count
    
End Function



Public Function Log(mmm)

    Text1.Text = Text1.Text + "[" + Time$ + "]" + mmm + vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
End Function



Public Function sxOnPlays()

    Dim allplays
    allplays = IndexName.keys
    List1.Clear
    For i = 0 To IndexName.Count - 1
        List1.AddItem IndexName.Item(allplays(i))
    Next
    For i = 0 To sckServer.Count - 1
        If sckServer(i).State <> sckClosed Then
            sckServer(i).SendData "reflash"
            'sckServer(i).Close
        End If
    Next

End Function



Public Function sxRooms()

    Dim allrooms
    Dim roommess
    Dim roomsc As Integer
    roommess = ""
    allrooms = IndexRoom.items
    List2.Clear
    For i = 0 To IndexRoom.Count - 1
        If Not allrooms(i) = 0 Then
            roomsc = roomsc + 1
            List2.AddItem allrooms(i)
            roommess = roommess + allrooms(i) + "###"
        End If
    Next
    roommess = "nowrooms " + Str(roomsc) + " " + roommess
    
    For i = 0 To sckServer.Count - 1
        If sckServer(i).State <> sckClosed Then
            sckServer(i).SendData roommess
            'sckServer(i).Close
        End If
    Next

End Function



