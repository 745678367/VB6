VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Server"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   19755
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   19755
   StartUpPosition =   1  '所有者中心
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   7950
      ItemData        =   "Server.frx":0442
      Left            =   4440
      List            =   "Server.frx":0444
      TabIndex        =   1
      Top             =   0
      Width           =   3000
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Height          =   7950
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
   Begin MSWinsockLib.Winsock wskListen 
      Left            =   8040
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskServer 
      Index           =   0
      Left            =   8160
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuServerSetting 
      Caption         =   "服务器设置"
      Begin VB.Menu mnuServerIP 
         Caption         =   "服务器IP"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenServer 
         Caption         =   "开启服务器"
      End
      Begin VB.Menu mnuCloseServer 
         Caption         =   "关闭服务器"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResize 
         Caption         =   "还原窗口"
      End
   End
   Begin VB.Menu mnuServerInfo 
      Caption         =   "服务器信息"
      Begin VB.Menu mnuServerState 
         Caption         =   "服务器状态"
      End
      Begin VB.Menu mnuClientQuantity 
         Caption         =   "当前连接数量"
      End
      Begin VB.Menu mnuClearLogs 
         Caption         =   "清空服务器日志"
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "关于"
      End
   End
   Begin VB.Menu mnuBroadcast 
      Caption         =   "广播信息"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    '服务器数据库初始化
    ServerIP = Me.wskListen.LocalIP
    ConnectAccess2.ConnectAccess DBChesscnn, DBChessset
    With DBChessset
        .Open "delete * from 五子棋", DBChesscnn
    End With
    With Me
        .Height = 7000
        .Width = 8000
        .Text4.Text = "服务器日志"
        .mnuCloseServer.Enabled = False
        .mnuServerIP.Caption = .wskListen.LocalIP
        .mnuServerState.Caption = "未开启"
        .mnuClientQuantity.Caption = "当前连接数量：0"
        .mnuBroadcast.Enabled = False
    End With
    With Me.Text4
        .Top = 0
        .Left = 0
        .Width = 5000
        .Height = Me.ScaleHeight
    End With
    With Me.List1
        .Top = 0
        .Left = Me.Text4.Width
        .Width = Me.ScaleWidth - Me.Text4.Width
        .Height = Me.ScaleHeight + 30
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    '发生错误继续执行
    On Error Resume Next
    For i = wskServer.UBound To 1 Step -1
        wskServer(i).Close
        Unload wskServer(i)
    Next
    wskListen.Close
    Module2.ClearDB
    End
End Sub

Private Sub mnuAbout_Click()
    MsgBox Info, vbOKOnly, "关于"
End Sub

'广播消息
Private Sub mnuBroadcast_Click()
    Dim message As String, i As Integer
    message = InputBox("输入要广播的信息", "服务器广播")
    If Len(Trim(message)) > 0 Then
        For i = 1 To Me.wskServer.UBound
            If Me.wskServer(i).State = 7 Then Me.wskServer(i).SendData "广播" & message
            DoEvents
        Next i
        Me.Text4.Text = Text4 & vbCrLf & Format$(Now, "General Date") & Space(2) & "服务器广播消息：" & message
        MsgBox message, vbInformation, "广播完成"
    End If
End Sub

Private Sub mnuClearLogs_Click()
    Me.Text4.Text = "服务器日志"
End Sub

Private Sub mnuCloseServer_Click()
    Dim i As Integer
    
    wskListen.Close
    For i = wskServer.UBound To 1 Step -1
        If wskServer(i).State <> 0 Then wskServer(i).Close
        Unload wskServer(i)
    Next i
    With Me
        .mnuOpenServer.Enabled = True
        .mnuCloseServer.Enabled = False
        .Text4.Text = Text4 & vbCrLf & Format$(Now, "General Date") & Space(2) & "服务器已关闭！"
        .mnuServerState.Caption = "未开启"
        .mnuBroadcast.Enabled = False
    End With
    ClientQuantities = 0
    Me.mnuClientQuantity.Caption = "当前连接数量：0"
    Module2.ClearDB
End Sub

Private Sub mnuOpenServer_Click()
try:
    On Error GoTo catch
    '绑定到设置的服务器IP的9999端口(127.0.0.1为本机地址)
    wskListen.Bind 9999, ServerIP
    DoEvents
    '监听网络连接
    wskListen.Listen
    With Me
        .Text4 = Text4 & vbCrLf & Format$(Now, "General Date") & Space(2) & "服务器已开启！"
        .mnuOpenServer.Enabled = False
        .mnuCloseServer.Enabled = True
        .mnuServerState.Caption = "已开启"
        .mnuBroadcast.Enabled = True
    End With
    Exit Sub
catch:
    MsgBox "设置了一个无效的IP地址！", vbInformation, "错误"
    Exit Sub
End Sub

Private Sub mnuResize_Click()
    If Me.WindowState = 2 Then Me.WindowState = 0
    Me.Height = 7000
    Me.Width = 8000
End Sub

Private Sub mnuServerIP_Click()
    Dim a As String
    a = InputBox("输入要设置的服务器IP", "服务器IP设置", ServerIP)
    If Len(Trim(a)) <> 0 Then
        ServerIP = a
        Me.mnuServerIP.Caption = ServerIP
    End If
End Sub

Private Sub Text4_Change()
    Text4.SelStart = Len(Text4)
End Sub

Private Sub wskListen_ConnectionRequest(ByVal requestID As Long)
    Dim i As Long
    
    '为新连接请求查找当前空闲服务器
    For i = 1 To wskServer.UBound
        If wskServer(i).State = sckClosed Then
            wskServer(i).Accept requestID
            Text4 = Text4 & vbCrLf & Format$(Time, "General Date") & Space(1) & "[" & wskListen.RemoteHostIP & "]" & Space(1) & "已成功连接"
            List1.AddItem wskListen.RemoteHostIP
            ClientQuantities = ClientQuantities + 1
            Me.mnuClientQuantity.Caption = "当前连接数量：" & ClientQuantities
            Exit Sub
        End If
    Next
    
    '若当前无空闲服务器，则新加载一个
    Load wskServer(i)
    wskServer(i).Accept requestID
    Text4 = Text4 & vbCrLf & Format$(Time, "General Date") & Space(1) & "[" & wskListen.RemoteHostIP & "]" & Space(1) & "已成功连接"
    List1.AddItem wskListen.RemoteHostIP
    ClientQuantities = ClientQuantities + 1
    Me.mnuClientQuantity.Caption = "当前连接数量：" & ClientQuantities
End Sub

Private Sub wskServer_Close(index As Integer)
    Dim i As Integer
    
    Me.wskServer(index).Close
    Text4 = Text4 & vbCrLf & Format$(Time, "General Date") & Space(1) & "[" & wskListen.RemoteHostIP & "]" & Space(1) & "已断开连接"
    For i = 0 To List1.ListCount - 1
        If List1.List(i) = wskServer(index).RemoteHostIP Then
            List1.RemoveItem i
            Exit For
        End If
    Next i
    
    ConnectAccess2.RefreshDB DBChesscnn
    With DBChessset
        If .State = 1 Then .Close
        .Open "delete * from 五子棋 where PlayerIndex='" & Trim(Str(index)) & "'", DBChesscnn
    End With
    ClientQuantities = ClientQuantities - 1
    Me.mnuClientQuantity.Caption = "当前连接数量：" & ClientQuantities
End Sub

Private Sub wskServer_DataArrival(index As Integer, ByVal bytesTotal As Long)
    Dim strdata As String
    wskServer(index).GetData strdata, vbString
    
    '用户连接
    If strdata = "!@#五子棋" Then
        Dim playerip As String, logintime As String
        
        playerip = Me.wskServer(index).RemoteHostIP
        logintime = Format$(Now, "General Date")
        Module2.ChessID index, playerip, logintime
    End If
    
    '用户邀请玩家
    If InStr(strdata, "五子棋邀请:") = 1 Then
        Dim guest_index As String
        
        guest_index = Trim(Mid(strdata, 7, 2))
        Module2.Invite Trim(Str(index)), guest_index
    End If
    
    '接受邀请
    If InStr(strdata, "五子棋接受邀请") = 1 Then
        Dim host_index1 As String
        host_index1 = Trim(Mid(strdata, 8, 2))
        Module2.Accept host_index1, Trim(Str(index))
    End If
    
    '拒绝邀请
    If InStr(strdata, "五子棋拒绝邀请") = 1 Then
        Module2.Refuse Mid(strdata, 8, 2), Trim(Str(index))
    End If
    
    '转发落棋的位置
    If InStr(strdata, "Location") = 1 Then
        Dim loc As String, index1 As Integer
        index1 = Val(Trim(Mid(strdata, 9, 2)))
        loc = Mid(strdata, 11)
        Me.wskServer(index1).SendData "Location" & loc
        DoEvents
    End If
    
    '重新开始
    If InStr(strdata, "五子棋重新开始") = 1 Then
        Dim index2 As Integer
        index2 = Val(Trim(Mid(strdata, 8)))
        Me.wskServer(index2).SendData "五子棋重新开始"
    End If
    
    '退出对战
    If InStr(strdata, "对方已退出对战") = 1 Then
        Dim index3 As Integer
        index3 = Val(Trim(Mid(strdata, 8)))
        Module2.ExitCompetition index, Trim(Str(index3))
        Me.wskServer(index3).SendData "对方已退出对战"
    End If
    
    '对方离线
    If InStr(strdata, "对方已离线") = 1 Then
        Dim index4 As Integer
        index4 = Val(Trim(Mid(strdata, 6, 2)))
        Module2.ExitCompetition index, Trim(Str(index4))
        Me.wskServer(index4).SendData "对方已离线"
    End If
End Sub
