VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Scnn 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "连接服务器"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   3405
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Scnn.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   3405
   StartUpPosition =   1  '所有者中心
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1440
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   600
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuServerIP 
      Caption         =   "配置服务器..."
      NegotiatePosition=   1  'Left
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuConn 
      Caption         =   "连接"
   End
   Begin VB.Menu mnuSIP 
      Caption         =   "服务器"
   End
End
Attribute VB_Name = "Scnn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    If OnlineGame = False Then
        Cls
        Print "服务器未连接", Time
    End If
End Sub

Private Sub Form_Load()
    Me.mnuSIP.Caption = ServerIP
    Cls
    Print "服务器未连接", Time
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.wskClient.Close
End Sub

Private Sub mnuConn_Click()
    If Len(Trim(ServerIP)) = 0 Then
        MsgBox "请配置服务器IP", vbOKOnly, "提示"
    Else
try:  On Error GoTo catch
        Cls
        Print "服务器IP地址:", ServerIP
        Print "服务器连接中", Time
        If Me.wskClient.State <> 0 Then Me.wskClient.Close
        Me.wskClient.Connect ServerIP, 9999
    End If
    Exit Sub
catch:
    MsgBox "配置了一个无效的IP地址", vbInformation, "错误"
    Exit Sub
End Sub

Private Sub mnuServerIP_Click()
     ServerIP = InputBox("请输入服务器IP地址", "连接服务器", ServerIP)
     Me.mnuSIP.Caption = ServerIP
End Sub

Private Sub wskClient_Close()
    wskClient.Close
    MsgBox "已与服务器断开连接！", vbInformation, "提示"
    Print "服务器未连接", Time
    OnlineGame = False
    With Form1
        .mnuState.Caption = "离线"
        .mnuInfo.Enabled = False
        .mnuDisconnect.Enabled = False
        .mnuConnect.Enabled = True
        .mnuLocation.Visible = False
    End With
    If Gaming = True Then
        Gaming = False
        Call Form1.UnloadChessboard
    End If
    '自动尝试重新连接
    'wskClient.Connect ServerIP, 9999
End Sub

Private Sub wskClient_Connect()
    OnlineGame = True
    MyIP = Me.wskClient.LocalIP
    LoginTime = Format$(Now, "General Date")
    Print "服务器已连接！", Time
    Form1.Show
    Me.Hide
    
    With Form1
        .mnuMyIP.Caption = MyIP
        .mnuState.Caption = "在线"
        .mnuInfo.Enabled = True
        .mnuConnect.Enabled = False
        .mnuDisconnect.Enabled = True
    End With
    
    Me.wskClient.SendData "!@#五子棋"
End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Me.wskClient.GetData strData, vbString
    
    If InStr(strData, "五子棋Index") = 1 Then
        MyIndex = Mid(strData, 9)
    End If
    
    If InStr(strData, "五子棋邀请失败:") = 1 Then
        With Form1
            .mnuState.Caption = "在线"
            .mnuSingle.Enabled = True
            .mnuDouble.Enabled = True
        End With
        MsgBox Mid(strData, 9), vbOKOnly, "邀请失败"
    End If
    
    If InStr(strData, "五子棋拒绝邀请") = 1 Then
        Competitor_Index = ""
        With Form1
            .mnuState.Caption = "在线"
            .mnuSingle.Enabled = True
            .mnuDouble.Enabled = True
        End With
        MsgBox "ID：" & Mid(strData, 8, 2) & "拒绝了您的邀请", vbOKOnly, "邀请拒绝"
    End If
    
    If InStr(strData, "五子棋邀请对战:") = 1 Then
        Dim a As Variant
        Competitor_Index = Mid(strData, 9, 2)
        Competitor_IP = Mid(strData, 11)
        a = MsgBox("ID：" & Competitor_Index & vbCrLf & "IP：" & Competitor_IP & vbCrLf & "邀请您对战", vbYesNo, "邀请对战")
        If a = vbYes Then
            Me.wskClient.SendData "五子棋接受邀请" & Competitor_Index
        Else
            Me.wskClient.SendData "五子棋拒绝邀请" & Competitor_Index
        End If
    End If
    
    If InStr(strData, "五子棋联机对战:开始") = 1 Then
        Competitor_IP = Mid(strData, 11)
        OnlineGame = True
        Call Form1.LoadChessboard
    End If
        
    If InStr(strData, "Location") = 1 Then
        Dim index1 As Integer
        index1 = Val(Mid(strData, 9))
        Module1.UpdateChessboard index1
    End If
    
    If InStr(strData, "五子棋重新开始") = 1 Then
        Call Form1.LoadChessboard
    End If
    
    If InStr(strData, "对方已退出对战") = 1 Then
        MsgBox Competitor_Index & vbCrLf & Competitor_IP & "已退出对战", vbOKOnly, "对战结束"
        Call Form1.UnloadChessboard
        Call Form1.MenuInit
        Form1.mnuState.Caption = "在线"
        Competitor_Index = ""
        Competitor_IP = ""
    End If
    
    If InStr(strData, "广播") = 1 Then
        Dim message As String
        message = Mid(strData, 3)
        MsgBox message, vbInformation, "服务器广播"
    End If
    
    If InStr(strData, "对方已离线") = 1 Then
        MsgBox Competitor_Index & vbCrLf & Competitor_IP & "已离线", vbOKOnly, "对战结束"
        Call Form1.UnloadChessboard
        Call Form1.MenuInit
        Form1.mnuState.Caption = "在线"
        Competitor_Index = ""
        Competitor_IP = ""
    End If
End Sub

Private Sub wskClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Print Description, Time
    Me.wskClient.Close
End Sub
