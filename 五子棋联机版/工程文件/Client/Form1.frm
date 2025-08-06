VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "五子棋"
   ClientHeight    =   11985
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   19770
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11985
   ScaleWidth      =   19770
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   270
      Left            =   12720
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   6000
      Width           =   700
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      ForeColor       =   &H80000008&
      Height          =   10500
      Left            =   0
      ScaleHeight     =   10470
      ScaleWidth      =   10470
      TabIndex        =   0
      Top             =   0
      Width           =   10500
      Begin VB.CommandButton Cmd3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   960
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3360
         Width           =   500
      End
      Begin VB.CommandButton Cmd2 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   0
         Left            =   360
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2280
         Width           =   500
      End
      Begin VB.CommandButton Cmd1 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   0
         Left            =   360
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1320
         Width           =   500
      End
      Begin VB.CommandButton Cell 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   500
      End
   End
   Begin VB.Menu mnuStartGame 
      Caption         =   "开始游戏"
      Begin VB.Menu mnuSingle 
         Caption         =   "单机游戏"
      End
      Begin VB.Menu mnuDouble 
         Caption         =   "联机游戏"
      End
   End
   Begin VB.Menu mnuSetting 
      Caption         =   "游戏设置"
      Begin VB.Menu mnuRestart 
         Caption         =   "重新开始"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出当前游戏"
      End
      Begin VB.Menu mnuResize 
         Caption         =   "重置窗口大小"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "关于"
      End
   End
   Begin VB.Menu mnuServer 
      Caption         =   "联机设置"
      Begin VB.Menu mnuConnect 
         Caption         =   "连接服务器..."
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "断开服务器连接"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "我的信息"
      Begin VB.Menu mnuMyIP 
         Caption         =   "我的IP"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMyID 
         Caption         =   "我的ID"
      End
      Begin VB.Menu mnuMyPC 
         Caption         =   "我的主机名称"
      End
      Begin VB.Menu mnuLoginTime 
         Caption         =   "我的登录时间"
      End
   End
   Begin VB.Menu mnuState 
      Caption         =   "状态"
   End
   Begin VB.Menu mnuLocation 
      Caption         =   "上一个棋落在了："
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'点击按钮下棋
Private Sub Cell_Click(index As Integer)
    If Gaming = True Then
        If OnlineGame = False Then
            Cell(index).Enabled = False
            If FlagA Then
                Cell(index).Caption = "●"
                Cmd3.Caption = "o"
            Else
                Cell(index).Caption = "o"
                Cmd3.Caption = "●"
            End If
            FlagA = Not FlagA
            '算法检测
            Call Module1.Detection(index)
        Else
            If (MyColor = "●" And FlagB = True) Or (MyColor = "o" And FlagB = False) Then
                Cell(index).Enabled = False
                Cell(index).Caption = MyColor
                Cmd3.Caption = Competitor_Color
                FlagB = Not FlagB
                Scnn.wskClient.SendData "Location" & Competitor_Index & Trim(Str(index))
                DoEvents
                Call Module1.Detection(index)
            Else
                MsgBox "请等待对方落棋！", vbOKOnly, "提示"
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.mnuDisconnect.Enabled = False
    Me.mnuInfo.Enabled = False
    Me.mnuExit.Enabled = False
    Me.mnuRestart.Enabled = False
    Me.mnuLocation.Visible = False
    Me.mnuState.Caption = "离线"
    
    Cell(0).Visible = False
    Cmd1(0).Visible = False
    Cmd2(0).Visible = False
    Cmd3.Visible = False
    Me.Height = 11385
    Me.Width = 10745
    
    Text1.Text = Info1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Gaming And Competitor_IP <> "" Then
        Scnn.wskClient.SendData "对方已离线" & Competitor_Index
        DoEvents
    End If
    End
End Sub

Private Sub mnuAbout_Click()
    MsgBox Info, vbOKOnly, "关于"
End Sub

Private Sub mnuConnect_Click()
    Scnn.Show
End Sub

Private Sub mnuDisconnect_Click()
    Dim a As Variant
    a = MsgBox("确定要断开服务器连接吗？", vbYesNo, "提示")
    If a = vbYes Then
        If Gaming And Competitor_IP <> "" Then
            Scnn.wskClient.SendData "对方已离线" & Competitor_Index
            DoEvents
            Call Form1.UnloadChessboard
        End If
        Scnn.wskClient.Close
        OnlineGame = False
        Me.mnuState.Caption = "离线"
        Me.mnuInfo.Enabled = False
        Me.mnuConnect.Enabled = True
        Me.mnuDisconnect.Enabled = False
    End If
End Sub

Private Sub mnuExit_Click()
    If MsgBox("确定要退出当前游戏吗？", vbYesNo, "提示") = vbYes Then
        If OnlineGame And Competitor_IP <> "" Then
            Scnn.wskClient.SendData "对方已退出对战" & Competitor_Index
            DoEvents
            Competitor_IP = ""
            Competitor_Index = ""
        End If
        Call UnloadChessboard
    End If
    Call MenuInit
End Sub

Private Sub mnuLoginTime_Click()
    MsgBox LoginTime, vbOKOnly, "我的登录时间"
End Sub

Private Sub mnuMyID_Click()
    MsgBox MyIndex, vbOKOnly, "我的ID"
End Sub

Private Sub mnuMyIP_Click()
    MsgBox Scnn.wskClient.LocalIP, vbOKOnly, "我的IP地址"
End Sub

Private Sub mnuMyPC_Click()
    MsgBox Scnn.wskClient.LocalHostName, vbOKOnly, "我的主机名称"
End Sub

Private Sub mnuDouble_Click()
    '邀请玩家
    If Scnn.wskClient.State <> 7 Then
        MsgBox "请先连接服务器！", vbOKOnly, "提示"
    Else
        Competitor_Index = Trim(InputBox("输入被邀请玩家的ID", "邀请玩家"))
        If Competitor_Index = MyIndex Then
            MsgBox "不能邀请自己！", vbOKOnly, "提示"
        Else
            Scnn.wskClient.SendData "五子棋邀请:" & Competitor_Index
            Me.mnuState.Visible = True
            Me.mnuState.Caption = "邀请玩家" & Competitor_Index & "中"
            Me.mnuSingle.Enabled = False
            Me.mnuDouble.Enabled = False
        End If
    End If
End Sub

Private Sub mnuResize_Click()
    If Me.WindowState = 2 Then Me.WindowState = 0
    Me.Height = 11385
    Me.Width = 10745
    Text1.Text = Info1
End Sub

'重新开始重新生成棋盘
Private Sub mnuRestart_Click()
    If OnlineGame Then Scnn.wskClient.SendData "五子棋重新开始" & Competitor_Index
    Call LoadChessboard
End Sub

Private Sub mnuSingle_Click()
    Me.mnuSingle.Enabled = False
    Me.mnuDouble.Enabled = False
    Call LoadChessboard
End Sub

'加载棋盘
Public Sub LoadChessboard()
    '如果棋盘已存在则卸载棋盘
    If Me.Cell.UBound > 0 Then
        Call UnloadChessboard
    End If
    
    '生成棋盘
    Dim i As Integer
    For i = 1 To 400
        DoEvents
        Load Cell(i)
        Set Cell(i).Container = Form1.Picture1
        Cell(i).Visible = True
        
        If i < 21 Then
                Load Cmd1(i)
                Load Cmd2(i)
                Set Cmd1(i).Container = Me.Picture1
                Set Cmd2(i).Container = Me.Picture1
                Cmd1(i).Caption = i
                Cmd2(i).Caption = Chr(64 + i)
                Cmd1(i).Visible = True
                Cmd2(i).Visible = True
            If i = 1 Then
                Cell(i).Top = 0
                Cell(i).Left = 0
                
                Cmd1(i).Top = 0
                Cmd1(i).Left = 20 * Cell(i).Width
                Cmd2(i).Top = 20 * Cell(i).Height
                Cmd2(i).Left = 0
            Else
                Cell(i).Top = Cell(1).Top
                Cell(i).Left = Cell(i - 1).Left + Cell(0).Width
                
                Cmd1(i).Top = Cmd1(i - 1).Top + 500
                Cmd1(i).Left = Cmd1(1).Left
                Cmd2(i).Top = Cmd2(1).Top
                Cmd2(i).Left = Cmd2(i - 1).Left + 500
            End If
            If i = 20 Then
                Cmd3.Left = 10000
                Cmd3.Top = 10000
                Cmd3.Visible = True
            End If
        Else
            Cell(i).Left = Cell(i - 20).Left
            Cell(i).Top = Cell(i - 20).Top + Cell(0).Height
        End If
        'Cell(i).Caption = i
    Next i
    Gaming = True
    Me.mnuRestart.Enabled = True
    Me.mnuExit.Enabled = True
    Me.mnuLocation.Visible = True
    Me.mnuLocation.Caption = "上一个棋落在了："
    
    '待优化：索引号大的先手
    If OnlineGame And Competitor_IP <> "" Then
        If MyIndex > Competitor_Index Then
            FlagB = True
            MyColor = "●"
            Competitor_Color = "o"
            MsgBox "你是黑棋●先手", vbOKOnly, "提示"
        Else
            FlagB = True
            MyColor = "o"
            Competitor_Color = "●"
            MsgBox "你是白棋o后手", vbOKOnly, "提示"
        End If
        Me.mnuState.Caption = "与" & Competitor_IP & "对战中"
    Else
        FlagA = True
    End If
    Cmd3.Caption = "●"
End Sub

'卸载棋盘
Public Sub UnloadChessboard()
    Gaming = False
    Erase Memory
    Me.mnuSingle.Enabled = False
    Me.mnuDouble.Enabled = False
    Me.mnuRestart.Enabled = False
    Me.mnuExit.Enabled = False
    Me.mnuLocation.Visible = False
    
    Dim i As Integer
    For i = 1 To 400
        DoEvents
        If i < 21 Then
            Unload Me.Cmd1(i)
            Unload Me.Cmd2(i)
            Me.Cmd3.Caption = ""
            Me.Cmd3.Visible = False
        End If
        Unload Me.Cell(i)
    Next i
End Sub

'菜单初始化
Public Sub MenuInit()
    Me.mnuSingle.Enabled = True
    Me.mnuDouble.Enabled = True
    Me.mnuRestart.Enabled = False
    Me.mnuExit.Enabled = False
    Me.mnuLocation.Visible = False
    If OnlineGame Then
        Me.mnuState.Caption = " 在线"
    Else
        Me.mnuState.Caption = "离线"
    End If
End Sub

