VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "文件传输助手"
   ClientHeight    =   10530
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   13620
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   13620
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8880
      Top             =   2760
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000003&
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   370
      Left            =   4560
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2640
      Width           =   2890
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   850
      Left            =   2760
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   3600
      Width           =   1570
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3120
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000F&
      Height          =   4740
      Left            =   2160
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   8040
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   370
      Left            =   1200
      TabIndex        =   9
      Top             =   5280
      Width           =   610
   End
   Begin VB.Menu mnuSetting 
      Caption         =   "设置"
      Begin VB.Menu mnuDelete 
         Caption         =   "删除"
      End
      Begin VB.Menu mnuSend 
         Caption         =   "发送"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "文件保存路径..."
      End
      Begin VB.Menu mnu0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "关于..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With Me
        .Width = 5000
        .Height = 6800
        
        .Text1.Move 0, 0, .ScaleWidth / 2, 300
        .Text2.Move Text1.Width, 0, .ScaleWidth - Text1.Width, 300
        .Text3.Move 0, 300, .ScaleWidth / 2, 300
        .Text4.Move Text3.Width, 300, .ScaleWidth / 2, 300
        .List1.Move 0, 600, .ScaleWidth, 5000
        .Text5.Move 0, 5500, .ScaleWidth, 300
        .Text6.Move 0, 5800, .ScaleWidth / 2, 300
        .Text7.Move Text6.Width, 5800, .ScaleWidth / 2, 300
        .Picture1.Move 0, 6100, .ScaleWidth, 300
        Set .Label1.Container = .Picture1
        .Label1.Move 0, 0, 0, .Picture1.ScaleHeight
        .Label1.BackColor = vbGreen
        
        .mnuSetting.Visible = False
        .Text1.Text = .Winsock1.LocalIP
        .Text2.Text = "9999"
        .Text3.Text = ""
        .Text4.Text = "9999"
        .Text5.Text = ""
        .Text6.Text = ""
        .Text7.Text = ""
        
        .Text1.ToolTipText = "我的IP"
        .Text2.ToolTipText = "我的端口"
        .Text3.ToolTipText = "对方IP"
        .Text4.ToolTipText = "对方端口"
        .Text5.ToolTipText = "文件位置"
        .Text6.ToolTipText = "文件格式"
        .Text7.ToolTipText = "文件大小"
        .List1.ToolTipText = "将文件拖到此处"
        .Picture1.ToolTipText = "进度条"
        .Label1.ToolTipText = "进度条"
        
        .Text1.BackColor = &H80000000
        .Text2.BackColor = &H80000000
        .Text3.BackColor = &H80000003
        .Text4.BackColor = &H80000003
        .Text5.BackColor = &H80000000
        .Text6.BackColor = &H80000000
        .Text7.BackColor = &H80000000
        
        
'        .Text1.Locked = True
'        .Text2.Locked = True
'        .Text4.Locked = True
        .Text5.Locked = True
        .Text6.Locked = True
        .Text7.Locked = True
        
        .Winsock1.Protocol = sckUDPProtocol
        .Winsock1.Bind 9999
    End With
    ProgressBarMinValue = Label1.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
End
End Sub

Private Sub List1_Click()
    Text5.Text = FilePath(List1.ListIndex)
    Text6.Text = FileExtName(List1.ListIndex)
    Text7.Text = FileSize(List1.ListIndex)
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Me.List1.ListIndex < 0 Then
        Me.mnuDelete.Enabled = False
        Me.mnuSend.Enabled = False
    Else
        Me.mnuDelete.Enabled = True
        Me.mnuSend.Enabled = True
    End If
    If Button = 2 Then Me.PopupMenu mnuSetting
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim t1 As String, t2 As String
    On Error GoTo errorhandler
    t1 = Data.Files(1)
    t2 = GetFileSize(t1)
    If t2 = "0B" Then GoTo errorhandler
    List1.AddItem GetFileName(t1)
    
    ReDim Preserve FilePath(List1.ListCount - 1)
    ReDim Preserve FileSize(List1.ListCount - 1)
    ReDim Preserve FileExtName(List1.ListCount - 1)
    
    FilePath(List1.ListCount - 1) = t1
    FileSize(List1.ListCount - 1) = GetFileSize(t1)
    FileExtName(List1.ListCount - 1) = GetFileExtName(t1)
    
    Text5.Text = FilePath(List1.ListCount - 1)
    Text6.Text = FileExtName(List1.ListCount - 1)
    Text7.Text = FileSize(List1.ListCount - 1)
    Exit Sub
    
errorhandler:
    MsgBox "不支持的文件", vbInformation
End Sub

Private Sub mnuAbout_Click()
    MsgBox Info
End Sub

Private Sub mnuDelete_Click()
    Dim i As Integer
    If Me.List1.ListIndex > -1 Then
        For i = List1.ListIndex To List1.ListCount - 2
            FilePath(i) = FilePath(i + 1)
            FileSize(i) = FileSize(i + 1)
            FileExtName(i) = FileExtName(i + 1)
        Next i
        Text5.Text = ""
        Text6.Text = ""
        Text7.Text = ""
        Me.List1.RemoveItem Me.List1.ListIndex
    End If
End Sub

Private Sub mnuOpen_Click()
    Dim p As String
    p = "explorer.exe " & MainPath
    Shell p, vbNormalFocus
End Sub

Private Sub mnuSend_Click()
    If Label1.Width > ProgressBarMinValue Then
        MsgBox "传输过程中请勿操作", vbInformation
        Exit Sub
    End If
    With Me.Winsock1
        .RemoteHost = Trim(Text3.Text)
        .RemotePort = Trim(Text4.Text)
    End With
    IpStrToIpByte Trim(Text1.Text), DataPackSend.IpBit, 0
    DataPackSend.ControlBit = 0
    
    SendBuffer List1.Text, FileExtName(List1.ListIndex), FileSize(List1.ListIndex)
    
    PackToMsg 0, 1
    
    On Error GoTo errorhandler
    Me.Winsock1.SendData Msg
    Timer1.Enabled = True
    Exit Sub

errorhandler:
    MsgBox "错误的IP地址"
End Sub

Private Sub Timer1_Timer()
    If DataPackSend.ControlBit = 0 Then
        MsgBox "请求超时", vbInformation
        Close
    End If
    Timer1.Enabled = False
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Erase Msg, DataPackRecv.DataBit
    ReDim Msg(bytesTotal - 1)
    Me.Winsock1.GetData Msg
    
    If bytesTotal > 5 Then
        PackToMsg 1, 1
    Else
        PackToMsg 1, 0
    End If
    
    Select Case DataPackRecv.ControlBit
        Case 0
            Dim s1 As String
            IpStrToIpByte s1, DataPackRecv.IpBit, 1
            If RemoteIp = "" Then
                RemoteIp = s1
                Text3.Text = s1
            Else
                Exit Sub
            End If
            RecvBuffer
            DataPackSend.ControlBit = 1
            PackToMsg 0, 0
            With Me.Winsock1
                .RemoteHost = RemoteIp
                .RemotePort = Trim(Text4.Text)
                .SendData Msg
            End With
        Case 1, 3
            If DataPackRecv.ControlBit = 1 Then
                Open FilePath(List1.ListIndex) For Binary Access Read As #1
            End If
            If LOF(1) - Loc(1) < MsgSize Then
                ReDim DataPackSend.DataBit(LOF(1) - Loc(1) - 1)
                DataPackSend.ControlBit = 4
            Else
                ReDim DataPackSend.DataBit(MsgSize - 1)
                DataPackSend.ControlBit = 2
            End If
            Get #1, Loc(1) + 1, DataPackSend.DataBit
            PackToMsg 0, 1
            UpdateProgressBar Loc(1), LOF(1)
            Me.Winsock1.SendData Msg
            If DataPackSend.ControlBit = 4 Then
                DoEvents
                MsgBox "发送完毕"
                Label1.Move 0, 0, 0
                Label1.Caption = ""
                Close 1
            End If
        Case 2, 4
            CreateFile "", ""
            If DataPackRecv.ControlBit = 2 Then
                DataPackSend.ControlBit = 3
                PackToMsg 0, 0
                Me.Winsock1.SendData Msg
            End If
    End Select

End Sub


