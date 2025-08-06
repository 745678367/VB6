Attribute VB_Name = "Module2"
Option Explicit

'五子棋玩家连接服务器
Public Sub ChessID(ByVal index As Integer, ByVal playerip As String, ByVal logintime As String)
    Dim i As Integer
    
    '写入数据库玩家信息
    ConnectAccess2.RefreshDB DBChesscnn
    With DBChessset
        If .State = 1 Then .Close
        .Open "select * from 五子棋", DBChesscnn
        .AddNew
        For i = 1 To 6
            .Fields(1).Value = playerip
            .Fields(2).Value = Trim(Str(index))
            .Fields(4).Value = logintime
        Next i
        .Update
        .Close
    End With
    '将唯一的索引号发给对应玩家
    Form1.wskServer(index).SendData "五子棋Index" & Trim(Str(index))
End Sub

'转发邀请
Public Sub Invite(ByVal host_index As String, ByVal guest_index As String)
    Dim host_ip As String, host_index1 As String * 2
    
    ConnectAccess2.RefreshDB DBChesscnn
    With DBChessset
        If .State = 1 Then .Close
        .Open "select * from 五子棋", DBChesscnn
        .Find "PlayerIndex='" & guest_index & "'"
        If .EOF Then
            Form1.wskServer(host_index).SendData "五子棋邀请失败:" & "没有找到ID为" & guest_index & "的玩家"
        ElseIf Len(.Fields(3).Value) > 0 Then
            Form1.wskServer(host_index).SendData "五子棋邀请失败:" & guest_index & "正在进行对战"
        Else
            .MoveFirst
            .Find "PlayerIndex='" & host_index & "'"
            host_ip = .Fields(1).Value
            host_index1 = host_index
            Form1.wskServer(Val(Trim(guest_index))).SendData "五子棋邀请对战:" & host_index1 & host_ip
        End If
        .Close
    End With
End Sub

'接受邀请
Public Sub Accept(ByVal host_index As String, ByVal guest_index As String)
    Dim host_ip As String, guest_ip As String
    ConnectAccess2.RefreshDB DBChesscnn
    With DBChessset
        If .State = 1 Then .Close
        .Open "select * from 五子棋", DBChesscnn
        .Find "PlayerIndex='" & host_index & "'"
        .Fields(3).Value = guest_index
        host_ip = .Fields(1).Value
        .MoveFirst
        .Find "PlayerIndex='" & guest_index & "'"
        .Fields(3).Value = host_index
        guest_ip = .Fields(1).Value
        .Update
        .Close
    End With

    Form1.wskServer(Val(Trim(host_index))).SendData "五子棋联机对战:开始" & guest_ip
    DoEvents
    Form1.wskServer(Val(Trim(guest_index))).SendData "五子棋联机对战:开始" & host_ip
    DoEvents
End Sub

'拒绝邀请
Public Sub Refuse(ByVal host_index As String, ByVal guest_index As String)
    Form1.wskServer(Val(Trim(host_index))).SendData "五子棋拒绝邀请" & guest_index
End Sub

'玩家退出对战，更新数据库数据
Public Sub ExitCompetition(ByVal host_id As String, ByVal guest_id As String)
    ConnectAccess2.RefreshDB DBChesscnn
    With DBChessset
        If .State = 1 Then .Close
        .Open "select * from 五子棋"
        .Find "PlayerIndex='" & host_id & "'"
        .Fields(3).Value = ""
        .MoveFirst
        .Find "PlayerIndex='" & guest_id & "'"
        .Fields(3).Value = ""
        .Update
        .Close
    End With
End Sub

'删除数据库信息
Public Sub ClearDB()
    ConnectAccess2.RefreshDB DBChesscnn
    With DBChessset
        If .State = 1 Then .Close
        .Open "delete * from 五子棋", DBChesscnn
    End With
End Sub
