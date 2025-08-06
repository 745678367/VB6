Attribute VB_Name = "Module1"
Option Explicit
'服务器IP
Public ServerIP As String
'已连接的客户端数量
Public ClientQuantities As Integer
'五子棋玩家信息数据库
Public DBChesscnn As ADODB.Connection
Public DBChessset As ADODB.Recordset

Public Const Info As String = "版本号：1.2.2" & vbCrLf & "作于：2024年7月24日" & vbCrLf & "最后修改于：2024年7月27日" & vbCrLf & vbCrLf & vbCrLf & "Written by 刘栋泽"

Public Sub Main()
    If App.PrevInstance Then
        MsgBox "程序已在运行中！", vbInformation, "提示"
    Else
        Form1.Show
    End If
End Sub
