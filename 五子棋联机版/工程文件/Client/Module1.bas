Attribute VB_Name = "Module1"
Option Explicit
'服务器IP
Public ServerIP As String
'我的ID、IP地址、我的棋子颜色
Public MyIndex As String * 2, MyIP As String, MyColor As String * 1
'我是否正在游戏中
Public Gaming As Boolean
'我是否在线
Public OnlineGame As Boolean
'登录时间
Public LoginTime As String
'对战玩家ID、IP地址、棋子颜色
Public Competitor_Index As String * 2, Competitor_IP As String, Competitor_Color As String * 1
'单机、联机模式下棋权交换,T黑F白
Public FlagA As Boolean, FlagB As Boolean
'将棋盘数据存放在数组中
Public Memory(20, 20) As String
'获取棋盘的横纵坐标
Public Type Location
    row As Integer
    column As Integer
End Type

Public Const Info As String = "版本号：1.2.2" & vbCrLf & "作于：2024年7月24日" & vbCrLf & "最后修改于：2024年7月27日" & vbCrLf & vbCrLf & vbCrLf & "Written by 刘栋泽"
Public Const Info1 As String = "LiUDinG"

Public Sub Main()
    If App.PrevInstance Then
        MsgBox "程序已在运行中！", vbInformation, "提示"
    Else
    ServerIP = ""
    Form1.Show
    End If
End Sub

'检测五子连珠
Public Sub Detection(ByVal index As Integer)
    Dim loc As Location, n1 As Integer
    n1 = index
    Do Until index < 21
        index = index - 20
        loc.row = loc.row + 1
    Loop
    loc.row = loc.row + 1
    loc.column = index
    Memory(loc.row, loc.column) = Form1.Cell(n1).Caption
    Form1.mnuLocation.Caption = "上一个棋落在了：" & loc.row & "," & Chr(loc.column + 64)
    
    Call Summary(loc.row, loc.column, Memory(loc.row, loc.column))
End Sub

'检测五子连珠
Public Sub Summary(ByVal i As Integer, ByVal j As Integer, ByRef Mark As String)
    Dim a As Integer, n As Integer, m As Integer, i1 As Integer, j1 As Integer, index As Integer
    Dim flag As Boolean
    m = 1
    For a = 0 To 4
        Select Case m
            Case 1
                If i - a > 0 Then
                    If Memory(i - a, j) = Mark Then
                        n = n + 1
                    Else
                        flag = True
                    End If
                Else
                    flag = True
                End If
                
                If flag Then
                    flag = False
                    n = 0
                    i1 = i - a + 1
                    j1 = j
                    a = -1
                    m = m + 1
                End If
            Case 2
                If i1 + a < 21 Then
                    If Memory(i1 + a, j1) = Mark Then
                        n = n + 1
                    Else
                        flag = True
                    End If
                Else
                    flag = True
                End If
                
                If flag Then
                    flag = False
                    n = 0
                    a = -1
                    m = m + 1
                End If
            Case 3
                If j + a < 21 Then
                    If Memory(i, j + a) = Mark Then
                        n = n + 1
                    Else
                        flag = True
                    End If
                Else
                    flag = True
                End If
                
                If flag Then
                    flag = False
                    n = 0
                    i1 = i
                    j1 = j + a - 1
                    a = -1
                    m = m + 1
                End If
            Case 4
                If j1 - a > 0 Then
                    If Memory(i1, j1 - a) = Mark Then
                        n = n + 1
                    Else
                        flag = True
                    End If
                Else
                    flag = True
                End If
                
                If flag Then
                    flag = False
                    n = 0
                    a = -1
                    m = m + 1
                End If
            Case 5
                If i - a > 0 And j + a < 21 Then
                    If Memory(i - a, j + a) = Mark Then
                        n = n + 1
                    Else
                        flag = True
                    End If
                Else
                    flag = True
                End If
                
                If flag Then
                    flag = False
                    n = 0
                    i1 = i - a + 1
                    j1 = j + a - 1
                    a = -1
                    m = m + 1
                End If
            Case 6
                If i1 + a < 21 And j1 - a > 0 Then
                    If Memory(i1 + a, j1 - a) = Mark Then
                        n = n + 1
                    Else
                        flag = True
                    End If
                Else
                    flag = True
                End If
                
                If flag Then
                    flag = False
                    n = 0
                    a = -1
                    m = m + 1
                End If
            Case 7
                If i - a > 0 And j - a > 0 Then
                    If Memory(i - a, j - a) = Mark Then
                        n = n + 1
                    Else
                        flag = True
                    End If
                Else
                    flag = True
                End If
                
                If flag Then
                    flag = False
                    n = 0
                    i1 = i - a + 1
                    j1 = j - a + 1
                    a = -1
                    m = m + 1
                End If
            Case 8
                If i1 + a < 21 And j1 + a < 21 Then
                    If Memory(i1 + a, j1 + a) = Mark Then
                        n = n + 1
                    Else
                        flag = True
                    End If
                Else
                    flag = True
                End If
                
                If flag Then
                    flag = False
                    n = 0
                    a = 4
                End If
        End Select
    Next a
    
    If n = 5 Then
        Gaming = False
        If m Mod 2 = 0 Then
            index = (i1 - 1) * 20 + j1
        Else
            index = (i - 1) * 20 + j
        End If
        Call Prompt(index, m)
        MsgBox Mark & "胜利", vbOKOnly, "游戏结束"
    End If
End Sub

'五子连珠提示
Public Sub Prompt(ByVal index As Integer, ByVal m As Integer)
    Dim i As Integer
    For i = 0 To 4
        Select Case m
            Case 1
                Form1.Cell(index - i * 20).BackColor = vbRed
            Case 2
                Form1.Cell(index + i * 20).BackColor = vbRed
            Case 3
                Form1.Cell(index + i).BackColor = vbRed
            Case 4
                Form1.Cell(index - i).BackColor = vbRed
            Case 5
                Form1.Cell(index - i * 20 + i).BackColor = vbRed
            Case 6
                Form1.Cell(index + i * 20 - i).BackColor = vbRed
            Case 7
                Form1.Cell(index - i * 20 - i).BackColor = vbRed
            Case 8
                Form1.Cell(index + i * 20 + i).BackColor = vbRed
        End Select
        Next i
            
End Sub

'联机模式下更新对手的落子
Public Sub UpdateChessboard(ByVal index As Integer)
    Form1.Cell(index).Enabled = False
    Form1.Cell(index).Caption = Competitor_Color
    Form1.Cmd3.Caption = MyColor
    Call Detection(index)
    FlagB = Not FlagB
End Sub

