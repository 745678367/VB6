Attribute VB_Name = "Module1"
Option Explicit
Option Base 0
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Type DataPack
    IpBit(3) As Byte
    ControlBit As Byte
    DataBit() As Byte
End Type

Public DataPackSend As DataPack
Public DataPackRecv As DataPack
Public Msg() As Byte

Public FilePath() As String
Public FileSize() As String
Public FileExtName() As String

'ControlBit位
'0-请求传输
'1-同意传输
'2-数据字节
'3-字节确认
'4-传输完毕

Public MainPath As String
Public RemoteIp As String
Public ProgressBarMinValue As Integer
Public MsgSize As Long

Public Fs As Long

Public Const Info As String = "版本：2.1.2" & vbCrLf & "传输协议：UDP" & vbCrLf & "作于：2024年11月12日" & vbCrLf & "最后修改于：2024年11月23日" & vbCrLf & vbCrLf & "Written by 柳丁"

Public Sub Main()
    On Error GoTo errorhandler
    Shell App.Path & "\注册运行库.bat", vbHide
    MainPath = App.Path & "\文件传输\"
    If Dir(MainPath, vbDirectory) = "" Then
    MkDir MainPath
    End If
    MsgSize = 1024 * 7
    Form1.Show
    Exit Sub
    
errorhandler:
    MsgBox "运行库丢失", vbInformation
End Sub

Public Function GetFileName(FilePath As String) As String
    Dim i As Integer, j As Integer
    i = InStrRev(FilePath, "\")
    j = InStrRev(FilePath, ".")
    GetFileName = Mid(FilePath, i + 1, j - i - 1)
End Function

Public Function GetFileSize(FilePath As String) As String
    Dim i As Single
    i = FileLen(FilePath)
    Select Case i
        Case Is < 1024
            GetFileSize = i & "B"
        Case Is < 1048576
            i = Round(i / 1024, 2)
            GetFileSize = i & "KB"
        Case Else
            i = Round(i / 1048576, 2)
            GetFileSize = i & "MB"
    End Select
End Function

Public Function GetFileExtName(FilePath As String) As String
    Dim i As Integer
    i = InStrRev(FilePath, ".")
    GetFileExtName = Mid(FilePath, i + 1)
End Function

Public Sub IpStrToIpByte(IpStr As String, Ipbyte() As Byte, Mode As Byte)
'mode
'0-字符串转字节数组
'1-字节数组转字符串
    Select Case Mode
        Case 0
            Dim t() As String, i As Integer
            t = Split(IpStr, ".")
            For i = 0 To 3
                Ipbyte(i) = t(i)
            Next
        Case 1
            IpStr = Ipbyte(0) & "." & Ipbyte(1) & "." & Ipbyte(2) & "." & Ipbyte(3)
    End Select
End Sub

Public Sub PackToMsg(Mode As Byte, HasDataBit As Byte)
'mode
'0-Pack转Msg
'1-Msg转Pack

'HasDataBit
'0-无数据位
'1-有数据位
    Select Case Mode
        Case 0
            ReDim Msg(4)
            CopyMemory Msg(0), DataPackSend.IpBit(0), 4
            CopyMemory Msg(4), DataPackSend.ControlBit, 1
            If HasDataBit = 1 Then
                ReDim Preserve Msg(5 + UBound(DataPackSend.DataBit))
                CopyMemory Msg(5), DataPackSend.DataBit(0), UBound(DataPackSend.DataBit) + 1
            End If
        Case 1
            CopyMemory DataPackRecv.IpBit(0), Msg(0), 4
            CopyMemory DataPackRecv.ControlBit, Msg(4), 1
            If HasDataBit = 1 Then
                ReDim Preserve DataPackRecv.DataBit(UBound(Msg) - 5)
                CopyMemory DataPackRecv.DataBit(0), Msg(5), UBound(DataPackRecv.DataBit) + 1
            End If
    End Select
    DoEvents
End Sub

Public Sub SendBuffer(FileName As String, FileExtName As String, FileSize As String)
    Open MainPath & "SendBuffer" & Trim(Form1.Text2.Text) & ".ldz" For Output As #9
        Write #9, FileName
        Write #9, FileExtName
        Write #9, FileSize
    Close #9
    
    Open MainPath & "SendBuffer" & Trim(Form1.Text2.Text) & ".ldz" For Binary Access Read As #9
        ReDim DataPackSend.DataBit(LOF(9) - 1)
        Get #9, 1, DataPackSend.DataBit
    Close #9
End Sub

Public Sub RecvBuffer()
    Open MainPath & "RecvBuffer" & Trim(Form1.Text2.Text) & ".ldz" For Binary Access Write As #8
        Put #8, 1, DataPackRecv.DataBit
    Close #8
    
    Dim FileName0 As String, FileExtName0 As String, FileSize0 As String
    Open MainPath & "RecvBuffer" & Trim(Form1.Text2.Text) & ".ldz" For Input As #8
        Input #8, FileName0
        Input #8, FileExtName0
        Input #8, FileSize0
    Close #8
    
    Dim t As String
    t = Time
    If Len(t) < 8 Then t = "0" & Time
    FileName0 = FileName0 & Mid(t, 1, 2) & Mid(t, 4, 2) & Mid(t, 7, 2)
    CreateFile FileName0, FileExtName0
    
    Form1.List1.AddItem FileName0
    
    ReDim Preserve FilePath(Form1.List1.ListCount - 1)
    ReDim Preserve FileSize(Form1.List1.ListCount - 1)
    ReDim Preserve FileExtName(Form1.List1.ListCount - 1)
    
    FilePath(Form1.List1.ListCount - 1) = MainPath & FileName0 & "." & FileExtName0
    FileSize(Form1.List1.ListCount - 1) = FileSize0
    FileExtName(Form1.List1.ListCount - 1) = FileExtName0
    
    Form1.Text5.Text = FilePath(Form1.List1.ListCount - 1)
    Form1.Text6.Text = FileExtName(Form1.List1.ListCount - 1)
    Form1.Text7.Text = FileSize(Form1.List1.ListCount - 1)
    
    '-----------------------
    Dim Fb As String
    Fb = Mid(FileSize0, Len(FileSize0) - 1, 2)
    If Mid(Fb, 1, 1) >= 0 And Mid(Fb, 1, 1) <= 9 Then Fb = Mid(Fb, 2, 1)
    Select Case Fb
        Case "B"
            Fs = Val(Mid(FileSize0, 1, Len(FileSize0) - Len(Fb)))
        Case "KB"
            Fs = Val(Mid(FileSize0, 1, Len(FileSize0) - Len(Fb))) * 1024
        Case "MB"
            Fs = Val(Mid(FileSize0, 1, Len(FileSize0) - Len(Fb))) * 1048576
    End Select
End Sub

Public Sub CreateFile(FileName As String, FileExtName As String)
    If DataPackRecv.ControlBit = 0 Then
        Open MainPath & FileName & "." & FileExtName For Binary Access Write As #7
    Else
        Put #7, Loc(7) + 1, DataPackRecv.DataBit
        UpdateProgressBar Loc(7), Fs
        If DataPackRecv.ControlBit = 4 Then
            Close 7
            MsgBox "接受完毕"
            Form1.Label1.Move 0, 0, 0
            Form1.Label1.Caption = ""
            RemoteIp = ""
        End If
    End If
End Sub

Public Sub UpdateProgressBar(CurrentValue As Long, MaxValue As Long)
    Form1.Label1.Move 0, 0, Round(CurrentValue / MaxValue, 2) * Form1.Picture1.ScaleWidth
    Form1.Label1.Caption = Round(CurrentValue / MaxValue, 2) * 100 & "%"
End Sub

