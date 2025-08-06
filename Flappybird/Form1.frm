VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "FlappyBird"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10980
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":25CA
   ScaleHeight     =   511
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   732
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   3120
      Top             =   7080
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   6960
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   6840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   600
      Top             =   6720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "最高分数:"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4275
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "分数:"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   825
      Picture         =   "Form1.frx":4818
      Top             =   600
      Width           =   2670
   End
   Begin VB.Image Image5 
      Height          =   1050
      Left            =   1200
      Picture         =   "Form1.frx":5710
      Top             =   4320
      Width           =   1740
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   720
      Picture         =   "Form1.frx":6514
      Top             =   2640
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   1680
      Index           =   2
      Left            =   5040
      Picture         =   "Form1.frx":6644
      Top             =   6120
      Width           =   5040
   End
   Begin VB.Image Image2 
      Height          =   1680
      Index           =   1
      Left            =   0
      Picture         =   "Form1.frx":7EC9
      Top             =   6120
      Width           =   5040
   End
   Begin VB.Image Image4 
      Height          =   4800
      Index           =   1
      Left            =   3120
      Picture         =   "Form1.frx":974E
      Top             =   4320
      Width           =   780
   End
   Begin VB.Image Image4 
      Height          =   4800
      Index           =   2
      Left            =   5400
      Picture         =   "Form1.frx":A5C4
      Top             =   4320
      Width           =   780
   End
   Begin VB.Image Image3 
      Height          =   4800
      Index           =   2
      Left            =   5520
      Picture         =   "Form1.frx":B43A
      Top             =   -1440
      Width           =   780
   End
   Begin VB.Image Image3 
      Height          =   4800
      Index           =   1
      Left            =   3120
      Picture         =   "Form1.frx":C2B9
      Top             =   -2520
      Width           =   780
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim h As Double, h1 As Double, t As Long, a As Integer, f As Integer, zhu As Integer, score As Long, maxscore As Long
Const g = 0.02
Dim hh(1 To 2) As Integer
Dim starting As Boolean
Private Sub Command1_Click()
Call Form_Load
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 And starting = True Then
    Image1.Top = Image1.Top - 25
    t = 0
    h1 = Image1.Top
End If
End Sub

Private Sub Form_Load()
Open App.Path & "\score.txt" For Append As #1
    Print #1, "0"
Close #1
Open App.Path & "\score.txt" For Input As #1
         If Not (EOF(1)) Then
         Line Input #1, Max
        maxscore = Val(Max)
        End If
Close #1




Timer3.Enabled = False
t = 0

Form1.Caption = "FlappyBird"
Form1.Height = 8250
Form1.Width = 4545

Image5.Visible = True
Image6.Visible = True

Image1.Left = 60
Image1.Top = 200
h1 = 200
a = 0
f = 2
zhu = 1
score = 0
starting = False
Label1.Caption = "分数:"
Label2.Caption = "最高分数:" & maxscore

Timer2.Enabled = True
Timer4.Enabled = True

Image3(1).Left = 2000
Image3(1).Top = -150
Image4(1).Top = 270
Image4(1).Left = Image3(1).Left
Label1.Visible = False
Label2.Visible = True

Randomize Timer

Image3(2).Left = Image3(1).Left + 250
Image4(2).Left = Image4(1).Left + 250
Image3(2).Top = Int(Rnd() * (-200)) - 100
hh(2) = Image3(2).Top + Image3(2).Height
Image4(2).Top = hh(2) + 100

End Sub

Private Sub Picture1_Click()

End Sub



Private Sub Image5_Click()
Call start
End Sub

Private Sub Timer1_Timer() '管道



'柱子左移
For i = 1 To 2
    Image3(i).Left = Image3(i).Left - 3
    Image4(i).Left = Image4(i).Left - 3
    
    If Image3(i).Left + Image3(i).Width < 0 Then
        Image3(i).Left = Image3(f).Left + 250
        Image4(i).Left = Image4(f).Left + 250
        Image3(i).Top = Int(Rnd() * (-200)) - 100: hh(i) = Image3(i).Top + Image3(i).Height
        Image4(i).Top = hh(i) + 100
        f = f + 1
        If f = 3 Then f = 1
    End If
Next i





End Sub

Private Sub gameover()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = True
Timer4.Enabled = False
t = 0
If score >= maxscore Then
    Open App.Path & "\score.txt" For Output As #1
        Print #1, Trim(Str(score))
    Close #1
End If

starting = False





End Sub
Private Sub start()
starting = True

Image5.Visible = False
Image6.Visible = False

Timer3.Enabled = False
Timer1.Enabled = True

Image3(1).Left = 350
Image3(1).Top = -150
Image4(1).Top = 270
Image4(1).Left = Image3(1).Left

Image1.Left = 60
Image1.Top = 200

Image3(2).Left = Image3(1).Left + 250
Image4(2).Left = Image4(1).Left + 250


Label1.Visible = True
Label2.Visible = False
End Sub

Private Sub Timer2_Timer() '鸟

a = a + 1
If a = 16 Then a = 1
fff = 0

'撞柱死亡
If Image1.Left >= Image3(zhu).Left - 35 And Image1.Left <= Image3(zhu).Left + Image3(zhu).Width - 10 And Image1.Top + 12 <= Image3(zhu).Top + Image3(zhu).Height Then GoTo En
If Image1.Left >= Image4(zhu).Left - 35 And Image1.Left <= Image4(zhu).Left + Image4(zhu).Width - 10 And Image1.Top + Image1.Height - 12 >= Image4(zhu).Top Then GoTo En

'加分
If Image1.Left > Image3(zhu).Left + Image3(zhu).Width Then
    score = score + 1
    zhu = zhu + 1
    If zhu = 3 Then zhu = 1
    Label1.Caption = "分数:" & score
End If
'落地触天死亡
If Image1.Top + 38 >= Image2(1).Top Or Image1.Top + 38 >= Image2(2).Top Or Image1.Top <= -10 Then GoTo En



'飞翔动画
If a <= 5 Then Image1.Picture = LoadPicture(App.Path & "\PIC\bird1_0.gif")
If a > 5 And a <= 10 Then Image1.Picture = LoadPicture(App.Path & "\PIC\bird1_1.gif")
If a > 10 And a <= 15 Then Image1.Picture = LoadPicture(App.Path & "\PIC\bird1_2.gif")

'小鸟下坠
If starting = True Then
    t = t + 1
    h = 1 / 2 * g * t * t + h1
    Image1.Top = h
End If
fff = 1



En:
If fff = 0 Then Call gameover
End Sub

Private Sub Timer3_Timer() '死亡动画

If Image1.Top < 380 Then
    t = t + 1
    h = 1 / 2 * 0.02 * t * t
    Image1.Top = h + Image1.Top
    If Image1.Top >= 380 Then
        
        Call Form_Load
     
        
    End If
End If
End Sub

Private Sub Timer4_Timer() '陆地
'陆地左移
Image2(1).Left = Image2(1).Left - 3
Image2(2).Left = Image2(2).Left - 3
 
'陆地循环
If Image2(1).Left + Image2(1).Width < 0 Then
    Image2(1).Left = Image2(2).Left + Image2(2).Width
End If
If Image2(2).Left + Image2(2).Width < 0 Then
    Image2(2).Left = Image2(1).Left + Image2(1).Width
End If
End Sub
