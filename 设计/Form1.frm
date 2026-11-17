VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "答案之书·心灵抽签"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6615
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4200
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000C8FF&
      Caption         =   "启动"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00222222&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   5400
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r%, p%
Dim start As Boolean
Private Sub Command1_Click(Index As Integer)
If start = False Then
MsgBox "请先点击启动"
Exit Sub
End If
Command1(Index).Visible = False
Label1(Index).Visible = False
MsgBox "" & Label1(Index).Caption
'Text1 = "" & Label1(Index).Caption



Dim s$, y$, c$
c = Label1(Index).Caption
s = y
y = c
's = y
Text1 = y
Text2 = s
Label1(Index).Caption = "/"
Label1(Index).Visible = True

start = False

End Sub

Private Sub Command2_Click()
start = True
Dim i%, j%
    MsgBox "在心里默念一个问题吧"
'Randomize Timer
For i = 1 To 10  '行
  For j = 1 To 10  '列
    p = i * 100 + j '编号
'Label1(i * 100 + j).Visible = True '显示格子
Command1(i * 100 + j).Visible = True
Label1(i * 100 + j).Caption = ""   '清除
'Text1 = ""
Randomize '重置文字
    r = Int(Rnd * 10) + 1 '给10种文字编号并放置
     If r = 1 Then
                Label1(p).Caption = "当然可以"
            ElseIf r = 2 Then
                Label1(p).Caption = "再等等看吧"
            ElseIf r = 3 Then
                Label1(p).Caption = "放心去做"
            ElseIf r = 4 Then
                Label1(p).Caption = "不太建议"
            ElseIf r = 5 Then
                Label1(p).Caption = "马上会好"
            ElseIf r = 6 Then
                Label1(p).Caption = "值得一试"
            ElseIf r = 7 Then
                Label1(p).Caption = "别太着急"
            ElseIf r = 8 Then
                Label1(p).Caption = "一定会成"
            ElseIf r = 9 Then
                Label1(p).Caption = "看你心意"
            ElseIf r = 10 Then
                Label1(p).Caption = "问问你的朋友吧"
    End If
    Next j
Next i


End Sub

Private Sub Form_Load()
Dim i%, j%, r%

'建格子
For i = 1 To 10  '行
  For j = 1 To 10  '列
    r = i * 100 + j
    
Load Label1(i * 100 + j)
Load Command1(i * 100 + j)
Label1(i * 100 + j).Visible = True '显示
Command1(i * 100 + j).Visible = True
'设置位置
Label1(i * 100 + j).Left = 295 + 315 * i
Command1(i * 100 + j).Left = 285 + 315 * i
Label1(i * 100 + j).Top = 320 + 315 * j
Command1(i * 100 + j).Top = 255 + 315 * j
  Next j '循环
Next i
start = False
Text1 = ""
Text2 = ""
End Sub

