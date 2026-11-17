VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "寻机头"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   13245
   StartUpPosition =   2  '屏幕中心
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
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
   Begin VB.Label Label3 
      Caption         =   "    202402081017毛昱彤"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "       飞机图样示例"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   5265
      Left            =   7080
      Picture         =   "Formz.frx":0000
      Top             =   720
      Width           =   5265
   End
   Begin VB.Label Score 
      Height          =   735
      Left            =   4920
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
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
Dim start As Boolean
Dim scoree As Integer
Dim h1x%, h1y%, dir1%, type1%
Dim h2x%, h2y%, dir2%, type2%
Dim found As Integer

Private Sub Command1_Click(Index As Integer)
If start = False Then Exit Sub '判断是否开始
    Command1(Index).Enabled = False '选择的按键不可再选择
'Command1(Index).Visible = False
    Select Case Command1(Index).Tag '判断点使用Select Case判断当前格子的Tag值
        Case 2 '机头
            Command1(Index).BackColor = vbRed '点击机头显示红色
            found = found + 1 '记录机头数
            scoree = scoree - 1
            If found = 2 Then '找到两个就结束游戏
                MsgBox "找到两架机头！最终得分：" & scoree
                start = False
            Command2.Caption = "开始游戏"
            End If
        Case 1 '机身
            Command1(Index).BackColor = vbBlue '点击机身显示蓝色
            scoree = scoree - 1 '积分-1
            Score.Caption = "当前积分：" & scoree '显示积分
        Case 0 '空白
            Command1(Index).BackColor = vbBlack '点击空格处显示黑色
            scoree = scoree - 1 '积分-1
            Score.Caption = "当前积分：" & scoree
    End Select
    
End Sub

Private Sub Command2_Click()
Dim success As Boolean '用来判断飞机是否放置好
Dim response As VbMsgBoxResult
 If start = False Or found = 2 Then
start = True
scoree = 109
found = 0
success = False
Command2.Caption = "游戏中"

For i = 1 To 11
    For j = 1 To 11
    p = i * 100 + j
    Command1(p).Tag = 0
    Command1(p).Enabled = True
    Command1(p).BackColor = vbButtonFace
    Next j
Next i

Score.Caption = "当前积分：" & scoree

Randomize '生成第一架飞机（类型1）
    type1 = 1
    Do
        success = False
        Call Create(h1x, h1y, dir1, type1, success)
    Loop While success = False
    
'生成第二架飞机（类型2）
    type2 = 2
    Do
        success = False
        Call Create(h2x, h2y, dir2, type2, success)
    Loop While success = False
End If


    response = MsgBox("确定要重置吗？重置后积分将清除。", vbQuestion + vbYesNo + vbDefaultButton2, "确认重置")
    
    If response = vbYes Then
        '重置游戏状态
        start = True
        scoree = 109
        found = 0
        Command2.Caption = "游戏中"
        '重置所有的颜色和状态
        For i = 1 To 11
            For j = 1 To 11
                p = i * 100 + j
                
                '重置颜色
                Command1(p).BackColor = vbButtonFace  '恢复默认钮颜色
                Label1(p).BackColor = vbWhite
                '重置状态
                Command1(p).Tag = 0                    '子重置为空白
                Command1(p).Enabled = True             '启用所有格子
            Next j
        Next i
        
        '更新分数显示
        Score.Caption = "当前积分：" & scoree
        
        Randomize
        
        '生成第一架飞机
        type1 = 1
        Do
            success = False
            Call Create(h1x, h1y, dir1, type1, success)
        Loop While success = False
        
        '生成第二架飞机
        type2 = 2
        Do
            success = False
            Call Create(h2x, h2y, dir2, type2, success)
        Loop While success = False
    End If
    '如果用户选择"否"，继续游戏
End Sub

Private Sub Create(x%, y%, dir%, ptyp%, ByRef success As Boolean) '创建一个子程序生成飞机
    Dim p%, i%, posList() As String, idx% '用数组暂时存放飞机的点
    dir = Int(Rnd * 4)  ' 0上 1下 2左 3右四个方向
    
    Select Case dir
        Case 0: x = Int(Rnd * 7) + 3: y = Int(Rnd * 8) + 1  ' 机头朝上
        Case 1: x = Int(Rnd * 7) + 3: y = Int(Rnd * 8) + 4  ' 机头朝下
        Case 2: x = Int(Rnd * 8) + 1: y = Int(Rnd * 7) + 3  ' 机头朝左
        Case 3: x = Int(Rnd * 8) + 4: y = Int(Rnd * 7) + 3  ' 机头朝右
    End Select



idx = 0 '给放的点编号
ReDim posList(0 To 9) '飞机有十个格子
    posList(idx) = x & "," & y ' 存机头
    idx = idx + 1
    
    
If ptyp = 1 Then ' 存第一架飞机
        Select Case dir
            Case 0 ' 上
                posList(idx) = (x) & "," & (y - 1): idx = idx + 1
                posList(idx) = (x) & "," & (y - 2): idx = idx + 1
                posList(idx) = (x) & "," & (y - 3): idx = idx + 1
                posList(idx) = (x - 1) & "," & (y - 1): idx = idx + 1
                posList(idx) = (x - 2) & "," & (y - 1): idx = idx + 1
                posList(idx) = (x + 1) & "," & (y - 1): idx = idx + 1
                posList(idx) = (x + 2) & "," & (y - 1): idx = idx + 1
                posList(idx) = (x - 1) & "," & (y - 3): idx = idx + 1
                posList(idx) = (x + 1) & "," & (y - 3): idx = idx + 1
            Case 1 ' 下
                posList(idx) = (x) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x) & "," & (y + 2): idx = idx + 1
                posList(idx) = (x) & "," & (y + 3): idx = idx + 1
                posList(idx) = (x - 1) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x - 2) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x + 1) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x + 2) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x - 1) & "," & (y + 3): idx = idx + 1
                posList(idx) = (x + 1) & "," & (y + 3): idx = idx + 1
            Case 2 ' 左
                posList(idx) = (x + 1) & "," & (y): idx = idx + 1
                posList(idx) = (x + 2) & "," & (y): idx = idx + 1
                posList(idx) = (x + 3) & "," & (y): idx = idx + 1
                posList(idx) = (x + 1) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x + 1) & "," & (y - 1): idx = idx + 1
                posList(idx) = (x + 1) & "," & (y + 2): idx = idx + 1
                posList(idx) = (x + 1) & "," & (y - 2): idx = idx + 1
                posList(idx) = (x + 3) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x + 3) & "," & (y - 1): idx = idx + 1
            Case 3 ' 右
                posList(idx) = (x - 1) & "," & (y): idx = idx + 1
                posList(idx) = (x - 2) & "," & (y): idx = idx + 1
                posList(idx) = (x - 3) & "," & (y): idx = idx + 1
                posList(idx) = (x - 1) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x - 1) & "," & (y - 1): idx = idx + 1
                posList(idx) = (x - 1) & "," & (y + 2): idx = idx + 1
                posList(idx) = (x - 1) & "," & (y - 2): idx = idx + 1
                posList(idx) = (x - 3) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x - 3) & "," & (y - 1): idx = idx + 1
        End Select
        
Else ' 第二架飞机
        Select Case dir
            Case 0 ' 上
                posList(idx) = (x) & "," & (y - 1): idx = idx + 1
                posList(idx) = (x) & "," & (y - 2): idx = idx + 1
                posList(idx) = (x) & "," & (y - 3): idx = idx + 1
                posList(idx) = (x - 1) & "," & (y - 1): idx = idx + 1
                posList(idx) = (x - 2) & "," & (y - 2): idx = idx + 1
                posList(idx) = (x + 1) & "," & (y - 1): idx = idx + 1
                posList(idx) = (x + 2) & "," & (y - 2): idx = idx + 1
                posList(idx) = (x - 1) & "," & (y - 3): idx = idx + 1
                posList(idx) = (x + 1) & "," & (y - 3): idx = idx + 1
            Case 1 ' 下
                posList(idx) = (x) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x) & "," & (y + 2): idx = idx + 1
                posList(idx) = (x) & "," & (y + 3): idx = idx + 1
                posList(idx) = (x - 1) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x - 2) & "," & (y + 2): idx = idx + 1
                posList(idx) = (x + 1) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x + 2) & "," & (y + 2): idx = idx + 1
                posList(idx) = (x - 1) & "," & (y + 3): idx = idx + 1
                posList(idx) = (x + 1) & "," & (y + 3): idx = idx + 1
            Case 2 '左
                posList(idx) = (x + 1) & "," & (y): idx = idx + 1
                posList(idx) = (x + 2) & "," & (y): idx = idx + 1
                posList(idx) = (x + 3) & "," & (y): idx = idx + 1
                posList(idx) = (x + 1) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x + 1) & "," & (y - 1): idx = idx + 1
                posList(idx) = (x + 2) & "," & (y + 2): idx = idx + 1
                posList(idx) = (x + 2) & "," & (y - 2): idx = idx + 1
                posList(idx) = (x + 3) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x + 3) & "," & (y - 1): idx = idx + 1
            Case 3 '右
                posList(idx) = (x - 1) & "," & (y): idx = idx + 1
                posList(idx) = (x - 2) & "," & (y): idx = idx + 1
                posList(idx) = (x - 3) & "," & (y): idx = idx + 1
                posList(idx) = (x - 1) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x - 1) & "," & (y - 1): idx = idx + 1
                posList(idx) = (x - 2) & "," & (y + 2): idx = idx + 1
                posList(idx) = (x - 2) & "," & (y - 2): idx = idx + 1
                posList(idx) = (x - 3) & "," & (y + 1): idx = idx + 1
                posList(idx) = (x - 3) & "," & (y - 1): idx = idx + 1
        End Select
    End If

Dim q%, checkX%, checkY%, checkP%
    For q = 0 To idx - 1 '检查到10个格子
        checkX = CInt(Split(posList(q), ",")(0)) '把坐标点分割成数组["x","y"]，返回整数
        checkY = CInt(Split(posList(q), ",")(1))
        checkP = checkX * 100 + checkY
        If checkX < 1 Or checkX > 11 Or checkY < 1 Or checkY > 11 Then
    Exit Sub '超出边界，生成失败
End If
        If Command1(checkP).Tag <> 0 Then '只要有一个格子被占了，就直接退出（检查重叠），不=0就推出
            Exit Sub
        End If
    Next q
    
    ' 打Tag
    For q = 0 To idx - 1
        checkX = CInt(Split(posList(q), ",")(0))
        checkY = CInt(Split(posList(q), ",")(1))
        checkP = checkX * 100 + checkY '飞机的点标号
        If q = 0 Then '第一个格子是机头，Tag=2
            Command1(checkP).Tag = 2
        Else '其他格子是机身，Tag=1
            Command1(checkP).Tag = 1
        End If
    Next q
    
    ' 标记生成成功
    success = True
End Sub




Private Sub Form_Load()
Dim i%, j%, p%
scoree = 109 '默认积分总数
found = 0 '默认找到的机头数0
For i = 1 To 11  '行
  For j = 1 To 11  '列
   p = i * 100 + j '格子编号
Load Label1(p) '加载格子
Load Command1(p)
Label1(p).Left = 295 + 315 * i '调整位置
Command1(p).Left = 285 + 315 * i
Label1(p).Top = 320 + 315 * j
Command1(p).Top = 255 + 315 * j
    Label1(p).Visible = True '显示格子
    Command1(p).Visible = True
    Command1(p).BackColor = vbButtonFace '设置格子默认颜色
   
    Next j
Next i
Score.Caption = "当前积分：0"
    start = False
Command2.Caption = "开始"
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Picture1_Click()

End Sub
