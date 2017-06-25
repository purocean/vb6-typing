VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "打字游戏"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   14850
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      Caption         =   "清除历史&C"
      Height          =   495
      Left            =   9960
      TabIndex        =   26
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "剩余时间"
      Height          =   1215
      Left            =   11640
      TabIndex        =   22
      Top             =   7680
      Width           =   2775
      Begin VB.Label time 
         AutoSize        =   -1  'True
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   42
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   840
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Width           =   1740
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退出&X"
      Height          =   495
      Left            =   4800
      TabIndex        =   15
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "关于&A"
      Height          =   495
      Left            =   3360
      TabIndex        =   14
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重玩(F5)"
      Height          =   495
      Left            =   1920
      TabIndex        =   13
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始(F2)"
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "得分"
      Height          =   1335
      Left            =   6840
      TabIndex        =   11
      Top             =   7560
      Width           =   4455
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1680
         TabIndex        =   21
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   960
         TabIndex        =   20
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1440
         TabIndex        =   19
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "历史最高分："
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "失误："
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "当前得分："
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11280
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Caption         =   "控制"
      Height          =   1335
      Left            =   360
      TabIndex        =   10
      Top             =   7560
      Width           =   6255
   End
   Begin VB.Timer zimuxialuo 
      Enabled         =   0   'False
      Index           =   9
      Left            =   13560
      Top             =   360
   End
   Begin VB.Timer zimuxialuo 
      Enabled         =   0   'False
      Index           =   8
      Left            =   13440
      Top             =   0
   End
   Begin VB.Timer zimuxialuo 
      Enabled         =   0   'False
      Index           =   7
      Left            =   12720
      Top             =   480
   End
   Begin VB.Timer zimuxialuo 
      Enabled         =   0   'False
      Index           =   6
      Left            =   12840
      Top             =   720
   End
   Begin VB.Timer zimuxialuo 
      Enabled         =   0   'False
      Index           =   5
      Left            =   13200
      Top             =   480
   End
   Begin VB.Timer zimuxialuo 
      Enabled         =   0   'False
      Index           =   4
      Left            =   12600
      Top             =   120
   End
   Begin VB.Timer zimuxialuo 
      Enabled         =   0   'False
      Index           =   3
      Left            =   13320
      Top             =   720
   End
   Begin VB.Timer zimuxialuo 
      Enabled         =   0   'False
      Index           =   2
      Left            =   12480
      Top             =   600
   End
   Begin VB.Timer zimuxialuo 
      Enabled         =   0   'False
      Index           =   1
      Left            =   13200
      Top             =   360
   End
   Begin VB.Timer zimuxialuo 
      Enabled         =   0   'False
      Index           =   0
      Left            =   12240
      Top             =   360
   End
   Begin WMPLibCtl.WindowsMediaPlayer mu2 
      Height          =   555
      Left            =   10560
      TabIndex        =   25
      Top             =   3120
      Visible         =   0   'False
      Width           =   3600
      URL             =   "2.mp3"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6350
      _cy             =   979
   End
   Begin WMPLibCtl.WindowsMediaPlayer mu1 
      Height          =   495
      Left            =   10320
      TabIndex        =   24
      Top             =   2400
      Visible         =   0   'False
      Width           =   3135
      URL             =   "1.mp3"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   5530
      _cy             =   873
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   15000
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   13800
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label zimu 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   60
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   9
      Left            =   9120
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label zimu 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   60
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   8
      Left            =   8160
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label zimu 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   60
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   7
      Left            =   7200
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label zimu 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   60
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   6
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label zimu 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   60
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   5
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label zimu 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   60
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   4
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label zimu 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   60
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   3
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label zimu 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   60
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label zimu 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   60
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label zimu 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   60
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'洋子，2012年
Option Base 0

Dim speed%     '设置一个字母1ms下落的长度
Dim a%     '游戏状态
Dim mypoint   '得分
Dim pb%       '得分基数
Dim p%      '正确个数
Dim lp%    '失误个数
Dim ps    '每次加速在上次速度基础上的加速倍数
Dim pj%   '设置加速阶梯,正确个数
Dim t(9) As Boolean
Dim tim%    '设置一轮游戏时长，秒

Private Type points
    p As Long
End Type

Dim pmax As points


Sub reset()
    
    speed = 20
    a = 0
    p = 0
    lp = 0
    pb = 10
    ps = 1.5
    pj = 5
    tim = 20
 '在上面配置游戏
    
    
    
    time.Caption = Str(tim)
    
    Open "max.dat" For Random As #2 Len = Len(pmax)
    Get #2, , pmax
    Close #2
    
    Label6.Caption = Str(pmax.p)
    Command1.Caption = "开始(F2)"
    mypoint = 0
    Timer1.Enabled = False
    For i = 0 To 9
        zimuxialuo(i).Enabled = False
        zimuxialuo(i).Interval = 15
        zimu(i).Visible = False
        zimu(i).Top = 120
    Next
End Sub







Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Command1_Click
    If KeyCode = vbKeyF5 Then Command2_Click
    If a = 1 Then
        If KeyCode >= Asc("A") And KeyCode <= Asc("Z") Then check KeyCode
    End If
End Sub

Private Sub Form_Load()
    Line1.Y1 = 120
    Line1.Y2 = Line1.Y1
    Line1.X2 = Form1.Width
    Line2.X2 = Form1.Width
    reset

End Sub



Private Sub Timer1_Timer()
    tim = tim - 1
    time.Caption = Str(tim)
    If tim <= 0 Then
    endgame
    End If
End Sub

Private Sub zimuxialuo_Timer(Index As Integer)
    zimu(Index).Top = zimu(Index).Top + speed
    If zimu(Index).Top >= Line2.Y1 - zimu(Index).Height Then
        nextone Index
        lp = lp + 1
    Label5.Caption = Str(lp)
    End If
End Sub





Sub start()
    Dim cl As Variant
    cl = Array(vbRed, vbBlue, vbBlack, vbGreen)
    Randomize
    i = Int(Rnd * 10)
    Randomize
    j = Int(Rnd * 4)
    Randomize
    k = Int(Rnd * 58) + 65
    If k > 90 And k < 97 Then k = Int(Rnd * 26) + 65
    Timer1.Enabled = True
    zimu(i).ForeColor = cl(j)
    zimu(i).Caption = Chr(k)
    zimu(i).Visible = True
    zimuxialuo(i).Enabled = True

End Sub



Sub pause()
    Timer1.Enabled = False
    For i = 0 To 9
        t(i) = zimuxialuo(i).Enabled
        zimuxialuo(i).Enabled = False
    Next
End Sub
Sub nextone(ByVal a%)
    zimuxialuo(a).Enabled = False
    zimu(a).Visible = False
    zimu(a).Top = 120
    start

End Sub

Sub check(ByVal a%)
    For i% = 0 To 9
        If zimu(i).Visible And a = Asc(UCase(zimu(i).Caption)) Then
            p = p + 1
            mypoint = p * pb
            mu1.Controls.play
            If p Mod pj = 0 Then speed = speed * ps
            nextone i
            Exit For
        End If
    Next
    If i > 9 Then
    mu2.Controls.play
    lp = lp + 1
    End If
    Label4.Caption = Str(mypoint)
    Label5.Caption = Str(lp)
End Sub


Sub endgame()

    pause
    reset
    
    Open "max.dat" For Random As #1 Len = Len(pmax)
    mypoint = Int(Label4.Caption)
    
    If mypoint > pmax.p Then
        pmax.p = mypoint
        Put #1, , pmax
        MsgBox "恭喜创造新纪录！" & vbCrLf & "你的最后得分为" & Label4.Caption & vbCrLf & "共失误" & Label5.Caption & "次！", vbOKOnly, "游戏结束！"
    Else
        MsgBox "你的最后得分为" & Label4.Caption & vbCrLf & "共失误" & Label5.Caption & "次！", vbOKOnly, "游戏结束！"
    
    End If
    
    Close #1
 
    reset

End Sub





Private Sub Command1_Click()
    a = a * -1
    If a = 1 Then
        Command1.Caption = "暂停(F2)"
        For i = o To 9
            zimuxialuo(i).Enabled = t(i)
        Next
    ElseIf a = -1 Then
        Command1.Caption = "开始(F2)"
        pause
    Else
        Command1.Caption = "暂停(F2)"
        start
        a = 1
    End If
End Sub


Private Sub Command2_Click()
    If a = 1 Then
    a = -1
    Command1.Caption = "开始(F2)"
    pause
    End If
    If MsgBox("确定要重新开始？", vbOKCancel, "提示：") = vbOK Then reset
End Sub

Private Sub Command3_Click()
    If a = 1 Then
    a = -1
    Command1.Caption = "开始(F2)"
    pause
    End If
    MsgBox "作者：洋子。2012", vbOKOnly, "关于："
End Sub



Private Sub Command4_Click()
    If a = 1 Then
    a = -1
    Command1.Caption = "开始(F2)"
    pause
    End If
    If MsgBox("确定要退出？", vbOKCancel, "提示：") = vbOK Then End
End Sub


Private Sub Command5_Click()
    If a = 1 Then
        a = -1
        Command1.Caption = "开始(F2)"
        pause
    End If
    If MsgBox("确定要清除历史最高分？", vbOKCancel, "提示：") = vbOK Then
        pmax.p = 0
       Open "max.dat" For Random As #3 Len = Len(pmax)
       Put #3, , pmax
    Label6.Caption = Str(pmax.p)

       Close #3
    End If
End Sub

