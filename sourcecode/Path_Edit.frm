VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Path_Edit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "修改路径-暂时只支持8个-点击文本框来选择"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   14460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   14460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton clear 
      Caption         =   "清空"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   18
      Top             =   4080
      Width           =   4335
   End
   Begin VB.CommandButton exit001 
      Caption         =   "不保存并退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      TabIndex        =   17
      Top             =   4080
      Width           =   4335
   End
   Begin VB.CommandButton save01 
      Caption         =   "保存"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6960
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame8 
      Caption         =   "路径7"
      Height          =   855
      Left            =   7320
      TabIndex        =   14
      Top             =   2040
      Width           =   6975
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "路径6"
      Height          =   855
      Left            =   7320
      TabIndex        =   12
      Top             =   1080
      Width           =   6975
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "路径5"
      Height          =   855
      Left            =   7320
      TabIndex        =   10
      Top             =   120
      Width           =   6975
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "路径8"
      Height          =   855
      Left            =   7320
      TabIndex        =   8
      Top             =   3000
      Width           =   6975
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "路径4"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   6975
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "路径3"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   6975
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "路径2"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   6975
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "路径1"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
      End
   End
End
Attribute VB_Name = "Path_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim readconfigpath As String
Dim MonitorSetFile As String
Dim pathini As String

Private Sub clear_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub

Private Sub exit001_Click()
mm = MsgBox("已取消，请重启软件", vbOKOnly, "需要重启")
End
End Sub

Private Sub Form_Load()
Dim iStr() As String, i As Long, c
i = 1
readconfigpath = App.Path + "\config.ini"
Open readconfigpath For Input As #2
Line Input #2, config
Close #2
If Mid(config, 3, 1) = "0" Then
pathini = "\con_path.ini"
xx = MsgBox("网页链接请以http开头，否则应用会出现错误", vbOKOnly, "请您注意！")
ElseIf Mid(config, 3, 1) = "1" Then
pathini = "\app_path.ini"
ElseIf Mid(config, 3, 1) = "2" Then
pathini = "\web_path.ini"
End If
Open App.Path + pathini For Input As #1
Do While Not EOF(1)
   ReDim Preserve iStr(i)
   Line Input #1, iStr(i)
   i = i + 1
Loop
Close #1
Text1.Text = iStr(1)
Text2.Text = iStr(2)
Text3.Text = iStr(3)
Text4.Text = iStr(4)
Text5.Text = iStr(5)
Text6.Text = iStr(6)
Text7.Text = iStr(7)
Text8.Text = iStr(8)
End Sub

Private Sub save01_Click()
MonitorSetFile = App.Path + pathini
 Dim ThisInst As String
 Open MonitorSetFile For Output As #1
 Print #1, Text1.Text
 Print #1, Text2.Text
 Print #1, Text3.Text
 Print #1, Text4.Text
 Print #1, Text5.Text
 Print #1, Text6.Text
 Print #1, Text7.Text
 Print #1, Text8.Text
 Close #1
mm = MsgBox("已保存，请继续下一步配置操作", vbOKOnly, "请继续")
Unload Me
End Sub

Private Sub Text1_Click()
 ' 设置“CancelError”为 True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' 设置标志
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' 设置过滤器
    CommonDialog1.Filter = "All Files (*.*)|*.*|应用程序" & _
    "(*.exe)|*.exe"
    ' 指定缺省的过滤器
    CommonDialog1.FilterIndex = 2
    ' 显示“打开”对话框
    CommonDialog1.ShowOpen
    ' 显示选定文件的名字
    Text1.Text = CommonDialog1.FileName '显示路径
    Exit Sub
ErrHandler:
    ' 用户按了“取消”按钮
    Exit Sub
End Sub
Private Sub Text2_Click()
 ' 设置“CancelError”为 True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' 设置标志
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' 设置过滤器
    CommonDialog1.Filter = "All Files (*.*)|*.*|应用程序" & _
    "(*.exe)|*.exe"
    ' 指定缺省的过滤器
    CommonDialog1.FilterIndex = 2
    ' 显示“打开”对话框
    CommonDialog1.ShowOpen
    ' 显示选定文件的名字
    Text2.Text = CommonDialog1.FileName '显示路径
    Exit Sub
ErrHandler:
    ' 用户按了“取消”按钮
    Exit Sub
End Sub
Private Sub Text3_Click()
 ' 设置“CancelError”为 True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' 设置标志
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' 设置过滤器
    CommonDialog1.Filter = "All Files (*.*)|*.*|应用程序" & _
    "(*.exe)|*.exe"
    ' 指定缺省的过滤器
    CommonDialog1.FilterIndex = 2
    ' 显示“打开”对话框
    CommonDialog1.ShowOpen
    ' 显示选定文件的名字
    Text3.Text = CommonDialog1.FileName '显示路径
    Exit Sub
ErrHandler:
    ' 用户按了“取消”按钮
    Exit Sub
End Sub
Private Sub Text4_Click()
 ' 设置“CancelError”为 True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' 设置标志
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' 设置过滤器
    CommonDialog1.Filter = "All Files (*.*)|*.*|应用程序" & _
    "(*.exe)|*.exe"
    ' 指定缺省的过滤器
    CommonDialog1.FilterIndex = 2
    ' 显示“打开”对话框
    CommonDialog1.ShowOpen
    ' 显示选定文件的名字
    Text4.Text = CommonDialog1.FileName '显示路径
    Exit Sub
ErrHandler:
    ' 用户按了“取消”按钮
    Exit Sub
End Sub
Private Sub Text5_Click()
 ' 设置“CancelError”为 True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' 设置标志
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' 设置过滤器
    CommonDialog1.Filter = "All Files (*.*)|*.*|应用程序" & _
    "(*.exe)|*.exe"
    ' 指定缺省的过滤器
    CommonDialog1.FilterIndex = 2
    ' 显示“打开”对话框
    CommonDialog1.ShowOpen
    ' 显示选定文件的名字
    Text5.Text = CommonDialog1.FileName '显示路径
    Exit Sub
ErrHandler:
    ' 用户按了“取消”按钮
    Exit Sub
End Sub
Private Sub Text6_Click()
 ' 设置“CancelError”为 True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' 设置标志
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' 设置过滤器
    CommonDialog1.Filter = "All Files (*.*)|*.*|应用程序" & _
    "(*.exe)|*.exe"
    ' 指定缺省的过滤器
    CommonDialog1.FilterIndex = 2
    ' 显示“打开”对话框
    CommonDialog1.ShowOpen
    ' 显示选定文件的名字
    Text6.Text = CommonDialog1.FileName '显示路径
    Exit Sub
ErrHandler:
    ' 用户按了“取消”按钮
    Exit Sub
End Sub
Private Sub Text7_Click()
 ' 设置“CancelError”为 True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' 设置标志
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' 设置过滤器
    CommonDialog1.Filter = "All Files (*.*)|*.*|应用程序" & _
    "(*.exe)|*.exe"
    ' 指定缺省的过滤器
    CommonDialog1.FilterIndex = 2
    ' 显示“打开”对话框
    CommonDialog1.ShowOpen
    ' 显示选定文件的名字
    Text7.Text = CommonDialog1.FileName '显示路径
    Exit Sub
ErrHandler:
    ' 用户按了“取消”按钮
    Exit Sub
End Sub
Private Sub Text8_Click()
 ' 设置“CancelError”为 True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' 设置标志
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' 设置过滤器
    CommonDialog1.Filter = "All Files (*.*)|*.*|应用程序" & _
    "(*.exe)|*.exe"
    ' 指定缺省的过滤器
    CommonDialog1.FilterIndex = 2
    ' 显示“打开”对话框
    CommonDialog1.ShowOpen
    ' 显示选定文件的名字
    Text8.Text = CommonDialog1.FileName '显示路径
    Exit Sub
ErrHandler:
    ' 用户按了“取消”按钮
    Exit Sub
End Sub


