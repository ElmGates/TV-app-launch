VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form imgpath 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "选择应用显示图标，否则则显示默认图标，可能会导致严重报错！"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   14295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   14295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Caption         =   "应用1"
      Height          =   855
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   6975
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "应用2"
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   960
      Width           =   6975
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "应用3"
      Height          =   855
      Left            =   0
      TabIndex        =   13
      Top             =   1920
      Width           =   6975
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "应用4"
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   2880
      Width           =   6975
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "应用8"
      Height          =   855
      Left            =   7200
      TabIndex        =   9
      Top             =   2880
      Width           =   6975
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "应用5"
      Height          =   855
      Left            =   7200
      TabIndex        =   7
      Top             =   0
      Width           =   6975
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "应用6"
      Height          =   855
      Left            =   7200
      TabIndex        =   5
      Top             =   960
      Width           =   6975
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "应用7"
      Height          =   855
      Left            =   7200
      TabIndex        =   3
      Top             =   1920
      Width           =   6975
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6735
      End
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
      Left            =   0
      TabIndex        =   2
      Top             =   3960
      Width           =   4215
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
      Left            =   9840
      TabIndex        =   1
      Top             =   3960
      Width           =   4335
   End
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
      Left            =   4800
      TabIndex        =   0
      Top             =   3960
      Width           =   4335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6840
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "imgpath"
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
pathini = "\img_path.ini"
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
xx = MsgBox("请注意：图片大小必须是200*200像素，24位深以下，否则会导致运行错误，如果您没有符合的图片，请留白即使用默认图片", vbOKOnly, "重要提示！！！")
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
mm = MsgBox("已保存，请重启软件", vbOKOnly, "需要重启")
End
End Sub

Private Sub Text1_Click()
 ' 设置“CancelError”为 True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' 设置标志
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' 设置过滤器
    CommonDialog1.Filter = "All Files (*.*)|*.*|jpg" & _
    "(*.jpg)|*.jpg"
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
    CommonDialog1.Filter = "All Files (*.*)|*.*|jpg" & _
    "(*.jpg)|*.jpg"
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
    CommonDialog1.Filter = "All Files (*.*)|*.*|jpg" & _
    "(*.jpg)|*.jpg"
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
    CommonDialog1.Filter = "All Files (*.*)|*.*|jpg" & _
    "(*.jpg)|*.jpg"
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
    CommonDialog1.Filter = "All Files (*.*)|*.*|jpg" & _
    "(*.jpg)|*.jpg"
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
    CommonDialog1.Filter = "All Files (*.*)|*.*|jpg" & _
    "(*.jpg)|*.jpg"
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
    CommonDialog1.Filter = "All Files (*.*)|*.*|jpg" & _
    "(*.jpg)|*.jpg"
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
    CommonDialog1.Filter = "All Files (*.*)|*.*|jpg" & _
    "(*.jpg)|*.jpg"
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






