VERSION 5.00
Begin VB.Form Frist_Run 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "应用初始化"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame4 
      Caption         =   "4.设置应用图标"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton Command4 
         Caption         =   "前往设置路径"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "3.设置应用名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton Command2 
         Caption         =   "前往设置路径"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "下一步"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      Caption         =   "2.设置路径"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton Command1 
         Caption         =   "前往设置路径"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "1.选择应用模式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.ComboBox chs 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Text            =   "混合网页和应用"
         Top             =   240
         Width           =   3375
      End
   End
End
Attribute VB_Name = "Frist_Run"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Path_Edit.Show
Frame3.Visible = True
End Sub

Private Sub Command2_Click()
appname.Show
Frame4.Visible = True
End Sub

Private Sub Command3_Click()
If chs.Text = "混合网页和应用" Then
xieru = "1,0"
ElseIf chs.Text = "纯应用" Then
xieru = "1,1"
ElseIf chs.Text = "纯网页" Then
xieru = "1,2"
End If
File0 = App.Path + "\config.ini"
Open File0 For Output As #1
Print #1, xieru
Close #1
Frame2.Visible = True
End Sub

Private Sub Command4_Click()
imgpath.Show
Unload Me
End Sub

Private Sub Form_Load()
chs.AddItem ("混合网页和应用")
chs.AddItem ("纯应用")
chs.AddItem ("纯网页")
End Sub
