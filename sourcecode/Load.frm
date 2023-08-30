VERSION 5.00
Begin VB.Form Load 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Loading……"
   ClientHeight    =   465
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Label1 
      Caption         =   "读取信息..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
configpath = App.Path + "\config.ini"
Open configpath For Input As #2
Line Input #2, config
Close #2
If Mid(config, 1, 1) = "0" Then
ii = MsgBox("程序还未初始化，即将前往初始化程序", vbOKOnly, "需要初始化")
Frist_Run.Show
Unload Me
ElseIf Mid(config, 1, 1) = "1" Then
If Mid(config, 3, 1) = "0" Then
Main_Combine.Show
Unload Me
ElseIf Mid(config, 3, 1) = "1" Then
Main_Apps.Show
Unload Me
ElseIf Mid(config, 3, 1) = "2" Then
Main_Web.Show
Unload Me
End If
End If
End Sub
