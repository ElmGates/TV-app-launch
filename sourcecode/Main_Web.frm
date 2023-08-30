VERSION 5.00
Begin VB.Form Main_Web 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "主程序-应用启动台"
   ClientHeight    =   10635
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10635
   ScaleWidth      =   19080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame name4 
      Caption         =   "name4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   14520
      TabIndex        =   17
      Top             =   480
      Width           =   4335
      Begin VB.PictureBox img4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   360
         Picture         =   "Main_Web.frx":0000
         ScaleHeight     =   3825
         ScaleWidth      =   3825
         TabIndex        =   18
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame name1 
      Caption         =   "name1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   240
      TabIndex        =   15
      Top             =   480
      Width           =   4335
      Begin VB.PictureBox img1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   360
         Picture         =   "Main_Web.frx":2223
         ScaleHeight     =   3825
         ScaleWidth      =   3825
         TabIndex        =   16
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame name2 
      Caption         =   "name2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   5040
      TabIndex        =   13
      Top             =   480
      Width           =   4335
      Begin VB.PictureBox img2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   360
         Picture         =   "Main_Web.frx":4446
         ScaleHeight     =   3825
         ScaleWidth      =   3825
         TabIndex        =   14
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame name3 
      Caption         =   "name3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   9840
      TabIndex        =   11
      Top             =   480
      Width           =   4335
      Begin VB.PictureBox img3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   360
         Picture         =   "Main_Web.frx":6669
         ScaleHeight     =   3825
         ScaleWidth      =   3825
         TabIndex        =   12
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame name5 
      Caption         =   "name5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      Width           =   4335
      Begin VB.PictureBox img5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   360
         Picture         =   "Main_Web.frx":888C
         ScaleHeight     =   3825
         ScaleWidth      =   3825
         TabIndex        =   10
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame name6 
      Caption         =   "name6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   4920
      TabIndex        =   7
      Top             =   5160
      Width           =   4335
      Begin VB.PictureBox img6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   360
         Picture         =   "Main_Web.frx":AAAF
         ScaleHeight     =   3825
         ScaleWidth      =   3825
         TabIndex        =   8
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame name7 
      Caption         =   "name7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   9720
      TabIndex        =   5
      Top             =   5160
      Width           =   4335
      Begin VB.PictureBox img7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   360
         Picture         =   "Main_Web.frx":CCD2
         ScaleHeight     =   3825
         ScaleWidth      =   3825
         TabIndex        =   6
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "APP启动台"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   18975
      Begin VB.Frame name8 
         Caption         =   "name8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   14520
         TabIndex        =   3
         Top             =   5160
         Width           =   4335
         Begin VB.PictureBox img8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   3855
            Left            =   360
            Picture         =   "Main_Web.frx":EEF5
            ScaleHeight     =   3825
            ScaleWidth      =   3825
            TabIndex        =   4
            Top             =   360
            Width           =   3855
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出启动台"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   1
      Top             =   9720
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11880
      TabIndex        =   0
      Top             =   9720
      Width           =   3855
   End
End
Attribute VB_Name = "Main_Web"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim csStr() As String
Dim img() As String

Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim strUrl As String
     
Private Sub OpenUrl(tUrl As String)
    ShellExecute Me.hwnd, "Open", tUrl, 0, 0, 0
End Sub
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Dim ii As Integer
ii = MsgBox("暂时还未完成设置的开发，将带您前往设置页面，如果不需要修改设置，请直接退出重起", vbOKOnly, "Error")
Frist_Run.Show
Unload Me
End Sub

Private Sub Form_Load()
Dim iStr() As String, i As Long, readconfigpath As String, pathini As String, config As String
i = 1
readconfigpath = App.Path + "\config.ini"
Open readconfigpath For Input As #2
Line Input #2, config
Close #2
If Mid(config, 3, 1) = "0" Then
pathini = "\con_path.ini"
ElseIf Mid(config, 3, 1) = "1" Then
pathini = "\app_path.ini"
ElseIf Mid(config, 3, 1) = "2" Then
pathini = "\web_path.ini"
End If
Open App.Path + pathini For Input As #1
Do While Not EOF(1)
   ReDim Preserve csStr(i)
   Line Input #1, csStr(i)
   i = i + 1
Loop
Close #1
i = 1
Open App.Path + "\appname.ini" For Input As #1
Do While Not EOF(1)
   ReDim Preserve iStr(i)
   Line Input #1, iStr(i)
   i = i + 1
Loop
Close #1
i = 1
Open App.Path + "\img_path.ini" For Input As #1
Do While Not EOF(1)
   ReDim Preserve img(i)
   Line Input #1, img(i)
   i = i + 1
Loop
Close #1
name1.Caption = iStr(1)
name2.Caption = iStr(2)
name3.Caption = iStr(3)
name4.Caption = iStr(4)
name5.Caption = iStr(5)
name6.Caption = iStr(6)
name7.Caption = iStr(7)
name8.Caption = iStr(8)
If img(1) = "" Then
img1.Picture = LoadPicture(App.Path & "\img\default.jpg")
Else
img1.Picture = LoadPicture(img(1))
End If
If img(2) = "" Then
img2.Picture = LoadPicture(App.Path & "\img\default.jpg")
Else
img2.Picture = LoadPicture(img(2))
End If
If img(3) = "" Then
img3.Picture = LoadPicture(App.Path & "\img\default.jpg")
Else
img3.Picture = LoadPicture(img(3))
End If
If img(4) = "" Then
img4.Picture = LoadPicture(App.Path & "\img\default.jpg")
Else
img4.Picture = LoadPicture(img(4))
End If
If img(5) = "" Then
img5.Picture = LoadPicture(App.Path & "\img\default.jpg")
Else
img5.Picture = LoadPicture(img(5))
End If
If img(6) = "" Then
img6.Picture = LoadPicture(App.Path & "\img\default.jpg")
Else
img6.Picture = LoadPicture(img(6))
End If
If img(7) = "" Then
img7.Picture = LoadPicture(App.Path & "\img\default.jpg")
Else
img7.Picture = LoadPicture(img(7))
End If
If img(8) = "" Then
img8.Picture = LoadPicture(App.Path & "\img\default.jpg")
Else
img8.Picture = LoadPicture(img(8))
End If
End Sub

Private Sub img1_Click()
OpenUrl (csStr(1))
End Sub

Private Sub img2_Click()
OpenUrl (csStr(2))
End Sub

Private Sub img3_Click()
OpenUrl (csStr(3))
End Sub

Private Sub img4_Click()
OpenUrl (csStr(4))
End Sub

Private Sub img5_Click()
OpenUrl (csStr(5))
End Sub

Private Sub img6_Click()
OpenUrl (csStr(6))
End Sub

Private Sub img7_Click()
OpenUrl (csStr(7))
End Sub

Private Sub img8_Click()
OpenUrl (csStr(8))
End Sub


