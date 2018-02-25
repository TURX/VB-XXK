VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   0  'None
   Caption         =   "关于我们 - About us"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   LinkTopic       =   "Form3"
   ScaleHeight     =   3720
   ScaleWidth      =   6615
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "技术支持"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00359F6A&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6615
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   6360
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5880
         TabIndex        =   3
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5520
         TabIndex        =   2
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "信息课管理系统 - 关于我们"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Copyright by (c) Crystal Studio - TURX"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   3720
      Picture         =   "About.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2760
   End
   Begin VB.Label Label6 
      Caption         =   "信息课管理系统"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   3615
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Sub Command1_Click()
 Support.Show
End Sub
Private Sub Form_Click()
 Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
 Unload Me
End Sub
Private Sub Label2_Click()
 Unload Me
End Sub
Private Sub Label1_Click()
 Me.WindowState = vbMinimized
End Sub
Private Sub Label3_Click()
 Unload Me
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label2.BackStyle = 1
 Label2.BackColor = &O0
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label1.BackStyle = 1
 Label1.BackColor = &O0
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label2.BackStyle = 1
 Label2.BackColor = &O0
End Sub
Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label2.BackStyle = 0
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label1.BackStyle = 0
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label2.BackStyle = 0
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ReleaseCapture
 SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ReleaseCapture
 SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
