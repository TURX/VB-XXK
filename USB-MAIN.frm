VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form TeacherMain 
   BorderStyle     =   0  'None
   Caption         =   "信息课管理系统"
   ClientHeight    =   4545
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6600
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame5 
      Caption         =   "其它"
      Height          =   1695
      Left            =   2520
      TabIndex        =   19
      Top             =   2760
      Width           =   2415
      Begin VB.CommandButton Command7 
         Caption         =   "获取文件"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "OFF"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   18
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "执行"
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "执行命令"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "OFF"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ON"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Interval        =   5
      Left            =   6000
      Top             =   3960
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   5640
      Top             =   3960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ON"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4920
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "IP地址"
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "接收到的消息。"
      Top             =   960
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "发送"
      Height          =   375
      Left            =   1200
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "要发送的消息。"
      Top             =   1440
      Width           =   6375
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   5280
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00359F6A&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6615
      TabIndex        =   6
      Top             =   0
      Width           =   6615
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   6360
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "信息课管理系统"
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
         TabIndex        =   9
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label3 
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
         TabIndex        =   8
         Top             =   0
         Width           =   375
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
         TabIndex        =   7
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "锁定电脑"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "发送消息"
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Frame Frame4 
      Caption         =   "广播"
      Height          =   735
      Left            =   2520
      TabIndex        =   17
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Powered by (C) Crystal Computer Studio - TURX"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "TeacherMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ip3 As String
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Sub Command1_Click()
 Winsock2.RemotePort = 23333
 Winsock2.RemoteHost = ip3
 Winsock2.SendData "text" & Text1.Text
End Sub
Private Sub Command2_Click()
 ip3 = "255.255.255.255"
 Timer2.Enabled = False
 Command2.Enabled = False
 Command5.Enabled = True
 Text3.Enabled = False
End Sub
Private Sub Command3_Click()
 Winsock2.RemotePort = 23333
 Winsock2.RemoteHost = ip3
 Winsock2.SendData "lock" & Text1.Text
 Command4.Enabled = True
End Sub
Private Sub Command4_Click()
 Winsock2.RemotePort = 23333
 Winsock2.RemoteHost = ip3
 Winsock2.SendData "unlock" & Text1.Text
End Sub
Private Sub Command5_Click()
 Text3.Enabled = True
 ip3 = Text3.Text
 Timer2.Enabled = True
 Command2.Enabled = True
 Command5.Enabled = False
End Sub
Private Sub Command6_Click()
 Winsock2.RemotePort = 23333
 Winsock2.RemoteHost = ip3
 Winsock2.SendData "shell" & Text1.Text
End Sub
Private Sub Command7_Click()
 TeacherGetFile.Show
End Sub
Private Sub Form_Load()
 Text1.Text = ""
 Text2.Text = ""
 Text3.Text = "127.0.0.1"
 Winsock1.Protocol = sckUDPProtocol
 Winsock2.Protocol = sckUDPProtocol
 Winsock1.LocalPort = 23333
 Winsock1.Bind 23333
End Sub
Private Sub Label1_Click()
 About.Show
End Sub
Private Sub Label4_Click()
 About.Show
End Sub
Private Sub Label2_Click()
 Unload Me
End Sub
Private Sub Label3_Click()
 Me.WindowState = vbMinimized
End Sub
Private Sub Label5_Click()
 Unload Me
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label2.BackStyle = 1
 Label2.BackColor = &O0
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label3.BackStyle = 1
 Label3.BackColor = &O0
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label2.BackStyle = 1
 Label2.BackColor = &O0
End Sub
Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label2.BackStyle = 0
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label3.BackStyle = 0
End Sub
Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
Private Sub Timer1_Timer()
 Me.Caption = "信息课管理系统"
 Label4.Caption = "信息课管理系统"
End Sub
Private Sub Timer2_Timer()
 ip3 = Text3.Text
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
 Dim got As String
  Winsock1.GetData got, vbString
 If left(got, 4) = "text" Then
  Text2.Text = "消息 - " & Mid(got, 5)
  Me.Caption = "信息课管理系统 - 你有新消息"
  Label4.Caption = "信息课管理系统 - 你有新消息"
 End If
 If left(got, 9) = "broadcast" Then
  Text2.Text = "广播 - " & Mid(got, 10)
  Me.Caption = "信息课管理系统 - 你有新广播"
  Label4.Caption = "信息课管理系统 - 你有新广播"
 End If
 If left(got, 4) = "lock" Then
  Text2.Text = "已锁定 - " & Mid(got, 5)
  Me.Caption = "信息课管理系统 - 已锁定"
  Label4.Caption = "信息课管理系统 - 已锁定"
 End If
 If left(got, 6) = "unlock" Then
  Text2.Text = "已解锁 - " & Mid(got, 7)
  Me.Caption = "信息课管理系统 - 已解锁"
  Label4.Caption = "信息课管理系统 - 已解锁"
 End If
 If left(got, 5) = "shell" Then
  Text2.Text = "已执行 - " & Mid(got, 6)
  Me.Caption = "信息课管理系统 - 已执行"
  Label4.Caption = "信息课管理系统 - 已执行"
 End If
End Sub
