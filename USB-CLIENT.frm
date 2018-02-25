VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form StudentMain 
   BorderStyle     =   0  'None
   Caption         =   "信息课管理系统"
   ClientHeight    =   4530
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   8370
   LinkTopic       =   "Form2"
   ScaleHeight     =   4530
   ScaleWidth      =   8370
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   840
      Top             =   1680
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Text            =   "输出"
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Text            =   "输入"
      Top             =   960
      Width           =   3375
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重启"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   480
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "StudentMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lockm As String
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Sub Command1_Click()
 Unload Me
End Sub
Private Sub Command2_Click()
 Shell "shutdown -i"
End Sub

Private Sub Command3_Click()
 StudentLock.Show
End Sub
Private Sub Form_Load()
 Winsock1.Protocol = sckUDPProtocol
 Winsock2.Protocol = sckUDPProtocol
 Winsock1.LocalPort = 23333
 Winsock1.Bind 23333
 Command1.Visible = True
 If Dir("F:\", vbDirectory) = "" Then
   Else
  Dim enter1
  enter1 = "检测到有U盘插入，请迅速拔出！拔出后请重启计算机。"
 End If
 If enter1 = "" Then
  Else
   Dim enter1m
    enter1m = "插入了U盘。"
    MsgBox (enter1m)
    Text3.Text = enter1m
    Command1.Visible = False
    Command2.Visible = True
    Winsock2.RemotePort = 23333
    Winsock2.RemoteHost = "172.17.30.99"
    Winsock2.SendData "text" & Winsock1.LocalIP & "试图插入U盘。"
 End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ReleaseCapture
 SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Winsock2.RemotePort = 23333
 Winsock2.RemoteHost = "172.17.30.99"
 Winsock2.SendData "text" & Winsock1.LocalIP & "试图退出程序。"
End Sub
Private Sub Timer1_Timer()
 Winsock2.RemotePort = 23333
 Winsock2.RemoteHost = "127.0.0.1"
 Winsock2.SendData "unlock" & "测试"
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
 Dim got As String
  Winsock1.GetData got, vbString
  If left(got, 4) = "text" Then
   Text1.Text = Mid(got, 5)
  End If
  If left(got, 9) = "broadcast" Then
   Text1.Text = Mid(got, 10)
  End If
  If left(got, 4) = "lock" Then
   lockm = Mid(got, 5)
   StudentLock.Show
  End If
  If left(got, 6) = "unlock" Then
   Unload StudentLock
   MsgBox Mid(got, 7), , "已解锁"
  End If
  If left(got, 5) = "shell" Then
   Shell Mid(got, 6)
  End If
End Sub
