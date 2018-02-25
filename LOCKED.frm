VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form StudentLock 
   BorderStyle     =   0  'None
   Caption         =   "Lock"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13680
   LinkTopic       =   "Form4"
   ScaleHeight     =   6360
   ScaleWidth      =   13680
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3480
      Top             =   3960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   3960
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   8055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "StudentLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Type RECT
 left As Long
 top As Long
 right As Long
 bottom As Long
End Type '以上代码请从API函数浏览器中复制即可。
Dim DENG As RECT
Dim SS As Boolean '是锁定或是释放的判断
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Const SPI_SCREENSAVERRUNNING = 97
Private Sub Form_LostFocus()
 Me.SetFocus
End Sub
Private Sub Timer1_Timer()
 hw = FindWindow(vbNullString, "任务管理器")
 hw = FindWindow(vbNullString, "Windows 任务管理器")
 SendMessage hw, &H10, 0, 0
 Me.SetFocus
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
 If SS = False Then
  Command1_Click
 End If
End Sub
Private Sub Form_Load() '作初始化设置。
 SS = True
 Command1.Caption = "LOCK"
 Command2.Caption = "END"
 Label1.Caption = "请认真听讲"
 Label2.Caption = StudentMain.lockm
 Winsock1.Protocol = sckUDPProtocol
 Winsock1.LocalPort = 23334
 Winsock1.Bind 23334
 Dim r As Integer
 Dim p As Boolean
 r = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, p, 0)
 Command1_Click
 Shell "taskkill /im explorer.exe /f"
End Sub
Private Sub Command1_Click() '锁定鼠标
 If SS = True Then '如果鼠标没有被锁定则锁定它。
  DENG.left = 0: DENG.top = 0 '必要设置,四个都为零。
  DENG.right = 0: DENG.bottom = 0
  ClipCursor DENG: SS = False  '锁定鼠标且重新赋值给SS。
    Else
  ClipCursor ByVal 0&: SS = True   '释放鼠标且重新赋值给SS
 End If
End Sub
Private Sub Command2_Click()
 Shell App.Path & "\res\explorer\RestartExplorer.exe"
 Unload Me '结束程序。
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Command1_Click
End Sub
