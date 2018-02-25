VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ClientFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "客户端"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3720
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   120
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "文件下载"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "接收进度:"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "服务器地址"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "ClientFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FName As String         '将接收的文件完整路径
Private Sub Command1_Click()
    If Winsock1.State <> sckClosed Then   '如果Winsock1当前状态非关闭
        Winsock1.Close                    '关闭连接
    End If
    Winsock1.RemoteHost = Text1.Text      '服务器地址
    Winsock1.RemotePort = 4567            '服务器端口
    Winsock1.Connect                      '连接
End Sub
 Private Sub Form_Load()
    FName = App.Path & "\" & Text2.Text      '指定接收文件完整路径
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    Winsock1.Close   '关闭连接
End Sub
 Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then     '如果当前输入的是回车键
        KeyAscii = 0          '取消当前输入
        Command1_Click        '激发Command1_Click过程
    End If
End Sub
 Private Sub Winsock1_Close()
    Winsock1.Close   '关闭连接
End Sub
 
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim TheFile() As Byte                    '接受数据的数组
    ReDim TheFile(bytesTotal)                '重定义数组下界
    Static YNLen As Boolean                  '是否接收了文件长度
    Winsock1.GetData TheFile                 '将接收的数据保存到数组
    If bytesTotal = 2 And Chr(TheFile(0)) = "C" And Chr(TheFile(1)) = "S" Then    '如果收到的是成功连接信息
        Me.Caption = "客户端-----成功连接"    '提示信息
        Winsock1.SendData "GetFileLen"        '发送要求文件长度信息
        Text1.Enabled = False                 'Text1不可输入
        Command1.Enabled = False              '按钮不可用
        ProgressBar1.Value = 0                '进度条当前值为0
        Exit Sub          '结束过程
    End If
    If bytesTotal = 3 And Chr(TheFile(0)) = "N" And Chr(TheFile(1)) = "o" And Chr(TheFile(2)) = "F" Then '如果收到的是无此文件的信息
        MsgBox "服务器并无此文件。", vbInformation, "提示"        '报错提示
        Exit Sub     '结束过程
    End If
    If bytesTotal = 3 And Chr(TheFile(0)) = "T" And Chr(TheFile(1)) = "h" And Chr(TheFile(2)) = "E" Then  '如果收到文件传送结束信息
        Close #1      '关闭文件
        YNLen = False '未接收文件长度描述信息
        Me.Caption = "客户端-----文件已成功接收"   '提示信息
        Winsock1.SendData "ConClose"    '关闭连接
        Text1.Enabled = True            'Text1可输入
        Command1.Enabled = True         '按钮可用
        Exit Sub
    End If
    If YNLen = True Then   '如果已经接收了文件长度信息
        Put #1, , TheFile                '将接收的数据包写入该文件
        Winsock1.SendData "NextB"        '发送要求下一数据包的信息
        ProgressBar1.Value = ProgressBar1.Value + bytesTotal '接收文件进度
    Else
        Me.Caption = "客户端-----正在接收数据"     '提示信息
        Dim i As Integer
        Dim Strs As String   '描述文件长度字符串
        For i = 0 To bytesTotal - 1
            Strs = Strs & Chr(TheFile(i))   '组合文件长度描述字符串
        Next i
        ProgressBar1.Max = Val(Strs)   '设置进度条最大值
        ProgressBar1.Min = 0           '设置进度条最小值
        YNLen = True                   '已经接收了文件长度描述信息
        Winsock1.SendData "FLA"        '发送已经收到文件长度描述信息的信息"FLA"

    End If
End Sub
 
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbInformation, "错误号:" & Number   '出错提示
    Winsock1.Close   '关闭连接
End Sub
