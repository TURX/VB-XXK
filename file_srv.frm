VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ServerFrm 
   Caption         =   "服务器端-------正监听"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   4995
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   240
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "ServerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FName As String      '要传送的文件完整路径（包括文件名）

Private Sub Command1_Click()
ClientFrm.Show
End Sub
Private Sub Form_Load()
    'Download by http://www.newxing.com
    CClose
    FName = App.Path & "\319.jpg"   '指定完整文件路径
End Sub
 Private Sub Form_Unload(Cancel As Integer)
    Winsock1.Close     '关闭连接
End Sub
 Private Sub Winsock1_Close()
    CClose
    Close #1    '为预防当客户端非法关闭时造成不能及时关闭已打开的文件的情况
End Sub
 
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then   '如果Winsock1当前状态非关闭
        Winsock1.Close                    '关闭连接
    End If
    Winsock1.Accept requestID             '关闭了的Winsock1接受连接
    Winsock1.SendData "CS"           '向工作站发送握手信息（表明已成功连接）
    Me.Caption = "服务器端-------已连接"   '提示信息
End Sub
 
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim TStr As String         '已到达的字符串信息数据
    Static FileLen As Long     '文件长度（以字节为单位）
    Static CurPos As Long      '文件传送进度
    Dim S As Long              '下一个数据包的大小（用以确定数组的下界）
    Dim TheFile() As Byte      '定义一个可变大小的数组，用以传送二进制数据包
    Winsock1.GetData TStr      '接收以到达的字符串数据
    Select Case TStr
    Case "ConClose"            '如果收到的是关闭连接信息
        CClose
    Case "GetFileLen"          '如果收到的是要求传送文件长度信息
        If Dir(FName) <> "" Then          '如果文件存在
            Open FName For Binary As #1   '打开文件以获得该文件的长度
            FileLen = LOF(1)              '获得已打开文件的长度（以字节为单位）
            Winsock1.SendData Trim(Str(FileLen)) '发送文件长度
        Else
            Winsock1.SendData "NoF"     '发送未找到文件信息
        End If
    Case "FLA"                            '如果收到的是文件长度已到达信息
        ReDim TheFile(1 To 3072)          '重定义数组下界为3072
        Get #1, , TheFile                 '获得已打开文件的一部分字节（3072个），并将其保存在TheFile数组中
        Winsock1.SendData TheFile         '发送第一个数据包
        CurPos = 3072                     '文件当前已读取了3072个字节
    Case "NextB"                          '如果收到的是要求下一个数据包的信息
        If CurPos = FileLen Then          '如果已经传送了所有的字节
            Winsock1.SendData "ThE"       '发送文件传送结束信息
            Close #1                      '关闭文件
            Exit Sub                      '结束过程
        End If
        S = CurPos + 3072                 '当前文件传送进度加3072个字节
        If S > FileLen Then               '如果超过了文件总长度
            S = FileLen - CurPos          '下一个数组下界应为余下的字节数
        Else
            S = 3072                      '下一个数组下界仍为3072
        End If
        ReDim TheFile(1 To S)             '重定义数组下界为S
        Get #1, , TheFile                 '读取该文件的下一部分数据，并保存在TheFile数组中
        Winsock1.SendData TheFile         '发送该包
        CurPos = CurPos + S               '确定传送进度
    End Select
End Sub
 
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbInformation, "错误号:" & Number   '出错提示
    CClose
End Sub
Private Sub CClose()   '关闭当前连接，并继续监听
    Winsock1.Close     '关闭当前连接
    Winsock1.LocalPort = 4567      '设置监听端口号为4567
    Winsock1.Listen                '开始监听
    Me.Caption = "服务器端-------连接已断开,现正监听" '提示信息
End Sub
