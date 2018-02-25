VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3960
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   2640
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1680
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1800
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   600
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long) As Boolean
Dim bytData() As Byte
Private Sub Command1_Click()
    Winsock1.RemoteHost = Text1.Text
    Dim arr() As Byte
    Dim i As New PropertyBag
    i.WriteProperty "image", Picture1.Picture
    ReDim arr(1 To LenB(i.Contents))
    arr = i.Contents
    If UBound(arr) <= 8192 Then '如果要发送的文件小于数据块大小，直接发送
        Winsock1.SendData arr '发送数据
        Exit Sub
    End If
End Sub
Private Sub Form_Load()
    With Winsock1    '信息发送与接收
        .Protocol = sckUDPProtocol '使用UDP协议
        .RemotePort = 9001 '要连接的端口
        .LocalPort = 9001
        .Bind  '绑定到本地的端口上
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Erase bytData
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim arr() As Byte
    ReDim arr(1 To bytesTotal)
    Winsock1.GetData arr
    ReDim Preserve bytData(1 To bytesTotal)
    CopyMemory bytData(1), arr(0), bytesTotal
    Dim i As New PropertyBag
    i.Contents = bytData
    Picture2.Picture = i.ReadProperty("image")
End Sub
