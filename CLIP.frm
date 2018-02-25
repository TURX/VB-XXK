VERSION 5.00
Begin VB.Form StudentClip 
   Caption         =   "Form5"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton StudentClipScreen 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   600
      Top             =   840
      Width           =   3255
   End
End
Attribute VB_Name = "StudentClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Sub Command1_Click()
 Call ScrnCap(0, 0, 800, 600) '调用函数，4个参数为左上，右下坐标
 Image1.Picture = Clipboard.GetData()
End Sub
Sub ScrnCap(Lt As Integer, top As Integer, Rt As Integer, Bot As Integer) '屏幕截图核心函数
 Dim rWidth, rHeight, SourceDC, DestDC, BHandle, Wnd, DHandle
 rWidth = Rt - Lt
 rHeight = Bot - top
 SourceDC = CreateDC("DISPLAY", 0, 0, 0)
 DestDC = CreateCompatibleDC(SourceDC)
 BHandle = CreateCompatibleBitmap(SourceDC, rWidth, rHeight)
 SelectObject DestDC, BHandle
 BitBlt DestDC, 0, 0, rWidth, rHeight, SourceDC, Lt, top, &HCC0020
 Wnd = Screen.ActiveForm.hwnd
OpenClipboard Wnd
 EmptyClipboard
 SetClipboardData 2, BHandle
 CloseClipboard
 DeleteDC DestDC
 ReleaseDC DHandle, SourceDC
End Sub
