VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form TeacherGetFile 
   Caption         =   "Form6"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2760
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   2040
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   360
      Picture         =   "TeacherGetClip.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "发送"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "此窗体暂时停用。"
      Height          =   735
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "TeacherGetFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'暂时停用
