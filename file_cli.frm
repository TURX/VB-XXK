VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ClientFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ͻ���"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3720
   StartUpPosition =   3  '����ȱʡ
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
      Caption         =   "�ļ�����"
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
      Caption         =   "���ս���:"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��������ַ"
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
Private FName As String         '�����յ��ļ�����·��
Private Sub Command1_Click()
    If Winsock1.State <> sckClosed Then   '���Winsock1��ǰ״̬�ǹر�
        Winsock1.Close                    '�ر�����
    End If
    Winsock1.RemoteHost = Text1.Text      '��������ַ
    Winsock1.RemotePort = 4567            '�������˿�
    Winsock1.Connect                      '����
End Sub
 Private Sub Form_Load()
    FName = App.Path & "\" & Text2.Text      'ָ�������ļ�����·��
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    Winsock1.Close   '�ر�����
End Sub
 Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then     '�����ǰ������ǻس���
        KeyAscii = 0          'ȡ����ǰ����
        Command1_Click        '����Command1_Click����
    End If
End Sub
 Private Sub Winsock1_Close()
    Winsock1.Close   '�ر�����
End Sub
 
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim TheFile() As Byte                    '�������ݵ�����
    ReDim TheFile(bytesTotal)                '�ض��������½�
    Static YNLen As Boolean                  '�Ƿ�������ļ�����
    Winsock1.GetData TheFile                 '�����յ����ݱ��浽����
    If bytesTotal = 2 And Chr(TheFile(0)) = "C" And Chr(TheFile(1)) = "S" Then    '����յ����ǳɹ�������Ϣ
        Me.Caption = "�ͻ���-----�ɹ�����"    '��ʾ��Ϣ
        Winsock1.SendData "GetFileLen"        '����Ҫ���ļ�������Ϣ
        Text1.Enabled = False                 'Text1��������
        Command1.Enabled = False              '��ť������
        ProgressBar1.Value = 0                '��������ǰֵΪ0
        Exit Sub          '��������
    End If
    If bytesTotal = 3 And Chr(TheFile(0)) = "N" And Chr(TheFile(1)) = "o" And Chr(TheFile(2)) = "F" Then '����յ������޴��ļ�����Ϣ
        MsgBox "���������޴��ļ���", vbInformation, "��ʾ"        '������ʾ
        Exit Sub     '��������
    End If
    If bytesTotal = 3 And Chr(TheFile(0)) = "T" And Chr(TheFile(1)) = "h" And Chr(TheFile(2)) = "E" Then  '����յ��ļ����ͽ�����Ϣ
        Close #1      '�ر��ļ�
        YNLen = False 'δ�����ļ�����������Ϣ
        Me.Caption = "�ͻ���-----�ļ��ѳɹ�����"   '��ʾ��Ϣ
        Winsock1.SendData "ConClose"    '�ر�����
        Text1.Enabled = True            'Text1������
        Command1.Enabled = True         '��ť����
        Exit Sub
    End If
    If YNLen = True Then   '����Ѿ��������ļ�������Ϣ
        Put #1, , TheFile                '�����յ����ݰ�д����ļ�
        Winsock1.SendData "NextB"        '����Ҫ����һ���ݰ�����Ϣ
        ProgressBar1.Value = ProgressBar1.Value + bytesTotal '�����ļ�����
    Else
        Me.Caption = "�ͻ���-----���ڽ�������"     '��ʾ��Ϣ
        Dim i As Integer
        Dim Strs As String   '�����ļ������ַ���
        For i = 0 To bytesTotal - 1
            Strs = Strs & Chr(TheFile(i))   '����ļ����������ַ���
        Next i
        ProgressBar1.Max = Val(Strs)   '���ý��������ֵ
        ProgressBar1.Min = 0           '���ý�������Сֵ
        YNLen = True                   '�Ѿ��������ļ�����������Ϣ
        Winsock1.SendData "FLA"        '�����Ѿ��յ��ļ�����������Ϣ����Ϣ"FLA"

    End If
End Sub
 
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbInformation, "�����:" & Number   '������ʾ
    Winsock1.Close   '�ر�����
End Sub
