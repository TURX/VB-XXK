VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ServerFrm 
   Caption         =   "��������-------������"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   4995
   StartUpPosition =   3  '����ȱʡ
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
Private FName As String      'Ҫ���͵��ļ�����·���������ļ�����

Private Sub Command1_Click()
ClientFrm.Show
End Sub
Private Sub Form_Load()
    'Download by http://www.newxing.com
    CClose
    FName = App.Path & "\319.jpg"   'ָ�������ļ�·��
End Sub
 Private Sub Form_Unload(Cancel As Integer)
    Winsock1.Close     '�ر�����
End Sub
 Private Sub Winsock1_Close()
    CClose
    Close #1    'ΪԤ�����ͻ��˷Ƿ��ر�ʱ��ɲ��ܼ�ʱ�ر��Ѵ򿪵��ļ������
End Sub
 
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then   '���Winsock1��ǰ״̬�ǹر�
        Winsock1.Close                    '�ر�����
    End If
    Winsock1.Accept requestID             '�ر��˵�Winsock1��������
    Winsock1.SendData "CS"           '����վ����������Ϣ�������ѳɹ����ӣ�
    Me.Caption = "��������-------������"   '��ʾ��Ϣ
End Sub
 
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim TStr As String         '�ѵ�����ַ�����Ϣ����
    Static FileLen As Long     '�ļ����ȣ����ֽ�Ϊ��λ��
    Static CurPos As Long      '�ļ����ͽ���
    Dim S As Long              '��һ�����ݰ��Ĵ�С������ȷ��������½磩
    Dim TheFile() As Byte      '����һ���ɱ��С�����飬���Դ��Ͷ��������ݰ�
    Winsock1.GetData TStr      '�����Ե�����ַ�������
    Select Case TStr
    Case "ConClose"            '����յ����ǹر�������Ϣ
        CClose
    Case "GetFileLen"          '����յ�����Ҫ�����ļ�������Ϣ
        If Dir(FName) <> "" Then          '����ļ�����
            Open FName For Binary As #1   '���ļ��Ի�ø��ļ��ĳ���
            FileLen = LOF(1)              '����Ѵ��ļ��ĳ��ȣ����ֽ�Ϊ��λ��
            Winsock1.SendData Trim(Str(FileLen)) '�����ļ�����
        Else
            Winsock1.SendData "NoF"     '����δ�ҵ��ļ���Ϣ
        End If
    Case "FLA"                            '����յ������ļ������ѵ�����Ϣ
        ReDim TheFile(1 To 3072)          '�ض��������½�Ϊ3072
        Get #1, , TheFile                 '����Ѵ��ļ���һ�����ֽڣ�3072�����������䱣����TheFile������
        Winsock1.SendData TheFile         '���͵�һ�����ݰ�
        CurPos = 3072                     '�ļ���ǰ�Ѷ�ȡ��3072���ֽ�
    Case "NextB"                          '����յ�����Ҫ����һ�����ݰ�����Ϣ
        If CurPos = FileLen Then          '����Ѿ����������е��ֽ�
            Winsock1.SendData "ThE"       '�����ļ����ͽ�����Ϣ
            Close #1                      '�ر��ļ�
            Exit Sub                      '��������
        End If
        S = CurPos + 3072                 '��ǰ�ļ����ͽ��ȼ�3072���ֽ�
        If S > FileLen Then               '����������ļ��ܳ���
            S = FileLen - CurPos          '��һ�������½�ӦΪ���µ��ֽ���
        Else
            S = 3072                      '��һ�������½���Ϊ3072
        End If
        ReDim TheFile(1 To S)             '�ض��������½�ΪS
        Get #1, , TheFile                 '��ȡ���ļ�����һ�������ݣ���������TheFile������
        Winsock1.SendData TheFile         '���͸ð�
        CurPos = CurPos + S               'ȷ�����ͽ���
    End Select
End Sub
 
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbInformation, "�����:" & Number   '������ʾ
    CClose
End Sub
Private Sub CClose()   '�رյ�ǰ���ӣ�����������
    Winsock1.Close     '�رյ�ǰ����
    Winsock1.LocalPort = 4567      '���ü����˿ں�Ϊ4567
    Winsock1.Listen                '��ʼ����
    Me.Caption = "��������-------�����ѶϿ�,��������" '��ʾ��Ϣ
End Sub
