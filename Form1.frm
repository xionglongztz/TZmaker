VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ͼ����������"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5040
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " �ϲ��ļ�(&G)"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ���ļ�(&O)"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ��ͼƬ(&P)"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) '�����ö�

Private Sub Form_Load()
Dim rtn
rtn = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3) '�����ö�
End Sub

Private Sub Label1_Click()
    CommonDialog1.CancelError = True ' ���á�CancelError��Ϊ True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly ' ���ñ�־
    CommonDialog1.InitDir = App.Path '���ó�ʼ·��
    CommonDialog1.Filter = "ͼ���ļ�(*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif" ' ���ù�����
    CommonDialog1.FilterIndex = 2    ' ָ��ȱʡ�Ĺ�����
    CommonDialog1.DialogTitle = "��ѡ��һ��ͼƬ�������ļ�" '���ñ���
    CommonDialog1.ShowOpen    ' ��ʾ���򿪡��Ի���
    Text1.Text = CommonDialog1.FileName  ' ��ʾѡ���ļ�������
    Exit Sub
ErrHandler:
    ' �û����ˡ�ȡ������ť
    Exit Sub
End Sub

Private Sub Label2_Click()

    CommonDialog1.CancelError = True ' ���á�CancelError��Ϊ True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly ' ���ñ�־
    CommonDialog1.InitDir = App.Path '���ó�ʼ·��
    CommonDialog1.Filter = "�����ļ�(*.*)|*.*" ' ���ù�����
    CommonDialog1.FilterIndex = 2    ' ָ��ȱʡ�Ĺ�����
    CommonDialog1.DialogTitle = "��ѡ��һ��Ҫ�����ص��ļ�" '���ñ���
    CommonDialog1.ShowOpen    ' ��ʾ���򿪡��Ի���
    Text2.Text = CommonDialog1.FileName  ' ��ʾѡ���ļ�������
    Exit Sub
ErrHandler:
    ' �û����ˡ�ȡ������ť
    Exit Sub
End Sub

Private Sub Label3_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "�㻹û��ѡ���ļ���", vbCritical, "ͼ����������"
Exit Sub
End If

    CommonDialog1.CancelError = True ' ���á�CancelError��Ϊ True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly ' ���ñ�־
    CommonDialog1.Filter = "ͼ���ļ�(*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif" ' ���ù�����
    CommonDialog1.FilterIndex = 2    ' ָ��ȱʡ�Ĺ�����
    CommonDialog1.DialogTitle = "��ѡ��ͼ�ֱ����λ��" '���ñ���
    CommonDialog1.ShowSave    ' ��ʾ�����Ϊ���Ի���
    'MsgBox CommonDialog1.FileName
    Shell "cmd.exe /c copy /b " & """" & Text1.Text & """" & " + " & """" & Text2.Text & """" & " " & """" & CommonDialog1.FileName & """", vbNormalNoFocus
    MsgBox "�ѳɹ�����ͼ�֣�" & CommonDialog1.FileName & "������������Ѱɣ�", 64
    Exit Sub
ErrHandler:
    ' �û����ˡ�ȡ������ť
    Exit Sub

End Sub
