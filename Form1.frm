VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "图种制作工具"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5040
   StartUpPosition =   3  '窗口缺省
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
      Caption         =   " 合并文件(&G)"
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
      Caption         =   " 打开文件(&O)"
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
      Caption         =   " 打开图片(&P)"
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
Private Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) '窗口置顶

Private Sub Form_Load()
Dim rtn
rtn = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3) '窗口置顶
End Sub

Private Sub Label1_Click()
    CommonDialog1.CancelError = True ' 设置“CancelError”为 True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly ' 设置标志
    CommonDialog1.InitDir = App.Path '设置初始路径
    CommonDialog1.Filter = "图像文件(*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif" ' 设置过滤器
    CommonDialog1.FilterIndex = 2    ' 指定缺省的过滤器
    CommonDialog1.DialogTitle = "请选择一个图片以隐藏文件" '设置标题
    CommonDialog1.ShowOpen    ' 显示“打开”对话框
    Text1.Text = CommonDialog1.FileName  ' 显示选定文件的名字
    Exit Sub
ErrHandler:
    ' 用户按了“取消”按钮
    Exit Sub
End Sub

Private Sub Label2_Click()

    CommonDialog1.CancelError = True ' 设置“CancelError”为 True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly ' 设置标志
    CommonDialog1.InitDir = App.Path '设置初始路径
    CommonDialog1.Filter = "所有文件(*.*)|*.*" ' 设置过滤器
    CommonDialog1.FilterIndex = 2    ' 指定缺省的过滤器
    CommonDialog1.DialogTitle = "请选择一个要被隐藏的文件" '设置标题
    CommonDialog1.ShowOpen    ' 显示“打开”对话框
    Text2.Text = CommonDialog1.FileName  ' 显示选定文件的名字
    Exit Sub
ErrHandler:
    ' 用户按了“取消”按钮
    Exit Sub
End Sub

Private Sub Label3_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "你还没有选择文件！", vbCritical, "图种制作工具"
Exit Sub
End If

    CommonDialog1.CancelError = True ' 设置“CancelError”为 True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly ' 设置标志
    CommonDialog1.Filter = "图像文件(*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif" ' 设置过滤器
    CommonDialog1.FilterIndex = 2    ' 指定缺省的过滤器
    CommonDialog1.DialogTitle = "请选择图种保存的位置" '设置标题
    CommonDialog1.ShowSave    ' 显示“另存为”对话框
    'MsgBox CommonDialog1.FileName
    Shell "cmd.exe /c copy /b " & """" & Text1.Text & """" & " + " & """" & Text2.Text & """" & " " & """" & CommonDialog1.FileName & """", vbNormalNoFocus
    MsgBox "已成功创建图种：" & CommonDialog1.FileName & "快分享给你的朋友吧！", 64
    Exit Sub
ErrHandler:
    ' 用户按了“取消”按钮
    Exit Sub

End Sub
