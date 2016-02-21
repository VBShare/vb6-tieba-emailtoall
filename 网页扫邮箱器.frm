VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "发邮箱神器（By：吸金大法之001 QQ：2523198627）"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   12855
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "邮箱配置"
      Height          =   2055
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   2895
      Begin VB.CheckBox chkLocked 
         Caption         =   "锁定设置"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1600
         Width           =   2655
      End
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtEmail 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "密码"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "邮箱"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox txtAttach 
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   5760
      Width           =   5415
   End
   Begin VB.TextBox txtTitle 
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   600
      Width           =   5415
   End
   Begin VB.TextBox txtContent 
      Height          =   4455
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   5415
   End
   Begin VB.TextBox txtLog 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2760
      Width           =   6495
   End
   Begin VB.CommandButton btnGetAndSend 
      Caption         =   "一键获取邮箱并发送"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   1
      Top             =   6240
      Width           =   5415
   End
   Begin VB.TextBox txtUrl 
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label5 
      Caption         =   "附件"
      Height          =   255
      Left            =   6840
      TabIndex        =   10
      Top             =   5835
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "邮件内容"
      Height          =   255
      Left            =   6480
      TabIndex        =   8
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "邮件标题"
      Height          =   255
      Left            =   6480
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "日志"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "有邮箱的网址"
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private web As New WebCode
Private mail As New CEmail
Private Sub chkLocked_Click()
  If chkLocked.Value = 1 Then
    '本次勾选了
  Else
    '本次取消勾选了
  End If
End Sub

Private Sub btnGetAndSend_Click()
  On Error Resume Next
  Dim pageHtml As String
  Dim re As RegExp
  Dim mh As Match
  Dim mhs As MatchCollection
  Dim retstr As String
  Dim r() As String
  Dim i As Integer, OkCount As Integer
  
  '检测地址是否为空
  If txtUrl.Text = "" Then
    MsgBox "网址为空", vbCritical, "发送大卫提示"
    Exit Sub
  End If

  If txtEmail.Text = "" Then
    MsgBox "发件邮箱为空", vbCritical, "发送大卫提示"
    Exit Sub
  End If
  
  If txtPass.Text = "" Then
    MsgBox "发件密码为空", vbCritical, "发送大卫提示"
    Exit Sub
  End If
  
  pageHtml = web.GetHTMLCode(txtUrl.Text)
  '用正则函数检测pageHtml中所有符合邮箱格式的字符串

  retstr = ""
  

  Set re = New RegExp
  re.IgnoreCase = False
  re.Global = True
  re.Pattern = "\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"
  Set mhs = re.Execute(pageHtml)
   For Each mh In mhs   '' Iterate Matches collection.
    retstr = retstr & mh.Value & ";" & vbCrLf
  Next

  '支持发送vip.qq.com的格式
  mail.SetUser txtEmail.Text, txtPass.Text
  mail.SetSMTP "smtp.163.com"
  If txtAttach.Text <> "" Then
    If Dir(txtAttach.Text) <> "" Then
      mail.AddFile txtAttach.Text
    End If
  End If
  r = Split(retstr, vbCrLf)
  txtLog.Text = Time & " - 启动任务"
  For i = 0 To UBound(r)
    DoEvents
    If Len(r(i)) <= 1 Then
      DoEvents
      txtLog.Text = txtLog.Text & vbCrLf & Time & " - [" & i + 1 & "/" & UBound(r) + 1 & "]空数据"
      DoEvents
    Else
      mail.SendMail txtEmail.Text, r(i), txtTitle.Text, txtContent.Text
      If Err.Number <> 0 Then
        DoEvents
        txtLog.Text = txtLog.Text & vbCrLf & Time & " - [" & i + 1 & "/" & UBound(r) + 1 & "]" & Err.Description
        DoEvents
        Err.Clear
      Else
        DoEvents
        txtLog.Text = txtLog.Text & vbCrLf & Time & " - [" & i + 1 & "/" & UBound(r) + 1 & "]" & "[" & r(i) & "]发送成功"
        DoEvents
        OkCount = OkCount + 1
        Text1.Text = "成功 " & OkCount & " 个"
        DoEvents
      End If
    End If
    txtLog.SelStart = Len(txtLog.Text)
  Next i
  txtLog.Text = txtLog.Text & vbCrLf & Time & " - 任务完成"
  txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub SaveSet(ByVal Key As String, ByVal Value As String)
  SaveSetting "vb6-tieba-emailtoall", "mailconfig", Key, Value
End Sub

Private Function GetSet(ByVal Key As String) As String
  GetSet = GetSetting("vb6-tieba-emailtoall", "mailconfig", Key)
End Function


Private Sub Form_Load()
  txtEmail.Text = GetSet("mail")
  txtPass.Text = GetSet("pass")
  txtUrl.Text = GetSet("url")
  txtTitle.Text = GetSet("title")
  txtContent.Text = GetSet("content")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSet "mail", txtEmail.Text
  SaveSet "pass", txtPass.Text
  SaveSet "url", txtUrl.Text
  SaveSet "title", txtTitle.Text
  SaveSet "content", txtContent.Text
  End
End Sub
