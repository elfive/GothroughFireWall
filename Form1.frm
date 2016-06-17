<<<<<<< HEAD
VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Update Password"
   ClientHeight    =   3045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   5520
      ScaleHeight     =   3015
      ScaleWidth      =   5775
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "&OK"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "科学上网，要的就是自由！"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "By：ELFive   2016/04/25"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "如果喜欢本工具，请支持作者，谢谢"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2895
      End
      Begin VB.Image Image1 
         Height          =   2250
         Left            =   3360
         Picture         =   "Form1.frx":0000
         Top             =   360
         Width           =   2250
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   375
      ItemData        =   "Form1.frx":49D9
      Left            =   1680
      List            =   "Form1.frx":49DB
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox List1 
      Height          =   1590
      Left            =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&About"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   525
   End
   Begin VB.Label Label1 
      Caption         =   "请选择接入点："
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   300
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2



Dim CountDown As Integer


Function ShowMessage(ByVal Message As String)
List1.AddItem Message
List1.ListIndex = List1.ListCount - 1
End Function


Public Function CheckConnection() As Boolean
Dim isConnected As Boolean

VBThreadHandle = CreateThread(ByVal 0&, ByVal 0&, AddressOf VBMultiThreadRun.SubIsConnectedState, isConnected, ByVal CREATE_DEFAULT, VBThreadID)
DoEvents

Do Until ThreadMonitorEnd = True
    DoEvents
    Wait 30
Loop

CheckConnection = isConnected

End Function


Private Sub Command1_Click()
List1.Clear
TestHost = "https://www.baidu.com/"
If CheckConnection = False Then
    ShowMessage " 无法连接至服务器，请稍检查您的网络连接"
    Exit Sub
End If
ShowMessage " 正在联系服务器获取密码..."

tmp = Replace(Replace(GetUrlResource(ConnectionPoint(Combo1.ListIndex).URL), Chr(13), ""), Chr(10), "")

If GetConnectionErr Then
    If CheckConnection = False Then
        ShowMessage " 密码获取失败，请稍检查您的网络连接"
    Else
        ShowMessage " 密码获取失败，服务器暂时不可访问，请稍后重试"
    End If
    'Call taskkill("Shadowsocks.exe")
    TestHost = "https://www.youtube.com/"
    Exit Sub
Else
    ShowMessage " 密码获取成功，正在准备写入配置..."
End If
'MsgBox Asc(Left(tmp, 1))
Text = "{" & vbCrLf & _
        Space(2) & Chr(34) & "configs" & Chr(34) & ": [" & vbCrLf & "  {" & vbCrLf & _
        Space(6) & Chr(34) & "server" & Chr(34) & ": " & Chr(34) & ConnectionPoint(Combo1.ListIndex).Account & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "server_port" & Chr(34) & ": 443," & vbCrLf & _
        Space(6) & Chr(34) & "password" & Chr(34) & ": " & Chr(34) & tmp & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "method" & Chr(34) & ": " & Chr(34) & "aes-256-cfb" & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "remarks" & Chr(34) & ": " & Chr(34) & ConnectionPoint(Combo1.ListIndex).Account & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "auth" & Chr(34) & ": false" & vbCrLf & "    }," & vbCrLf & "    {" & vbCrLf & _
        Space(6) & Chr(34) & "server" & Chr(34) & ": " & Chr(34) & "107.181.154.72" & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "server_port" & Chr(34) & ": 453," & vbCrLf & _
        Space(6) & Chr(34) & "password" & Chr(34) & ": " & Chr(34) & "ChEnBo" & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "method" & Chr(34) & ": " & Chr(34) & "aes-256-cfb" & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "remarks" & Chr(34) & ": " & Chr(34) & "ChEnBo" & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "auth" & Chr(34) & ": false" & vbCrLf & "    }" & vbCrLf & "  ]," & vbCrLf & _
        Space(2) & Chr(34) & "strategy" & Chr(34) & ": null," & vbCrLf & _
        Space(2) & Chr(34) & "index" & Chr(34) & ": 0," & vbCrLf & _
        Space(2) & Chr(34) & "global" & Chr(34) & ": false," & vbCrLf & _
        Space(2) & Chr(34) & "enabled" & Chr(34) & ": true," & vbCrLf & _
        Space(2) & Chr(34) & "shareOverLan" & Chr(34) & ": true," & vbCrLf & _
        Space(2) & Chr(34) & "isDefault" & Chr(34) & ": false," & vbCrLf & _
        Space(2) & Chr(34) & "localPort" & Chr(34) & ": 1081," & vbCrLf & _
        Space(2) & Chr(34) & "pacUrl" & Chr(34) & ": null," & vbCrLf & _
        Space(2) & Chr(34) & "useOnlinePac" & Chr(34) & ": false," & vbCrLf & _
        Space(2) & Chr(34) & "availabilityStatistics" & Chr(34) & ": false," & vbCrLf & _
        Space(2) & Chr(34) & "autoCheckUpdate" & Chr(34) & ": true," & vbCrLf & Space(2) & Chr(34) & "logViewer" & Chr(34) & ": null" & vbCrLf & "}"

'MsgBox Text
TmpPath = App.Path
Path = IIf(Right(TmpPath, 1) = "\", TmpPath, TmpPath & "\")
ConfigPath = Path & "gui-config.json"

ShowMessage " 正在关闭SS进程..."
Call taskkill("Shadowsocks.exe")

Open ConfigPath For Output As #1
    'Print #1, StrConv(Text, vbUnicode)
    Print #1, Text
Close #1

ShowMessage " 配置写入成功！正在启动SS..."

'Clipboard.Clear                                     '清空剪贴板
'Clipboard.SetText Text

ExePath = Path & "Shadowsocks.exe"
If Dir(ExePath) = "" Then ShowMessage " SS程序未找到，请将其复制到本程序目录下再运行": GoTo SubExit

a = Shell(ExePath, vbNormalNoFocus)

If a = 0 Then
    ShowMessage " 程序启动失败！"
    MsgBox "SS启动失败，请从本程序目录下运行Shadowsocks.exe程序"
Else
    ShowMessage " 程序启动成功！正在检测是否连接正常..."
    Wait 10000
    DoEvents
    If CheckConnection = True Then
        ShowMessage " 连接正常，可以开始科学上网！"
        Call ShellExecute(Form1.hwnd, "open", "https://www.youtube.com/", vbNullString, vbNullString, &H0)
SubExit:
        Timer1.Interval = 1000
    Else
        ShowMessage " 连接失败，请切换接入点后重试，如还不能上网，则有可能是流量用完，待下月初即可恢复"
        'Call taskkill("Shadowsocks.exe")
    End If

End If


End Sub

Private Sub Command2_Click()
End: Set Form1 = Nothing
End Sub

Private Sub Command3_Click()
Picture1.Visible = False
End Sub

Private Sub Form_Load()
Form1.Show
SetWindowPos Form1.hwnd, -1, 0, 0, 0, 0, 3 '窗体置顶
Timer1.Interval = 0
CountDown = 3

List1.AddItem " 本工具适用于2.3-3.0版本的SS"
List1.AddItem " 请在IPv4协议下使用本工具"

ConnectionPoint(0).Account = "mgm1.sspw.pw"
ConnectionPoint(0).URL = "http://www.shadowsocks.asia/ssmm1.php"
ConnectionPoint(1).Account = "mgm2.sspw.pw"
ConnectionPoint(1).URL = "http://www.shadowsocks.asia/ssmm2.php"
ConnectionPoint(2).Account = "mgm5.sspw.pw"
ConnectionPoint(2).URL = "http://www.shadowsocks.asia/ssmm5.php"

For I = LBound(ConnectionPoint) To UBound(ConnectionPoint)
    Combo1.AddItem ConnectionPoint(I).Account
Next
Combo1.ListIndex = 0


Picture1.Top = 0
Picture1.Left = 0
Picture1.Height = Form1.Height
Picture1.Width = Form1.Width


TestHost = "https://www.nasa.gov"


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = RGB(0, 0, 0)
End Sub

Private Sub Label2_Click()
Picture1.Visible = True
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = RGB(0, 0, 255)
End Sub

Private Sub Picture1_LostFocus()
Picture1.Visible = False
End Sub

Private Sub Timer1_Timer()
ShowMessage " 程序将在" & CountDown & "秒后自动退出..."
CountDown = CountDown - 1
If CountDown = -1 Then End: Set Form1 = Nothing
End Sub
=======
VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Update Password"
   ClientHeight    =   3045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   5520
      ScaleHeight     =   3015
      ScaleWidth      =   5775
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "&OK"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "科学上网，要的就是自由！"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "By：ELFive   2016/04/25"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "如果喜欢本工具，请支持作者，谢谢"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2895
      End
      Begin VB.Image Image1 
         Height          =   2250
         Left            =   3360
         Picture         =   "Form1.frx":0000
         Top             =   360
         Width           =   2250
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   375
      ItemData        =   "Form1.frx":49D9
      Left            =   1680
      List            =   "Form1.frx":49DB
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox List1 
      Height          =   1590
      Left            =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&About"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   525
   End
   Begin VB.Label Label1 
      Caption         =   "请选择接入点："
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   300
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2



Dim CountDown As Integer


Function ShowMessage(ByVal Message As String)
List1.AddItem Message
List1.ListIndex = List1.ListCount - 1
End Function


Public Function CheckConnection() As Boolean
Dim isConnected As Boolean

VBThreadHandle = CreateThread(ByVal 0&, ByVal 0&, AddressOf VBMultiThreadRun.SubIsConnectedState, isConnected, ByVal CREATE_DEFAULT, VBThreadID)
DoEvents

Do Until ThreadMonitorEnd = True
    DoEvents
    Wait 30
Loop

CheckConnection = isConnected

End Function


Private Sub Command1_Click()
List1.Clear
TestHost = "https://www.baidu.com/"
If CheckConnection = False Then
    ShowMessage " 无法连接至服务器，请稍检查您的网络连接"
    Exit Sub
End If
ShowMessage " 正在联系服务器获取密码..."

tmp = Replace(Replace(GetUrlResource(ConnectionPoint(Combo1.ListIndex).URL), Chr(13), ""), Chr(10), "")

If GetConnectionErr Then
    If CheckConnection = False Then
        ShowMessage " 密码获取失败，请稍检查您的网络连接"
    Else
        ShowMessage " 密码获取失败，服务器暂时不可访问，请稍后重试"
    End If
    'Call taskkill("Shadowsocks.exe")
    TestHost = "https://www.youtube.com/"
    Exit Sub
Else
    ShowMessage " 密码获取成功，正在准备写入配置..."
End If
'MsgBox Asc(Left(tmp, 1))
Text = "{" & vbCrLf & _
        Space(2) & Chr(34) & "configs" & Chr(34) & ": [" & vbCrLf & "  {" & vbCrLf & _
        Space(6) & Chr(34) & "server" & Chr(34) & ": " & Chr(34) & ConnectionPoint(Combo1.ListIndex).Account & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "server_port" & Chr(34) & ": 443," & vbCrLf & _
        Space(6) & Chr(34) & "password" & Chr(34) & ": " & Chr(34) & tmp & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "method" & Chr(34) & ": " & Chr(34) & "aes-256-cfb" & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "remarks" & Chr(34) & ": " & Chr(34) & ConnectionPoint(Combo1.ListIndex).Account & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "auth" & Chr(34) & ": false" & vbCrLf & "    }," & vbCrLf & "    {" & vbCrLf & _
        Space(6) & Chr(34) & "server" & Chr(34) & ": " & Chr(34) & "107.181.154.72" & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "server_port" & Chr(34) & ": 453," & vbCrLf & _
        Space(6) & Chr(34) & "password" & Chr(34) & ": " & Chr(34) & "ChEnBo" & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "method" & Chr(34) & ": " & Chr(34) & "aes-256-cfb" & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "remarks" & Chr(34) & ": " & Chr(34) & "ChEnBo" & Chr(34) & "," & vbCrLf & _
        Space(6) & Chr(34) & "auth" & Chr(34) & ": false" & vbCrLf & "    }" & vbCrLf & "  ]," & vbCrLf & _
        Space(2) & Chr(34) & "strategy" & Chr(34) & ": null," & vbCrLf & _
        Space(2) & Chr(34) & "index" & Chr(34) & ": 0," & vbCrLf & _
        Space(2) & Chr(34) & "global" & Chr(34) & ": false," & vbCrLf & _
        Space(2) & Chr(34) & "enabled" & Chr(34) & ": true," & vbCrLf & _
        Space(2) & Chr(34) & "shareOverLan" & Chr(34) & ": true," & vbCrLf & _
        Space(2) & Chr(34) & "isDefault" & Chr(34) & ": false," & vbCrLf & _
        Space(2) & Chr(34) & "localPort" & Chr(34) & ": 1081," & vbCrLf & _
        Space(2) & Chr(34) & "pacUrl" & Chr(34) & ": null," & vbCrLf & _
        Space(2) & Chr(34) & "useOnlinePac" & Chr(34) & ": false," & vbCrLf & _
        Space(2) & Chr(34) & "availabilityStatistics" & Chr(34) & ": false," & vbCrLf & _
        Space(2) & Chr(34) & "autoCheckUpdate" & Chr(34) & ": true," & vbCrLf & Space(2) & Chr(34) & "logViewer" & Chr(34) & ": null" & vbCrLf & "}"

'MsgBox Text
TmpPath = App.Path
Path = IIf(Right(TmpPath, 1) = "\", TmpPath, TmpPath & "\")
ConfigPath = Path & "gui-config.json"

ShowMessage " 正在关闭SS进程..."
Call taskkill("Shadowsocks.exe")

Open ConfigPath For Output As #1
    'Print #1, StrConv(Text, vbUnicode)
    Print #1, Text
Close #1

ShowMessage " 配置写入成功！正在启动SS..."

'Clipboard.Clear                                     '清空剪贴板
'Clipboard.SetText Text

ExePath = Path & "Shadowsocks.exe"
If Dir(ExePath) = "" Then ShowMessage " SS程序未找到，请将其复制到本程序目录下再运行": GoTo SubExit

a = Shell(ExePath, vbNormalNoFocus)

If a = 0 Then
    ShowMessage " 程序启动失败！"
    MsgBox "SS启动失败，请从本程序目录下运行Shadowsocks.exe程序"
Else
    ShowMessage " 程序启动成功！正在检测是否连接正常..."
    Wait 10000
    DoEvents
    If CheckConnection = True Then
        ShowMessage " 连接正常，可以开始科学上网！"
        Call ShellExecute(Form1.hwnd, "open", "https://www.youtube.com/", vbNullString, vbNullString, &H0)
SubExit:
        Timer1.Interval = 1000
    Else
        ShowMessage " 连接失败，请切换接入点后重试，如还不能上网，则有可能是流量用完，待下月初即可恢复"
        'Call taskkill("Shadowsocks.exe")
    End If

End If


End Sub

Private Sub Command2_Click()
End: Set Form1 = Nothing
End Sub

Private Sub Command3_Click()
Picture1.Visible = False
End Sub

Private Sub Form_Load()
Form1.Show
SetWindowPos Form1.hwnd, -1, 0, 0, 0, 0, 3 '窗体置顶
Timer1.Interval = 0
CountDown = 3

List1.AddItem " 本工具适用于2.3-3.0版本的SS"
List1.AddItem " 请在IPv4协议下使用本工具"

ConnectionPoint(0).Account = "mgm1.sspw.pw"
ConnectionPoint(0).URL = "http://www.shadowsocks.asia/ssmm1.php"
ConnectionPoint(1).Account = "mgm2.sspw.pw"
ConnectionPoint(1).URL = "http://www.shadowsocks.asia/ssmm2.php"
ConnectionPoint(2).Account = "mgm5.sspw.pw"
ConnectionPoint(2).URL = "http://www.shadowsocks.asia/ssmm5.php"

For I = LBound(ConnectionPoint) To UBound(ConnectionPoint)
    Combo1.AddItem ConnectionPoint(I).Account
Next
Combo1.ListIndex = 0


Picture1.Top = 0
Picture1.Left = 0
Picture1.Height = Form1.Height
Picture1.Width = Form1.Width


TestHost = "https://www.nasa.gov"


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = RGB(0, 0, 0)
End Sub

Private Sub Label2_Click()
Picture1.Visible = True
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = RGB(0, 0, 255)
End Sub

Private Sub Picture1_LostFocus()
Picture1.Visible = False
End Sub

Private Sub Timer1_Timer()
ShowMessage " 程序将在" & CountDown & "秒后自动退出..."
CountDown = CountDown - 1
If CountDown = -1 Then End: Set Form1 = Nothing
End Sub
>>>>>>> 2abdb2e3a24a7144f6e9a89c8664e5d0469ec446
