<<<<<<< HEAD
Attribute VB_Name = "ModKillProcess"
'VB关闭进程

Option Explicit
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long

Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * 1024
End Type

Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Dim pid As Long
Dim pname As String
'-------------结束进程通用函数 注意进程名要区分大小写
Public Sub taskkill(ByVal taskname As String)
Dim my As PROCESSENTRY32
Dim l As Long
Dim l1 As Long
Dim flag As Boolean
Dim mName As String
Dim I As Integer

l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If l Then
my.dwSize = 1060
If (Process32First(l, my)) Then '遍历第一个进程
Do
I = InStr(1, my.szExeFile, Chr$(0))
mName = LCase$(Left$(my.szExeFile, I - 1))
If mName = LCase$(taskname) Then
pid = my.th32ProcessID
pname = mName
Dim mProcID As Long
mProcID = OpenProcess(1&, -1&, pid)
TerminateProcess mProcID, 0&
flag = True
Exit Sub
Else
flag = False
End If
Loop Until (Process32Next(l, my) < 1) '遍历所有进程知道返回值为False
End If
l1 = CloseHandle(l)
End If
End Sub
=======
Attribute VB_Name = "ModKillProcess"
'VB关闭进程

Option Explicit
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long

Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * 1024
End Type

Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Dim pid As Long
Dim pname As String
'-------------结束进程通用函数 注意进程名要区分大小写
Public Sub taskkill(ByVal taskname As String)
Dim my As PROCESSENTRY32
Dim l As Long
Dim l1 As Long
Dim flag As Boolean
Dim mName As String
Dim I As Integer

l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If l Then
my.dwSize = 1060
If (Process32First(l, my)) Then '遍历第一个进程
Do
I = InStr(1, my.szExeFile, Chr$(0))
mName = LCase$(Left$(my.szExeFile, I - 1))
If mName = LCase$(taskname) Then
pid = my.th32ProcessID
pname = mName
Dim mProcID As Long
mProcID = OpenProcess(1&, -1&, pid)
TerminateProcess mProcID, 0&
flag = True
Exit Sub
Else
flag = False
End If
Loop Until (Process32Next(l, my) < 1) '遍历所有进程知道返回值为False
End If
l1 = CloseHandle(l)
End If
End Sub
>>>>>>> 2abdb2e3a24a7144f6e9a89c8664e5d0469ec446
