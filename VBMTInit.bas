<<<<<<< HEAD
Attribute VB_Name = "VBMTInit"
'*******************************************************************************************************************************
'非原创的代码部分已标明原作者及出处。其余部分皆为本人【秋枫萧萧】（百度ID:hhyjq007）自己研究而成
'感谢老汉、TGY、IZERO等等为VB6完美、稳定多线程铺平道路的前辈们！
'若有任何疑问可到【VBGOOD论坛】或百度贴吧【VB吧】发帖交流！
'*******************************************************************************************************************************
Option Explicit

'ClassID
Public Type mIID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(7) As Byte
End Type

'线程安全属性数据结构（已修正，但通常用不到）
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Byte
End Type

'临界区数据结构（已修正）
Public Type CRITICAL_SECTION
    DebugInfo As Long
    LockCount As Long
    RecursionCount As Long
    OwningThread As Long
    LockSemaphore As Long
    SpinCount As Long
End Type

'*关键！VB6内部隐藏函数，由VBGOOD论坛的老汉（DOWNLOAD）首次发现！能初始化VB6运行库中绝大部分的内容
Public Declare Function CreateIExprSrvObj Lib "msvbvm60.dll" (ByVal p1_0 As Long, ByVal p2_4 As Long, ByVal p3_0 As Long) As Long

'创建线程API，已针对VB6修正声明
Public Declare Function CreateThread Lib "kernel32" (ByVal lpSecurityAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Const CREATE_DEFAULT = &H0&          '默认值，创建一个立即运行的线程
Public Const CREATE_SUSPENDED = &H4&        '线程创建后不立即运行，需要ResumeThread激活才开始运行
Public Const STACK_SIZE_CUSTOM = &H10000    '自定义堆栈大小（StackSize），以字节为单位

Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long        '获取当前线程ID
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long       '获取当前进程ID
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long     '继续运行被暂停的线程
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long    '暂停某个线程的运行（使用时注意临界区问题！某线程进入临界区后被暂停会导致临界区被永久占用。直到线程被继续运行到离开临界区）
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long    '强行终止线程（强烈建议只在程序要全部退出时使用！某线程进入临界区后被强制退出会导致临界区被永久占用，几乎没有补救措施。）
Public Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '初始化临界区
Public Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '删除临界区
Public Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '进入临界区
Public Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '离开临界区

Public Declare Function VBGetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModName As Long) As Long     '获取模块基址/句柄
Public Declare Sub UserDllMain Lib "msvbvm60.dll" (u1 As Long, u2 As Long, ByVal u3_h As Long, ByVal u4_1 As Long, ByVal u5_0 As Long)      'VB6内部隐藏函数 用户DLL入口
Public Declare Function VBDllGetClassObject Lib "msvbvm60.dll" (g1 As Long, g2 As Long, ByVal g3_vbHeader As Long, REFCLSID As Long, REFIID As mIID, ppv As Long) As Long   ''VB6内部隐藏函数 获取类对象

Public Declare Sub InitCommonControls Lib "comctl32" ()     '初始化通用控件
Public Declare Function CoInitializeEx Lib "ole32.dll" (ByVal pvReserved As Long, ByVal dwCoInit As Long) As Long       '初始化COM组件
Public Const COINIT_APARTMENTTHREADED = &H2&            '以单线程模式运行COM相关模块
Public Const COINIT_MULTITHREADED = &H0&                    '以多线程模式运行COM相关模块
Public Const COINIT_DISABLE_OLE1DDE = &H4&                  '关闭OLE1DDE
Public Const COINIT_SPEED_OVER_MEMORY = &H8&            '牺牲更多的内存为代价，让COM相关模块运行的更快（强烈推荐！）

Public Declare Sub CoUninitialize Lib "ole32.dll" ()                '卸载COM组件

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long      '关闭句柄
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)        '拷贝内存
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)          '线程休眠一段时间

'Unicode消息框API
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxW" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
'以下都是返回值的常量
Public Const IDOK = 1
Public Const IDCANCEL = 2
Public Const IDABORT = 3
Public Const IDRETRY = 4
Public Const IDIGNORE = 5
Public Const IDYES = 6
Public Const IDNO = 7
Public Const IDPROMPT = &HFFFF&

'           '防止重入标志                    '运行在IDE或EXE标志    '存放VB数据头
Public AvoidReentrant  As Boolean, IsIDEorEXE As Long, VBHeader As Long

'***********************************VBGOOD论坛的TGY首创的读取VB6数据头的方法，这里略作修改和优化*******************
Public Function GETVBHeader() As Boolean            '获取VB6数据头          此函数内容看不懂没关系，直接使用即可
    Dim BaseAds As Long, GetOffSet As Long
    Dim VBHdChar(3) As Byte, MemData(&H1FDA&) As Byte
    IsIDEorEXE = App.LogMode
    If IsIDEorEXE <> 0 Then
        VBHdChar(0) = &H56          'V
        VBHdChar(1) = &H42          'B
        VBHdChar(2) = &H35          '5
        VBHdChar(3) = &H21          '!
        BaseAds = VBGetModuleHandle(ByVal 0&) + &H1024&
        CopyMemory MemData(0), ByVal (BaseAds), &H1FDB&
        GetOffSet = InStrB(1, MemData, VBHdChar, vbBinaryCompare)
            If GetOffSet > 0 Then
                VBHeader = GetOffSet + BaseAds - 1
                GETVBHeader = True
            Else
                MessageBox ByVal 0&, StrPtr("无法获取VB数据头！程序将运行的不稳定。"), StrPtr("VB数据头定位标志被篡改！"), vbExclamation
                VBHeader = 0
                GETVBHeader = False
            End If
    Else
    VBHeader = 0
    GETVBHeader = False
    End If
End Function

'********************VBGOOD论坛的老汉(DOWNLOAD)，基于国际友人izero@Slovakia的多线程DLL首创的方法，这里略作修改*******************
Public Function InitVBdll() As Boolean                          '诱导VB6运行库其它部分完全初始化 此函数内容看不懂没关系，直接使用即可
    Dim pIID As mIID, pDummy As Long
    Dim u1 As Long, u2 As Long, u3 As Long
    If VBHeader > 0 Then
            ' Set pIID = IID of IClassFactory
            ' = {00000001-0000-0000-C000-000000000046}
            pIID.data1 = &H1&
            pIID.data4(0) = &HC0
            pIID.data4(7) = &H46
            u3 = VBGetModuleHandle(ByVal 0&)
            'get u1,u2 for VBDllGetClassObject
            UserDllMain u1, u2, u3, 1&, 0&
            VBDllGetClassObject u1, u2, VBHeader, pDummy, pIID, pDummy
            InitVBdll = True
    Else
            InitVBdll = False
    End If

End Function
=======
Attribute VB_Name = "VBMTInit"
'*******************************************************************************************************************************
'非原创的代码部分已标明原作者及出处。其余部分皆为本人【秋枫萧萧】（百度ID:hhyjq007）自己研究而成
'感谢老汉、TGY、IZERO等等为VB6完美、稳定多线程铺平道路的前辈们！
'若有任何疑问可到【VBGOOD论坛】或百度贴吧【VB吧】发帖交流！
'*******************************************************************************************************************************
Option Explicit

'ClassID
Public Type mIID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(7) As Byte
End Type

'线程安全属性数据结构（已修正，但通常用不到）
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Byte
End Type

'临界区数据结构（已修正）
Public Type CRITICAL_SECTION
    DebugInfo As Long
    LockCount As Long
    RecursionCount As Long
    OwningThread As Long
    LockSemaphore As Long
    SpinCount As Long
End Type

'*关键！VB6内部隐藏函数，由VBGOOD论坛的老汉（DOWNLOAD）首次发现！能初始化VB6运行库中绝大部分的内容
Public Declare Function CreateIExprSrvObj Lib "msvbvm60.dll" (ByVal p1_0 As Long, ByVal p2_4 As Long, ByVal p3_0 As Long) As Long

'创建线程API，已针对VB6修正声明
Public Declare Function CreateThread Lib "kernel32" (ByVal lpSecurityAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Const CREATE_DEFAULT = &H0&          '默认值，创建一个立即运行的线程
Public Const CREATE_SUSPENDED = &H4&        '线程创建后不立即运行，需要ResumeThread激活才开始运行
Public Const STACK_SIZE_CUSTOM = &H10000    '自定义堆栈大小（StackSize），以字节为单位

Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long        '获取当前线程ID
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long       '获取当前进程ID
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long     '继续运行被暂停的线程
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long    '暂停某个线程的运行（使用时注意临界区问题！某线程进入临界区后被暂停会导致临界区被永久占用。直到线程被继续运行到离开临界区）
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long    '强行终止线程（强烈建议只在程序要全部退出时使用！某线程进入临界区后被强制退出会导致临界区被永久占用，几乎没有补救措施。）
Public Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '初始化临界区
Public Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '删除临界区
Public Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '进入临界区
Public Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '离开临界区

Public Declare Function VBGetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModName As Long) As Long     '获取模块基址/句柄
Public Declare Sub UserDllMain Lib "msvbvm60.dll" (u1 As Long, u2 As Long, ByVal u3_h As Long, ByVal u4_1 As Long, ByVal u5_0 As Long)      'VB6内部隐藏函数 用户DLL入口
Public Declare Function VBDllGetClassObject Lib "msvbvm60.dll" (g1 As Long, g2 As Long, ByVal g3_vbHeader As Long, REFCLSID As Long, REFIID As mIID, ppv As Long) As Long   ''VB6内部隐藏函数 获取类对象

Public Declare Sub InitCommonControls Lib "comctl32" ()     '初始化通用控件
Public Declare Function CoInitializeEx Lib "ole32.dll" (ByVal pvReserved As Long, ByVal dwCoInit As Long) As Long       '初始化COM组件
Public Const COINIT_APARTMENTTHREADED = &H2&            '以单线程模式运行COM相关模块
Public Const COINIT_MULTITHREADED = &H0&                    '以多线程模式运行COM相关模块
Public Const COINIT_DISABLE_OLE1DDE = &H4&                  '关闭OLE1DDE
Public Const COINIT_SPEED_OVER_MEMORY = &H8&            '牺牲更多的内存为代价，让COM相关模块运行的更快（强烈推荐！）

Public Declare Sub CoUninitialize Lib "ole32.dll" ()                '卸载COM组件

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long      '关闭句柄
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)        '拷贝内存
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)          '线程休眠一段时间

'Unicode消息框API
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxW" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
'以下都是返回值的常量
Public Const IDOK = 1
Public Const IDCANCEL = 2
Public Const IDABORT = 3
Public Const IDRETRY = 4
Public Const IDIGNORE = 5
Public Const IDYES = 6
Public Const IDNO = 7
Public Const IDPROMPT = &HFFFF&

'           '防止重入标志                    '运行在IDE或EXE标志    '存放VB数据头
Public AvoidReentrant  As Boolean, IsIDEorEXE As Long, VBHeader As Long

'***********************************VBGOOD论坛的TGY首创的读取VB6数据头的方法，这里略作修改和优化*******************
Public Function GETVBHeader() As Boolean            '获取VB6数据头          此函数内容看不懂没关系，直接使用即可
    Dim BaseAds As Long, GetOffSet As Long
    Dim VBHdChar(3) As Byte, MemData(&H1FDA&) As Byte
    IsIDEorEXE = App.LogMode
    If IsIDEorEXE <> 0 Then
        VBHdChar(0) = &H56          'V
        VBHdChar(1) = &H42          'B
        VBHdChar(2) = &H35          '5
        VBHdChar(3) = &H21          '!
        BaseAds = VBGetModuleHandle(ByVal 0&) + &H1024&
        CopyMemory MemData(0), ByVal (BaseAds), &H1FDB&
        GetOffSet = InStrB(1, MemData, VBHdChar, vbBinaryCompare)
            If GetOffSet > 0 Then
                VBHeader = GetOffSet + BaseAds - 1
                GETVBHeader = True
            Else
                MessageBox ByVal 0&, StrPtr("无法获取VB数据头！程序将运行的不稳定。"), StrPtr("VB数据头定位标志被篡改！"), vbExclamation
                VBHeader = 0
                GETVBHeader = False
            End If
    Else
    VBHeader = 0
    GETVBHeader = False
    End If
End Function

'********************VBGOOD论坛的老汉(DOWNLOAD)，基于国际友人izero@Slovakia的多线程DLL首创的方法，这里略作修改*******************
Public Function InitVBdll() As Boolean                          '诱导VB6运行库其它部分完全初始化 此函数内容看不懂没关系，直接使用即可
    Dim pIID As mIID, pDummy As Long
    Dim u1 As Long, u2 As Long, u3 As Long
    If VBHeader > 0 Then
            ' Set pIID = IID of IClassFactory
            ' = {00000001-0000-0000-C000-000000000046}
            pIID.data1 = &H1&
            pIID.data4(0) = &HC0
            pIID.data4(7) = &H46
            u3 = VBGetModuleHandle(ByVal 0&)
            'get u1,u2 for VBDllGetClassObject
            UserDllMain u1, u2, u3, 1&, 0&
            VBDllGetClassObject u1, u2, VBHeader, pDummy, pIID, pDummy
            InitVBdll = True
    Else
            InitVBdll = False
    End If

End Function
>>>>>>> 2abdb2e3a24a7144f6e9a89c8664e5d0469ec446
