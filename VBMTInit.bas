<<<<<<< HEAD
Attribute VB_Name = "VBMTInit"
'*******************************************************************************************************************************
'��ԭ���Ĵ��벿���ѱ���ԭ���߼����������ಿ�ֽ�Ϊ���ˡ�������������ٶ�ID:hhyjq007���Լ��о�����
'��л�Ϻ���TGY��IZERO�ȵ�ΪVB6�������ȶ����߳���ƽ��·��ǰ���ǣ�
'�����κ����ʿɵ���VBGOOD��̳����ٶ����ɡ�VB�ɡ�����������
'*******************************************************************************************************************************
Option Explicit

'ClassID
Public Type mIID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(7) As Byte
End Type

'�̰߳�ȫ�������ݽṹ������������ͨ���ò�����
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Byte
End Type

'�ٽ������ݽṹ����������
Public Type CRITICAL_SECTION
    DebugInfo As Long
    LockCount As Long
    RecursionCount As Long
    OwningThread As Long
    LockSemaphore As Long
    SpinCount As Long
End Type

'*�ؼ���VB6�ڲ����غ�������VBGOOD��̳���Ϻ���DOWNLOAD���״η��֣��ܳ�ʼ��VB6���п��о��󲿷ֵ�����
Public Declare Function CreateIExprSrvObj Lib "msvbvm60.dll" (ByVal p1_0 As Long, ByVal p2_4 As Long, ByVal p3_0 As Long) As Long

'�����߳�API�������VB6��������
Public Declare Function CreateThread Lib "kernel32" (ByVal lpSecurityAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Const CREATE_DEFAULT = &H0&          'Ĭ��ֵ������һ���������е��߳�
Public Const CREATE_SUSPENDED = &H4&        '�̴߳������������У���ҪResumeThread����ſ�ʼ����
Public Const STACK_SIZE_CUSTOM = &H10000    '�Զ����ջ��С��StackSize�������ֽ�Ϊ��λ

Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long        '��ȡ��ǰ�߳�ID
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long       '��ȡ��ǰ����ID
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long     '�������б���ͣ���߳�
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long    '��ͣĳ���̵߳����У�ʹ��ʱע���ٽ������⣡ĳ�߳̽����ٽ�������ͣ�ᵼ���ٽ���������ռ�á�ֱ���̱߳��������е��뿪�ٽ�����
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long    'ǿ����ֹ�̣߳�ǿ�ҽ���ֻ�ڳ���Ҫȫ���˳�ʱʹ�ã�ĳ�߳̽����ٽ�����ǿ���˳��ᵼ���ٽ���������ռ�ã�����û�в��ȴ�ʩ����
Public Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '��ʼ���ٽ���
Public Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  'ɾ���ٽ���
Public Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '�����ٽ���
Public Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '�뿪�ٽ���

Public Declare Function VBGetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModName As Long) As Long     '��ȡģ���ַ/���
Public Declare Sub UserDllMain Lib "msvbvm60.dll" (u1 As Long, u2 As Long, ByVal u3_h As Long, ByVal u4_1 As Long, ByVal u5_0 As Long)      'VB6�ڲ����غ��� �û�DLL���
Public Declare Function VBDllGetClassObject Lib "msvbvm60.dll" (g1 As Long, g2 As Long, ByVal g3_vbHeader As Long, REFCLSID As Long, REFIID As mIID, ppv As Long) As Long   ''VB6�ڲ����غ��� ��ȡ�����

Public Declare Sub InitCommonControls Lib "comctl32" ()     '��ʼ��ͨ�ÿؼ�
Public Declare Function CoInitializeEx Lib "ole32.dll" (ByVal pvReserved As Long, ByVal dwCoInit As Long) As Long       '��ʼ��COM���
Public Const COINIT_APARTMENTTHREADED = &H2&            '�Ե��߳�ģʽ����COM���ģ��
Public Const COINIT_MULTITHREADED = &H0&                    '�Զ��߳�ģʽ����COM���ģ��
Public Const COINIT_DISABLE_OLE1DDE = &H4&                  '�ر�OLE1DDE
Public Const COINIT_SPEED_OVER_MEMORY = &H8&            '����������ڴ�Ϊ���ۣ���COM���ģ�����еĸ��죨ǿ���Ƽ�����

Public Declare Sub CoUninitialize Lib "ole32.dll" ()                'ж��COM���

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long      '�رվ��
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)        '�����ڴ�
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)          '�߳�����һ��ʱ��

'Unicode��Ϣ��API
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxW" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
'���¶��Ƿ���ֵ�ĳ���
Public Const IDOK = 1
Public Const IDCANCEL = 2
Public Const IDABORT = 3
Public Const IDRETRY = 4
Public Const IDIGNORE = 5
Public Const IDYES = 6
Public Const IDNO = 7
Public Const IDPROMPT = &HFFFF&

'           '��ֹ�����־                    '������IDE��EXE��־    '���VB����ͷ
Public AvoidReentrant  As Boolean, IsIDEorEXE As Long, VBHeader As Long

'***********************************VBGOOD��̳��TGY�״��Ķ�ȡVB6����ͷ�ķ��������������޸ĺ��Ż�*******************
Public Function GETVBHeader() As Boolean            '��ȡVB6����ͷ          �˺������ݿ�����û��ϵ��ֱ��ʹ�ü���
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
                MessageBox ByVal 0&, StrPtr("�޷���ȡVB����ͷ���������еĲ��ȶ���"), StrPtr("VB����ͷ��λ��־���۸ģ�"), vbExclamation
                VBHeader = 0
                GETVBHeader = False
            End If
    Else
    VBHeader = 0
    GETVBHeader = False
    End If
End Function

'********************VBGOOD��̳���Ϻ�(DOWNLOAD)�����ڹ�������izero@Slovakia�Ķ��߳�DLL�״��ķ��������������޸�*******************
Public Function InitVBdll() As Boolean                          '�յ�VB6���п�����������ȫ��ʼ�� �˺������ݿ�����û��ϵ��ֱ��ʹ�ü���
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
'��ԭ���Ĵ��벿���ѱ���ԭ���߼����������ಿ�ֽ�Ϊ���ˡ�������������ٶ�ID:hhyjq007���Լ��о�����
'��л�Ϻ���TGY��IZERO�ȵ�ΪVB6�������ȶ����߳���ƽ��·��ǰ���ǣ�
'�����κ����ʿɵ���VBGOOD��̳����ٶ����ɡ�VB�ɡ�����������
'*******************************************************************************************************************************
Option Explicit

'ClassID
Public Type mIID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(7) As Byte
End Type

'�̰߳�ȫ�������ݽṹ������������ͨ���ò�����
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Byte
End Type

'�ٽ������ݽṹ����������
Public Type CRITICAL_SECTION
    DebugInfo As Long
    LockCount As Long
    RecursionCount As Long
    OwningThread As Long
    LockSemaphore As Long
    SpinCount As Long
End Type

'*�ؼ���VB6�ڲ����غ�������VBGOOD��̳���Ϻ���DOWNLOAD���״η��֣��ܳ�ʼ��VB6���п��о��󲿷ֵ�����
Public Declare Function CreateIExprSrvObj Lib "msvbvm60.dll" (ByVal p1_0 As Long, ByVal p2_4 As Long, ByVal p3_0 As Long) As Long

'�����߳�API�������VB6��������
Public Declare Function CreateThread Lib "kernel32" (ByVal lpSecurityAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Const CREATE_DEFAULT = &H0&          'Ĭ��ֵ������һ���������е��߳�
Public Const CREATE_SUSPENDED = &H4&        '�̴߳������������У���ҪResumeThread����ſ�ʼ����
Public Const STACK_SIZE_CUSTOM = &H10000    '�Զ����ջ��С��StackSize�������ֽ�Ϊ��λ

Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long        '��ȡ��ǰ�߳�ID
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long       '��ȡ��ǰ����ID
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long     '�������б���ͣ���߳�
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long    '��ͣĳ���̵߳����У�ʹ��ʱע���ٽ������⣡ĳ�߳̽����ٽ�������ͣ�ᵼ���ٽ���������ռ�á�ֱ���̱߳��������е��뿪�ٽ�����
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long    'ǿ����ֹ�̣߳�ǿ�ҽ���ֻ�ڳ���Ҫȫ���˳�ʱʹ�ã�ĳ�߳̽����ٽ�����ǿ���˳��ᵼ���ٽ���������ռ�ã�����û�в��ȴ�ʩ����
Public Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '��ʼ���ٽ���
Public Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  'ɾ���ٽ���
Public Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '�����ٽ���
Public Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '�뿪�ٽ���

Public Declare Function VBGetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModName As Long) As Long     '��ȡģ���ַ/���
Public Declare Sub UserDllMain Lib "msvbvm60.dll" (u1 As Long, u2 As Long, ByVal u3_h As Long, ByVal u4_1 As Long, ByVal u5_0 As Long)      'VB6�ڲ����غ��� �û�DLL���
Public Declare Function VBDllGetClassObject Lib "msvbvm60.dll" (g1 As Long, g2 As Long, ByVal g3_vbHeader As Long, REFCLSID As Long, REFIID As mIID, ppv As Long) As Long   ''VB6�ڲ����غ��� ��ȡ�����

Public Declare Sub InitCommonControls Lib "comctl32" ()     '��ʼ��ͨ�ÿؼ�
Public Declare Function CoInitializeEx Lib "ole32.dll" (ByVal pvReserved As Long, ByVal dwCoInit As Long) As Long       '��ʼ��COM���
Public Const COINIT_APARTMENTTHREADED = &H2&            '�Ե��߳�ģʽ����COM���ģ��
Public Const COINIT_MULTITHREADED = &H0&                    '�Զ��߳�ģʽ����COM���ģ��
Public Const COINIT_DISABLE_OLE1DDE = &H4&                  '�ر�OLE1DDE
Public Const COINIT_SPEED_OVER_MEMORY = &H8&            '����������ڴ�Ϊ���ۣ���COM���ģ�����еĸ��죨ǿ���Ƽ�����

Public Declare Sub CoUninitialize Lib "ole32.dll" ()                'ж��COM���

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long      '�رվ��
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)        '�����ڴ�
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)          '�߳�����һ��ʱ��

'Unicode��Ϣ��API
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxW" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
'���¶��Ƿ���ֵ�ĳ���
Public Const IDOK = 1
Public Const IDCANCEL = 2
Public Const IDABORT = 3
Public Const IDRETRY = 4
Public Const IDIGNORE = 5
Public Const IDYES = 6
Public Const IDNO = 7
Public Const IDPROMPT = &HFFFF&

'           '��ֹ�����־                    '������IDE��EXE��־    '���VB����ͷ
Public AvoidReentrant  As Boolean, IsIDEorEXE As Long, VBHeader As Long

'***********************************VBGOOD��̳��TGY�״��Ķ�ȡVB6����ͷ�ķ��������������޸ĺ��Ż�*******************
Public Function GETVBHeader() As Boolean            '��ȡVB6����ͷ          �˺������ݿ�����û��ϵ��ֱ��ʹ�ü���
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
                MessageBox ByVal 0&, StrPtr("�޷���ȡVB����ͷ���������еĲ��ȶ���"), StrPtr("VB����ͷ��λ��־���۸ģ�"), vbExclamation
                VBHeader = 0
                GETVBHeader = False
            End If
    Else
    VBHeader = 0
    GETVBHeader = False
    End If
End Function

'********************VBGOOD��̳���Ϻ�(DOWNLOAD)�����ڹ�������izero@Slovakia�Ķ��߳�DLL�״��ķ��������������޸�*******************
Public Function InitVBdll() As Boolean                          '�յ�VB6���п�����������ȫ��ʼ�� �˺������ݿ�����û��ϵ��ֱ��ʹ�ü���
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
