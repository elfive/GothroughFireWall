<<<<<<< HEAD
Attribute VB_Name = "VBMultiThreadRun"
Option Explicit

'线程之间“交流”最好的方式是用加锁的公共变量（byte boolen integer long可以不加锁）或者信号变量等等。
'给线程传递参数并不是高效方便的“交流“方式，因为参数传递会由创建线程的API经手。也就是说参数先传入API中，下一步才会传入子线程。
'某些情况下确实要在线程创建时给它传递一个或几个私有变量，但是API只留了一个参数的位置给我们传递。这意味着要传递多个变量就必须将其
'做成自定义数据类型。具体方法见本示例中的代码。
'特别注意：程序必须编译为EXE后，传递的参数才是正确的。在IDE下除4字节以下的数字外，其余的会数据混乱。


Public ThreadMonitorEnd As Boolean          '线程结束标识，用于主线程判断子线程是否已经结束
Public VBThreadHandle As Long               '定义线程句柄
Public VBThreadID As Long                   '定义线程ID


'************************************注意：VB6多线程必须以SUB MAIN为启动对象***************************************************
'***************************本示例中已经设置好了，自己使用时注意在工程——属性——启动对象中自行选择************************
Sub Main()
If AvoidReentrant = False Then       '防止主线程重复运行
    AvoidReentrant = True
        If App.PrevInstance Then        '防止程序重复运行
            MessageBox ByVal 0&, StrPtr("程序正在运行或未完全退出"), StrPtr("重复运行"), vbCritical
            Exit Sub
        Else
            InitCommonControls      '初始化通用控件
            GETVBHeader                 '获取VB数据头
            
            Form1.Show     '在这里加载主窗体
        End If
End If
End Sub

Public Sub SubIsConnectedState(ByRef isConnected As Boolean)        '子线程1
'***********************************（重要！）VB6线程环境初始化*************************************************
CreateIExprSrvObj 0&, 4&, 0&            'VB6运行库初始化
CoInitializeEx ByVal 0&, ByVal (COINIT_MULTITHREADED Or COINIT_SPEED_OVER_MEMORY)   'COM组件初始化
InitVBdll               '诱导VB6运行库内部其他部分的初始化
'***********************************（重要！）VB6线程环境初始化*************************************************

'**********************************************自定义程序主体***************************************************

isConnected = CheckConnection

ThreadMonitorEnd = True
'**********************************************自定义程序主体***************************************************
CoUninitialize      '卸载COM组件（省掉也不会影响稳定性，但可能造成句柄或内存泄漏。为了养成好习惯，还是写上）
End Sub
=======
Attribute VB_Name = "VBMultiThreadRun"
Option Explicit

'线程之间“交流”最好的方式是用加锁的公共变量（byte boolen integer long可以不加锁）或者信号变量等等。
'给线程传递参数并不是高效方便的“交流“方式，因为参数传递会由创建线程的API经手。也就是说参数先传入API中，下一步才会传入子线程。
'某些情况下确实要在线程创建时给它传递一个或几个私有变量，但是API只留了一个参数的位置给我们传递。这意味着要传递多个变量就必须将其
'做成自定义数据类型。具体方法见本示例中的代码。
'特别注意：程序必须编译为EXE后，传递的参数才是正确的。在IDE下除4字节以下的数字外，其余的会数据混乱。


Public ThreadMonitorEnd As Boolean          '线程结束标识，用于主线程判断子线程是否已经结束
Public VBThreadHandle As Long               '定义线程句柄
Public VBThreadID As Long                   '定义线程ID


'************************************注意：VB6多线程必须以SUB MAIN为启动对象***************************************************
'***************************本示例中已经设置好了，自己使用时注意在工程——属性——启动对象中自行选择************************
Sub Main()
If AvoidReentrant = False Then       '防止主线程重复运行
    AvoidReentrant = True
        If App.PrevInstance Then        '防止程序重复运行
            MessageBox ByVal 0&, StrPtr("程序正在运行或未完全退出"), StrPtr("重复运行"), vbCritical
            Exit Sub
        Else
            InitCommonControls      '初始化通用控件
            GETVBHeader                 '获取VB数据头
            
            Form1.Show     '在这里加载主窗体
        End If
End If
End Sub

Public Sub SubIsConnectedState(ByRef isConnected As Boolean)        '子线程1
'***********************************（重要！）VB6线程环境初始化*************************************************
CreateIExprSrvObj 0&, 4&, 0&            'VB6运行库初始化
CoInitializeEx ByVal 0&, ByVal (COINIT_MULTITHREADED Or COINIT_SPEED_OVER_MEMORY)   'COM组件初始化
InitVBdll               '诱导VB6运行库内部其他部分的初始化
'***********************************（重要！）VB6线程环境初始化*************************************************

'**********************************************自定义程序主体***************************************************

isConnected = CheckConnection

ThreadMonitorEnd = True
'**********************************************自定义程序主体***************************************************
CoUninitialize      '卸载COM组件（省掉也不会影响稳定性，但可能造成句柄或内存泄漏。为了养成好习惯，还是写上）
End Sub
>>>>>>> 2abdb2e3a24a7144f6e9a89c8664e5d0469ec446
