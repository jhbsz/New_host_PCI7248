Attribute VB_Name = "AlcorMP"
'1. must load MPFiler.sys to system32/driver
'2. replace 29_k9HCG.bin at /86 subdirectory

Option Explicit
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const GWL_WNDPROC = (-4)
Public prevWndProc As Long

Public Declare Function fnScsi2usb2K_KillEXE Lib "usbreset.dll" () As Integer
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
 
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const PROCESS_TERMINATE = &H1
Public Const WM_CLOSE = &H10
Public Const WM_DESTROY = &H2
Public Const WM_QUIT = &H12
Public Const INFINITE = &HFFFF      '  Infinite timeout

Public Const SW_SHOW = 5
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Declare Function WaitMessage Lib "user32" () As Boolean
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long

Public Const PM_NOREMOVE = &H0
Public Const PM_REMOVE = &H1

Const SWP_NOMOVE = &H2 '不更動目前視窗位置
Const SWP_NOSIZE = &H1 '不更動目前視窗大小
Public Const HWND_TOPMOST = -1 '設定為最上層
Const HWND_NOTOPMOST = -2 '取消最上層設定
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
'ExitWindowsEx EWX_FORCE Or EWX_SHUTDOWN, 0

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

Const TH32CS_SNAPPROCESS = 2
Const MAX_PATH = 260

Private Type PROCESSENTRY32
    dwSize               As Long
    cntUsage             As Long
    th32ProcessID        As Long
    th32DefaultHeapID    As Long
    th32ModuleID         As Long
    cntThreads           As Long
    th32ParentProcessID  As Long
    pcPriClassBase       As Long
    dwFlags              As Long
    szexeFile            As String * MAX_PATH
End Type

Public AU6254CMediaFailCounter As Byte

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Function WndProc(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    AlcorMPMessage = MSG
    'debug.print hwnd, MSG, wParam, lParam
    WndProc = CallWindowProc(prevWndProc, hwnd, MSG, wParam, lParam)
    
End Function

Public Sub AU6988DLF20TestSub()

Dim OldTime
Dim PassTime

'1. if continueFail >=5 then executte " Alcor Micro USF manufacture Program
'1a. load MP , until AlcorMPtest.status= "ReadY'

    OldTime = Timer
    MsgBox "begin"

'2. execute RW.tester

End Sub

Public Function GetDeviceName(Vid_PID As String) As String

Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long
 
    With HidGuid ' CD-ROM Gerateklassen GUID
        .Data1 = &HA5DCBF10
        .Data2 = &H6530
        .Data3 = &H11D2
        .Data4(0) = &H90
        .Data4(1) = &H1F
        .Data4(2) = &H0
        .Data4(3) = &HC0
        .Data4(4) = &H4F
        .Data4(5) = &HB9
        .Data4(6) = &H51
        .Data4(7) = &HED
    End With

    DevicePathName = ""
    DeviceInfoSet = SetupDiGetClassDevs _
                (HidGuid, _
                vbNullString, _
                0, _
                (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
    
    MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)
    
    result = SetupDiEnumDeviceInterfaces _
                (DeviceInfoSet, _
                0, _
                HidGuid, _
                MemberIndex, _
                MyDeviceInterfaceData)
        
    If result <> 0 Then
        MyDeviceInfoData.cbSize = Len(MyDeviceInfoData)
            
        result = SetupDiGetDeviceInterfaceDetail _
                (DeviceInfoSet, _
                MyDeviceInterfaceData, _
                0, _
                0, _
                Needed, _
                0)
        
        DetailData = Needed
        
        'Store the structure's size.
        MyDeviceInterfaceDetailData.cbSize = Len(MyDeviceInterfaceDetailData)
        
        'Use a byte array to allocate memory for the MyDeviceInterfaceDetailData structure
        ReDim DetailDataBuffer(Needed)
        
        Call RtlMoveMemory _
            (DetailDataBuffer(0), _
            MyDeviceInterfaceDetailData, _
            4)
            
        result = SetupDiGetDeviceInterfaceDetail _
                   (DeviceInfoSet, _
                   MyDeviceInterfaceData, _
                   VarPtr(DetailDataBuffer(0)), _
                   DetailData, _
                   Needed, _
                   0)
           
        'Convert the byte array to a string.
        DevicePathName = CStr(DetailDataBuffer())
        
        'Convert to Unicode.
        DevicePathName = StrConv(DevicePathName, vbUnicode)
        
        'Strip cbSize (4 bytes) from the beginning.
        DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4)
 
    End If
 
    If InStr(1, DevicePathName, Vid_PID) Or InStr(1, DevicePathName, "9101") Or InStr(1, DevicePathName, "9102") Then
        GetDeviceName = DevicePathName
        MPTester.Print "VID="; Mid(DevicePathName, InStr(1, DevicePathName, "vid"), 8); " ; ";
        MPTester.Print "PID="; Mid(DevicePathName, InStr(1, DevicePathName, "pid"), 8)
    Else
        GetDeviceName = ""
        MPTester.Print "VID= unknow"; " ; ";
        MPTester.Print "PID= unknow"
    End If

End Function

Public Sub KillProcess(NameProcess As String)

Const PROCESS_ALL_ACCESS = &H1F0FFF
Const TH32CS_SNAPPROCESS As Long = 2&
Dim uProcess  As PROCESSENTRY32
Dim RProcessFound As Long
Dim hSnapshot As Long
Dim SzExename As String
Dim ExitCode As Long
Dim MyProcess As Long
Dim AppKill As Boolean
Dim AppCount As Integer
Dim i As Integer
Dim WinDirEnv As String
        
    If NameProcess <> "" Then
        AppCount = 0

        uProcess.dwSize = Len(uProcess)
        hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
        RProcessFound = ProcessFirst(hSnapshot, uProcess)
  
        Do
            i = InStr(1, uProcess.szexeFile, Chr(0))
            SzExename = LCase$(Left$(uProcess.szexeFile, i - 1))
            WinDirEnv = Environ("Windir") + "\"
            WinDirEnv = LCase$(WinDirEnv)
        
            If Right$(SzExename, Len(NameProcess)) = LCase$(NameProcess) Then
                AppCount = AppCount + 1
                MyProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
                AppKill = TerminateProcess(MyProcess, ExitCode)
                Call CloseHandle(MyProcess)
            End If
            RProcessFound = ProcessNext(hSnapshot, uProcess)
        Loop While RProcessFound
        Call CloseHandle(hSnapshot)
        
    End If

End Sub
