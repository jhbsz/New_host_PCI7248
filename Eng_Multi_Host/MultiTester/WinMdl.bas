Attribute VB_Name = "WinMdl"
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowContextHelpId Lib "user32" (ByVal hwnd As Long, ByVal dw As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOMOVE = &H2 '����ʥثe������m
Public Const SWP_NOSIZE = &H1 '����ʥثe�����j�p
Public Const HWND_TOPMOST = -1 '�]�w���̤W�h
Public Const HWND_NOTOPMOST = -2 '�����̤W�h�]�w
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4

Public NameofPC  As String
Public ProgramName As String
Public ProgramRevisionCode As String
Public DeviceID As String
Public RunCardNO As String
Public LotID As String
Public StartAt As String
Public StartAtMin As String
Public EndAt As String
Public EndAtMax As String
Public HandlerID As String
Public OperatorName As String
Public ProcessID As String
Public Sites As String
 
Public Const AllenDebug = 0
Public Const ReportDebug = 0
Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '�`?�G  Maintenance string for PSS usage
End Type

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
        (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias _
        "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long




Public Bin1Counter(0 To 7) As Long
Public Bin2Counter(0 To 7) As Long
Public Bin3Counter(0 To 7) As Long
Public Bin4Counter(0 To 7) As Long
Public Bin5Counter(0 To 7) As Long

Public ReportBegin As Byte

'==========================================='
Public Bin1Sum(0 To 7) As Long
Public Bin2Sum(0 To 7) As Long
Public Bin3Sum(0 To 7) As Long
Public Bin4Sum(0 To 7) As Long
Public Bin5Sum(0 To 7) As Long
 
'=========================================
  Public EndDay As String
  Public EndSecond As String
  Public SNow As String
  Public OutFileName As String
  
  Public ProcessIDSum As String
  
  
   
  



 
