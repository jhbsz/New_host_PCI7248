Attribute VB_Name = "VaribleDefine"
Option Explicit

Public Const AtheistDebug = False
Public Const AtheistDebugName = "AU9540BSF20"

Public AU6256Unknow
Public buf
Public DeviceHandle As Long
Public HubPort As Long
Public Const GPO_FAIL = 51
Public UsbSpeedTestResult As Integer
Public FlashCapacityError As Integer
Public GPOFail As Integer
Public card  As Integer   'card handle
Public PCI7248InitFinish As Integer
Public PCI7248InitFinish_Sync As Integer
Public TestName(0 To 300) As String
Public BadBlock(0 To 511) As Byte
Public BadBlockCounter As Byte
Public ChipNo As Byte
Public FixCounter As Byte
Public Zone As Byte
Public OS_Result As Byte
Public GPIBCard_Exist As Boolean
Public OSFileName As String

Public ChipName As String
Public DumpTableNormal(0 To 4096) As String
Public DumpTableInverse(0 To 4096) As String
 
 '//// 20050519
Public RequestSenseData(0 To 17) As Byte
Public ReadData(0 To 65536) As Byte
Public MP3Data(0 To 100) As Integer
Public MP3Data_1(0 To 100) As Integer
Public MP3Data_2(0 To 100) As Integer
Public MP3Data_A(0 To 100) As Integer
Public MP3Data_B(0 To 100) As Integer
Public MP3Data_B1(0 To 100) As Integer
Public MP3Data_B2(0 To 100) As Integer
Public MP3Data_C(0 To 100) As Integer
Public MP3Data_3(0 To 100) As Integer
Public MP3Data_C1(0 To 100) As Integer
Public MP3Data_4(0 To 100) As Integer
Public MP3Data_C2(0 To 100) As Integer
Public MP3Data_5(0 To 100) As Integer
Public MP3Data_6(0 To 100) As Integer
Public MP3Data_7(0 To 100) As Integer
Public MP3Data_8(0 To 100) As Integer
Public MP3Data_9(0 To 100) As Integer '=============2007.06.14 CHEYENNE CHANG
Public MP3Data_BL(0 To 100) As Integer
Public MP3Data_BL1(0 To 100) As Integer
Public MP3Data_CW1(0 To 100) As Integer '=============2007.06.14 CHEYENNE CHANG
Public MP3Data_10(0 To 100) As Integer '
Public MP3Data_11(0 To 100) As Integer '
Public MP3Data_12(0 To 100) As Integer
Public MP3Data_13(0 To 100) As Integer

Public MP3Data_15(0 To 100) As Integer
Public MP3Data_16(0 To 100) As Integer
Public MP3Data_161(0 To 100) As Integer
'Public MP3Data_162(0 To 100) As Integer
Public MP3Data_17(0 To 100) As Integer
Public MP3Data_18(0 To 100) As Integer

Public MP3Data_19(0 To 100) As Integer
Public MP3Data_20(0 To 100) As Integer

Public MP3Data_21(0 To 100) As Integer
Public MP3Data_22(0 To 100) As Integer
Public MP3Data_23(0 To 100) As Integer
Public MP3WMA_23(0 To 100) As Integer
Public MP3WMA_231(0 To 100) As Integer
Public MP3WMA_232(0 To 100) As Integer
Public MP3Data_CL(0 To 100) As Integer
Public MP3Data_3150J(0 To 100) As Integer
Public MP3Data_3150J1(0 To 100) As Integer
Public MP3Data_AU6254(0 To 100) As Integer
Public MP3Data_AU62541(0 To 100) As Integer
Public MP3Data_3152A1(0 To 100) As Integer
Public MP3Data_3152A2(0 To 100) As Integer
Public MP3Data_3152A3(0 To 100) As Integer
'Public MP3Data_3152A4(0 To 100) As Integer
Public MP3Data_3150A221(0 To 100) As Integer
Public MP3Data_3150A222(0 To 100) As Integer

Public MP3Data_3150KL1(0 To 100) As Integer
Public MP3Data_3150KL2(0 To 100) As Integer

Public MP3Data_3150KL21(0 To 100) As Integer
Public MP3Data_3150KL22(0 To 100) As Integer
Public MP3Data_3150KL23(0 To 100) As Integer
Public MP3WMA_3150KL23(0 To 100) As Integer
Public MP3WMA_3150KL231(0 To 100) As Integer
Public MP3WMA_3150KL232(0 To 100) As Integer
Public bAlertable As Long
Public Capabilities As HIDP_CAPS
Public DataString As String
Public DetailData As Long
Public DetailDataBuffer() As Byte

Public DeviceAttributes As HIDD_ATTRIBUTES
Public DevicePathName As String
Public DeviceInfoSet As Long
 
Public WriteHandle As Long
Public WriteHandle6331 As Long
Public WriteHandle9369 As Long
Public ReadHandle As Long
Public Rv1ContinueFail As Integer ' for AU6254 testing
 
Public ErrorString As String
Public EventObject As Long
Public HIDHandle As Long
Public HIDOverlapped As OVERLAPPED
Public LastDevice As Boolean

Public MyDeviceDetected As Boolean
Public MyDeviceInfoData As SP_DEVINFO_DATA
Public MyDeviceInterfaceDetailData As SP_DEVICE_INTERFACE_DETAIL_DATA
Public MyDeviceInterfaceData As SP_DEVICE_INTERFACE_DATA
Public Needed As Long
Public OutputReportData(7) As Byte
Public PreparsedData As Long
'public ReadHandle As Long
Public result As Long
Public Security As SECURITY_ATTRIBUTES
Public ReaderExist As Byte
 
  '//// 20050519
 
 Public ArchTest As Boolean
Public TestResultadd As String
Public TestResultSmartCard As String

 
Public bStop As Boolean
Public strAppPath As String
Public lngLineNum As Long
Public bReadyToClose As Boolean
Public NBCount  As Long
Public TimeCounter As Integer
Public TimeCounterBegin As Boolean
Public Present
Public slot As String
Public hDevice As Integer
Public Pattern(0 To 4096) As Byte
Public Pattern_64k(0 To 65536) As Byte
Public Pattern_AU6982(0 To 65536) As Byte
Public Pattern_AU6377(0 To 65536) As Byte
Public Pattern_AU6375(0 To 65536) As Byte
Public AU6371Pattern(0 To 65536) As Byte
Public AU6981Pattern(0 To 4096) As Byte
Public IncPattern(0 To 4096) As Byte
Public a As Integer
Public tmp(2)  As Byte
Public tmp1(2)  As Byte
Public b As Byte
Public flag As Byte
Public ret As Integer
Public ret1 As Integer
Public Declare Function OpenLinkDevice Lib "DevLink.dll" (ByVal hInst As Integer, ByRef pHandle As Integer) As Byte
Public Declare Function WriteIOData Lib "DevLink.dll" (ByVal hDevice As Integer, ByRef pSrc As Byte, ByVal dwSize As Integer, ByRef dwRet As Integer) As Byte
Public Declare Function ReadIOData Lib "DevLink.dll" (ByVal hDevice As Integer, ByRef pSrc As Byte, ByVal dwSize As Integer, ByRef dwRet As Integer) As Byte
Public PCcount  As Long
'\\\\\\\\\\\\\\\\\\\\\\\
Public CBW(0 To 30) As Byte
Public CSW(0 To 13) As Byte
Public CBWDataTransferLen(0 To 3) As Byte
Public LBAByte(0 To 3) As Byte
Public TransData(0 To 4096) As Byte

Public Lun As Byte
Public opcode As Byte
Public LBA As Long
Public CBWFlag As Byte
Public CDBLen As Byte
Public Const SD_Disk = "F:\"
Public Const CF_Disk = "G:\"
Public Const XD_Disk = "H:\"
Public Const MS_Disk = "I:\"

Public value_a(0 To 1) As Long, value_b(0 To 1) As Long, value_cu(0 To 1) As Long, value_cl(0 To 1) As Long
Public status_a(0 To 1) As Integer, status_b(0 To 1) As Integer, status_cu(0 To 1) As Long, status_cl(0 To 1) As Integer
Public CardResult As Integer   'for card control result for ADLINK PCI card
Public LightOff As Long
Public LightOn As Long
Public TestResult As String
Public PreviousStatus As Byte
Public SaveOSCounter As Long
Public CurrentOSFileName As String

Public rv As Byte
Public rv0 As Byte
Public rv1 As Byte
Public rv2 As Byte
Public rv3 As Byte
Public rv4 As Byte
Public rv5 As Byte
Public rv6 As Byte
Public rv7 As Byte



Public InquiryString(0 To 3, 0 To 36) As Byte
Public AU6375ASHangTime As Single
Public OldTimer
Public Const UNKNOW = 0
Public Const PASS = 1
Public Const WRITE_FAIL = 2
Public Const READ_FAIL = 3
Public Const PREVIOUS_SLOT_FAIL = 4

Public SDWriteFail As Long
Public SDReadFail As Long

Public CFWriteFail As Long
Public CFReadFail As Long

Public XDWriteFail As Long
Public XDReadFail As Long

Public MSWriteFail As Long
Public MSReadFail As Long

Public MiniSDWriteFail As Long
Public MiniSDReadFail As Long
Public UnknowDeviceFail As Long
Public AU6254TestMsg As String
Public OldChipName As String
Public AU6258ProgVer As String
Public PreChipName As String
 

Public mfn As String

Public CMediaTestResult As Byte
Public CMediaDiappearCounter As Byte
Public CMediaPassCounter As Byte
Public Declare Function CAndValue Lib "DLL7.dll" (ByVal In1 As Byte, ByVal In2 As Byte) As Integer
'Declare Function D2K_Register_Card Lib "D2K-Dask.dll" (ByVal cardType As Integer, ByVal card_num As Integer) As Integer
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

' INI file
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpapplicationname As String, ByVal lpkeyname As Any, ByVal lpstring As Any, ByVal lpfilename As String) As LoadPictureColorConstants


'Sync for different site
Public Const SiteReady = &H0
Public Const RunMP = &H1
Public Const MPDone = &H2
Public Const RunHV = &H3
Public Const HVDone = &H4
Public Const RunLV = &H5
Public Const LVDone = &H6
Public Const SiteUnknow = &HF


Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
End Type

Public bNeedsReStart As Boolean

Public startProgress As Boolean

'for Auto re-open Tester.exe (AU9540 testing will cause memory leak)
Public Sub CheckAvailMemSize(minTargetMB)
Dim MemBuffers As MEMORYSTATUS
Dim lTotalMem As Long
Dim lAvailMem As Long

    MemBuffers.dwLength = Len(MemBuffers)
    
    GlobalMemoryStatus MemBuffers
    'lTotalMem = MemBuffers.dwTotalPhys \ 1024 \ 1024
    lAvailMem = MemBuffers.dwAvailPhys \ 1024 \ 1024
    
    If (lAvailMem <= minTargetMB) Then
        Shell App.Path & "\RestartTester.exe", vbNormalFocus
        End
    End If
    

End Sub




Public Sub SetSiteStatus(SetVal As Long)

Dim StatusOutCH As Integer

    If GPIBCard_Exist Then
        StatusOutCH = Channel_P1CH
    Else
        StatusOutCH = Channel_P1CL
    End If
    
    'DoEvents
    CardResult = DO_WritePort(card, StatusOutCH, SetVal)
    
End Sub

Public Function GetAnOtherStatus() As Byte

Dim StatusInCH As Integer
Dim StatusVal As Long
    
    GetAnOtherStatus = &HF
    
    If GPIBCard_Exist Then
        StatusInCH = Channel_P1CL
    Else
        StatusInCH = Channel_P1CH
    End If
    
    'DoEvents
    CardResult = DO_ReadPort(card, StatusInCH, StatusVal)
    GetAnOtherStatus = StatusVal
    
End Function

Public Sub WaitAnotherSiteDone(FlowItem As Long, SetTimeOut As Single)

Dim ReadVal As Long
Dim PassTime As Single
Dim OldTimer As Single

    OldTimer = Timer

    Do
        ReadVal = GetAnOtherStatus
        If (ReadVal = SiteUnknow) Or (ReadVal = FlowItem) Then
            Exit Sub
        End If
        Call MsecDelay(0.06)
        PassTime = Timer - OldTimer
        
        If Timer < OldTimer Then
            Exit Do
        End If
        
    Loop Until (ReadVal = FlowItem) Or (PassTime > SetTimeOut)
    
End Sub
