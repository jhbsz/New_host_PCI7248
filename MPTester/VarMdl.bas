Attribute VB_Name = "VarMdl"
Option Explicit

Public winHwnd As Long
Public OldTimer
Public StopFlag As Byte
Public ChipName As String
Public OldChipName As String
Public NewChipFlag As Byte
Public TestResult As String
Public ContFail As Integer
Public AllenTest As Byte
Public AllenStop As Byte
Public ScanContFail As Byte
Public MPContFail As Byte
Public MPFlag As Byte
Public ResetFlag As Byte
Public ReMP_Counter As Integer
Public Const ReMP_Limit = 2
Public EQC_HV As Boolean
Public EQC_LV As Boolean
Public Dual_Flag As Boolean
Public EQC_Flag As Boolean
Public ForceMP_Flag As Boolean
Public ReMP_Flag As Byte
Public Const MPIdleTime = 12
Public SendGPIBFlag As Boolean
Public DeviceFolderPath As String
Public ShortChipName As String
'Public UpdateModuleName As String
Public CurrentChip As String
Public CurrentMPCaption As String
Public CurrentMPCaption1 As String
Public GPIBCard_Exist As Boolean
Public SetP1CInput_Flag As Boolean
Public SetP1CST4Cond_Flag As Boolean
Public FWFail_Counter As Integer
Public PLFlag As Boolean
Public KLFlag As Boolean

Public Check3826VCC5V As Boolean

Public ResetHubString As String

Public Lon As Boolean, Loff As Boolean
Public U3MPFlag As Integer

Public Function WaitDevOn(Vid_PID As String) As Boolean

Dim PassTime As Single
Dim OldTimer As Single

    OldTimer = Timer
    WaitDevOn = False
    
    Do
        'DoEvents
        If GetDeviceName_NoReply(Vid_PID) <> "" Then
            WaitDevOn = True
        End If
        Call MsecDelay(0.1)
        'DoEvents
        PassTime = Timer - OldTimer
    Loop Until (WaitDevOn) Or (PassTime >= 4)
    
End Function

Public Function WaitDevOFF(Vid_PID As String) As Boolean

Dim PassTime As Single
Dim OldTimer As Single

    OldTimer = Timer
    WaitDevOFF = False
    
    Do
        'DoEvents
        If GetDeviceName_NoReply(Vid_PID) = "" Then
            WaitDevOFF = True
        End If
        Call MsecDelay(0.1)
        'DoEvents
        PassTime = Timer - OldTimer
    Loop Until (WaitDevOFF) Or (PassTime >= 12)
    
End Function

Public Function GetDeviceName_NoReply(Vid_PID As String) As String

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

For MemberIndex = 0 To 10

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
        MyDeviceInterfaceDetailData.cbSize = _
            Len(MyDeviceInterfaceDetailData)
        
        'Use a byte array to allocate memory for
        'the MyDeviceInterfaceDetailData structure
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
 
    If (InStr(1, DevicePathName, Vid_PID)) And (InStr(1, DevicePathName, "6387")) Then
        GetDeviceName_NoReply = DevicePathName
        Exit Function
    ElseIf (InStr(1, DevicePathName, Vid_PID) <> 0) And (InStr(1, DevicePathName, "1234") <> 0) Then
        GetDeviceName_NoReply = DevicePathName
        Exit Function
    ElseIf (InStr(1, DevicePathName, Vid_PID) <> 0) And (InStr(1, DevicePathName, "9420") <> 0) Then
        GetDeviceName_NoReply = DevicePathName
        Exit Function
    ElseIf (InStr(1, DevicePathName, Vid_PID) <> 0) And (InStr(1, DevicePathName, "9380") <> 0) Then
        GetDeviceName_NoReply = DevicePathName
        Exit Function
    ElseIf (InStr(1, DevicePathName, Vid_PID) <> 0) And (InStr(1, DevicePathName, "3822") <> 0) Then
        GetDeviceName_NoReply = DevicePathName
        Exit Function
    ElseIf (InStr(1, DevicePathName, Vid_PID) <> 0) And (InStr(1, DevicePathName, "3826") <> 0) Then
        GetDeviceName_NoReply = DevicePathName
        Exit Function
    ElseIf (InStr(1, DevicePathName, Vid_PID) <> 0) And (InStr(1, DevicePathName, "3825") <> 0) Then
        GetDeviceName_NoReply = DevicePathName
        Exit Function
    ElseIf (InStr(1, DevicePathName, Vid_PID) <> 0) And (InStr(1, DevicePathName, "3821") <> 0) Then
        GetDeviceName_NoReply = DevicePathName
        Exit Function
    ElseIf (InStr(1, DevicePathName, Vid_PID) <> 0) And (InStr(1, DevicePathName, "1000") <> 0) Then
        GetDeviceName_NoReply = DevicePathName
        Exit Function
    Else
        If Left(ChipName, 6) = "AU7310" Then
            GetDeviceName_NoReply = DevicePathName
            Exit Function
        Else
            GetDeviceName_NoReply = ""
        End If
    End If
Next

'    DevicePathName = ""
'    DeviceInfoSet = SetupDiGetClassDevs _
'                    (HidGuid, _
'                    vbNullString, _
'                    0, _
'                    (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
'
'    If MPTester.skip_other_usb.Value = 0 Then
'        MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)
'        result = SetupDiEnumDeviceInterfaces _
'                (DeviceInfoSet, _
'                0, _
'                HidGuid, _
'                MemberIndex, _
'                MyDeviceInterfaceData)
'    Else
'        MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)
'        result = SetupDiEnumDeviceInterfaces _
'                (DeviceInfoSet, _
'                0, _
'                HidGuid, _
'                MPTester.Text2, _
'                MyDeviceInterfaceData)
'    End If
'
'    If result <> 0 Then
'        MyDeviceInfoData.cbSize = Len(MyDeviceInfoData)
'        result = SetupDiGetDeviceInterfaceDetail _
'               (DeviceInfoSet, _
'               MyDeviceInterfaceData, _
'               0, _
'               0, _
'               Needed, _
'               0)
'
'        DetailData = Needed
'
'        'Store the structure's size.
'        MyDeviceInterfaceDetailData.cbSize = _
'        Len(MyDeviceInterfaceDetailData)
'
'        'Use a byte array to allocate memory for
'        'the MyDeviceInterfaceDetailData structure
'        ReDim DetailDataBuffer(Needed)
'
'        Call RtlMoveMemory _
'            (DetailDataBuffer(0), _
'            MyDeviceInterfaceDetailData, _
'            4)
'
'        result = SetupDiGetDeviceInterfaceDetail _
'                (DeviceInfoSet, _
'                MyDeviceInterfaceData, _
'                VarPtr(DetailDataBuffer(0)), _
'                DetailData, _
'                Needed, _
'                0)
'
'        'Convert the byte array to a string.
'        DevicePathName = CStr(DetailDataBuffer())
'
'        'Convert to Unicode.
'        DevicePathName = StrConv(DevicePathName, vbUnicode)
'
'        'Strip cbSize (4 bytes) from the beginning.
'        DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4)
'    End If
'
'    If InStr(1, DevicePathName, Vid_PID) Then
'        GetDeviceName_NoReply = DevicePathName
'    Else
'        GetDeviceName_NoReply = ""
'    End If

End Function

