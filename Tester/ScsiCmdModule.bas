Attribute VB_Name = "ScsiCmdModule"
Public Const Direction_In = &H80
Public Const Direction_Out = &H0
Public Const UnloadDisk = &H8B
Public Const Write_Password = &H71
Public Const Read_Password = &H72
Public Const Erase_Password = &H73
Public Const Send_Password = &H74
Public Const Check_Password = &H75
Public Const Write_Password_2 = &H76
Public Const Read_1Mega = &H7A
Public Const Write_1Mega = &H7B
Public Const Write_VID = &H81
Public Const Read_VID = &H82
Public Const Erase_VID = &H83
Public Const DiagnosticCmd = &H91
Public Const LowLevelCmd = &H92
Public Const Low_Format = &H92
Public Const GetChipNumber = &H93
Public Const SingleZoneDiagnostic = &H94
Public Const SingleZoneLowFormat = &H95
Public Const GetConfiguration = &H96
Public Const Software_Reset = &H97
Public Const GetScanInfo = &H98
Public Const Get_FAT = &H99
Public Const Read_OrignialCapacity = &H9A
Public Const SET_BOOTABLE = &H9B
Public Const PHYSICAL_READ = &H9C
Public Const DUMP_FAT = &H9D
Public Const DUMP_CACHE = &H9E
Public Const DUMP_CONFIGURATION = &H9F
Public Const FIXED_TYPE_ONLY = 1
 
Public Type Dev
       DeviceCode  As Byte
       ZoneNumber As Byte
       BlockNumber As Integer
       SectorNumber As Integer
End Type
 
Public Declare Function fnScsi2usb_ReadCapacity _
       Lib "usb2scsi.dll" _
       (ByVal filehandle As Long) _
As Integer
 
Public Declare Function fnScsi2usb_ModeSense _
       Lib "usb2scsi.dll" _
       (ByVal filehandle As Long) _
As Integer

Public Declare Function fnScsi2usb_Inquiry _
       Lib "usb2scsi.dll" _
       (ByVal filehandle As Long) _
As Long

Public Declare Function fnScsi2usb_Read _
       Lib "usb2scsi.dll" _
       (ByVal filehandle As Long, _
        ByVal TL As Long, _
        ByVal LBA As Long) _
As Integer
       
Public Declare Function fnScsi2usb_Write _
       Lib "usb2scsi.dll" _
       (ByVal filehandle As Long, _
        ByVal TL As Long, _
        ByVal LBA As Long, _
        DataBuffer As Any) _
As Integer

Public Declare Function fnScsi2usb_TestUnitReady _
       Lib "usb2scsi.dll" _
       (ByVal filehandle As Long) _
As Integer

Public Declare Function fnScsi2usb_VendorR _
       Lib "usb2scsi.dll" _
       (ByVal filehandle As Long, _
       ByVal opcode As Long, _
       ByVal arg1 As Long, _
            ByVal TL As Long _
    ) _
As Integer

Public Declare Function fnScsi2usb_Vendor _
       Lib "usb2scsi.dll" _
       (ByVal filehandle As Long, _
       ByVal opcode As Long, _
       ByVal arg1 As Long, _
       ByVal arg2 As Long, _
       ByVal TL As Long _
    ) _
As Integer

Public Declare Function fnScsi2usb_VendorScan _
       Lib "usb2scsi.dll" _
       (ByVal filehandle As Long, _
       ByVal opcode As Long, _
       ByVal arg1 As Long, _
       ByVal arg2 As Long, _
       ByVal arg3 As Long, _
       ByVal arg4 As Long, _
       ByVal arg5 As Long, _
       ByVal TL As Long _
    ) _
As Integer

Public Declare Function fnScsi2usb2K_KillEXE _
       Lib "usbreset.dll" _
       () _
As Integer

Public Declare Function fnScsi2usb_VendorW _
       Lib "usb2scsi.dll" _
       (ByVal filehandle As Long, _
       ByVal opcode As Long, _
       ByVal TL As Long, _
       DataBuffer As Any _
    ) _
As Integer
 
Public Function SetConfiguration(DiskIDnumber As Integer, WriteToDisk As Boolean)
    Dim result As Integer, pointer As Integer
    Dim start, Finish, TotalTime
    Dim FileName As String
    Dim checksum As Integer
    Dim bArr() As Byte, bArr1() As Byte
    Dim iStringLen As Integer, iStringLen1 As Integer
    Dim ScsiRevTxt As String
    On Error Resume Next
    
    ScsiRevTxt = "7.77"
    'initialize DatabufferArray
    For i = 1 To 512
        DataBufferArray(i) = 0
    Next i
    'check null exist or not
    If (VidTxt <> "" And PidTxt <> "" And _
       UsbManufacturerTxt <> "" And UsbProductTxt <> "" And _
       ScsiManufacturerTxt <> "" And ScsiProductTxt <> "") Then
       DataBufferArray(1) = 153    '0x99
       DataBufferArray(2) = 7      '0x07
       ' Unicode & String length & String type
       DataBufferArray(3) = Len(UsbManufacturerTxt) * 2 + 2
       DataBufferArray(4) = Len(UsbProductTxt) * 2 + 2
       DataBufferArray(5) = 18     '0x12
       DataBufferArray(6) = 1      '0x01
       DataBufferArray(7) = 16     '0x10
       DataBufferArray(8) = 1      '0x01
       DataBufferArray(9) = 0      '0x00
       DataBufferArray(10) = 0     '0x00
       DataBufferArray(11) = 0     '0x00
       DataBufferArray(12) = 8     '0x08
       DataBufferArray(13) = Val("&H" & (Mid(VidTxt, 3, 2)))
       DataBufferArray(14) = Val("&H" & (Mid(VidTxt, 1, 2)))
       DataBufferArray(15) = Val("&H" & (Mid(PidTxt, 3, 2)))
       DataBufferArray(16) = Val("&H" & (Mid(PidTxt, 1, 2)))
       DataBufferArray(17) = 0     '0x00
       DataBufferArray(18) = 1     '0x01
       DataBufferArray(19) = 1     '0x01
       DataBufferArray(20) = 2     '0x02
       DataBufferArray(21) = 0     '0x00
       DataBufferArray(22) = 1     '0x01
       
       DataBufferArray(23) = Len(UsbManufacturerTxt) * 2 + 2
       DataBufferArray(24) = 3
       For i = 1 To Len(UsbManufacturerTxt)
              DataBufferArray(2 * i - 1 + 24) = Asc(Mid(UsbManufacturerTxt, i, 1))
       Next i
       pointer = 2 * (i - 1) + 25 'i=len(usbmanufacturetxt.text)+1
       DataBufferArray(pointer) = Len(UsbProductTxt) * 2 + 2
       DataBufferArray(pointer + 1) = 3 '33
       For i = 1 To Len(UsbProductTxt)
           DataBufferArray(2 * i - 1 + pointer + 1) = Asc(Mid(UsbProductTxt, i, 1))
       Next i
       
       pointer = pointer + 2 * (i - 1) + 2
       bArr = StrConv(ScsiManufacturerTxt, vbFromUnicode)
       iStringLen = LenB(StrConv(ScsiManufacturerTxt, vbFromUnicode))
       For i = 0 To 7
           If (i < iStringLen) Then
               DataBufferArray(i + pointer) = bArr(i)
           Else
               DataBufferArray(i + pointer) = 32 'Space key
           End If
       Next i
       pointer = pointer + i
       bArr1 = StrConv(ScsiProductTxt, vbFromUnicode)
       iStringLen1 = LenB(StrConv(ScsiProductTxt, vbFromUnicode))
       For i = 0 To 15
           If (i < iStringLen1) Then
               DataBufferArray(i + pointer) = bArr1(i)
           Else
               DataBufferArray(i + pointer) = 32 'Space key
           End If
       Next i
       pointer = pointer + i
        For i = 0 To 4
             If (i < Len(ScsiRevTxt)) Then
                DataBufferArray(i + pointer) = Asc(Mid(ScsiRevTxt, i + 1, 1))
             Else
                 DataBufferArray(i + pointer) = 0
             End If
        Next i
        pointer = pointer + i
       For i = 1 To pointer
           checksum = (checksum + CInt(DataBufferArray(i))) And &HFF
       Next i
    
        'send to scsi
        DataBufferArray(pointer - 1) = CByte(checksum)
        DataBufferArray(pointer) = &HAA
        DataBufferArray(pointer + 1) = &H55
    
        If WriteToDisk = True Then
            'vid, pid,..... inquiry string
            result = fnScsi2usb_VendorW(Disk(DiskIDnumber).Handle, Write_VID, 512, DataBufferArray(1))
            If result = 0 Then
                SetConfiguration = 0
                Disk(DiskIDnumber).ErrStatus = &HF6
            Else
                SetConfiguration = 1
            End If
        End If
    Else
        SetConfiguration = 2
        Disk(DiskIDnumber).ErrStatus = &HF7
    End If
End Function
Public Function CheckConfiguration(DiskIDnumber As Integer)
Dim result As Integer
Dim i As Integer
    On Error Resume Next
    
    result = fnScsi2usb_VendorR(Disk(DiskIDnumber).Handle, Read_VID, 0, 512)
    If result = 0 Then
        Disk(DiskIDnumber).ErrStatus = Read_VID
        CheckConfiguration = 0
    Else
        Receive_DataBuffer
        
        SetConfiguration 1, False   'load configuration data to DataBufferArray
        For i = 1 To 512
            If (DataBufferArray(i) <> ReceiveDataBuffer(i)) Then
                Disk(DiskIDnumber).ErrStatus = &HF4
                CheckConfiguration = 1
                Exit Function
            End If
        Next i
        
        CheckConfiguration = 2
    End If
End Function
 

Public Function CheckWP(DiskIDnumber As Integer)
Dim result As Integer
'Dim i As Integer
    On Error Resume Next
    
    'fnScsi2usb_TestUnitReady(ByVal filehandle As Long)
    'MUST issue 2 times of command TestUnitReady
    result = fnScsi2usb_TestUnitReady(Disk(DiskIDnumber).Handle)
    result = fnScsi2usb_TestUnitReady(Disk(DiskIDnumber).Handle)

    'fnScsi2usb_ModeSense(ByVal filehandle As Long)
    result = fnScsi2usb_ModeSense(Disk(DiskIDnumber).Handle)
    
    Receive_DataBuffer
    
    If ReceiveDataBuffer(83) = 128 Then
        CheckWP = 1
    Else
        CheckWP = 0
        Disk(DiskIDnumber).ErrStatus = &HF5
    End If
    

    'fnScsi2usb_Read(ByVal filehandle As Long, ByVal TL As Long, ByVal LBA As Long)
    'result = fnScsi2usb_Read(Disk(DiskIDnumber).Handle, 1, 19)
    'If result = 0 Then
    '    Disk(DiskIDnumber).ErrStatus = &H28
    '    CheckWP = 0
    'Else
    '    Receive_DataBuffer
    '    'fnScsi2usb_Write(ByVal filehandle As Long, ByVal TL As Long, ByVal LBA As Long, DataBuffer As Any)
    '    result = fnScsi2usb_Write(Disk(DiskIDnumber).Handle, 1, 19, ReceiveDataBuffer(1))
    '    If result = 0 Then
    '        Disk(DiskIDnumber).ErrStatus = &H2A
    '        CheckWP = 1
    '    Else
    '        CheckWP = 0
    '    End If
    'End If
End Function

Public Sub LowLevelFormat(DiskIDnumber As Integer, op As Integer)
    Dim result As Integer, rowtmp As Integer
    Dim i As Integer, j As Integer, tmpk As Integer, k As Integer
    Dim doCount As Long
    Dim tmp As Integer
    Dim SScanOneStep As Long
    Dim SScanValue  As Long
    Dim SScanMax As Long
    Dim Increment As Integer, TotalZone As Integer
    On Error Resume Next
    
    If (Disk(DiskIDnumber).Handle = 0) Then
       Disk(DiskIDnumber).ErrStatus = &HFE
       Exit Sub
    End If
    result = fnScsi2usb_VendorR(Disk(DiskIDnumber).Handle, GetChipNumber, 0, 512)
    If result = 0 Then
        Disk(DiskIDnumber).ErrStatus = GetChipNumber
        Exit Sub
    End If
    Receive_DataBuffer
    Call InterpreteDeviceCodeOP(DiskIDnumber)
    TotalZone = 0
    For i = 1 To Disk(DiskIDnumber).ChipNumber
        TotalZone = TotalZone + DCode(i).ZoneNumber
    Next i
    Increment = TestFrm.ProgressBar(DiskIDnumber - 1).Max / TotalZone
    
        Disk(DiskIDnumber).BadBlockNumber = 0
        For i = 1 To Disk(DiskIDnumber).ChipNumber
            For j = 0 To DCode(i).ZoneNumber - 1
                doCount = doCount + 1
                result = fnScsi2usb_Vendor(Disk(DiskIDnumber).Handle, op, i - 1, j, 512)
                
                Receive_DataBuffer
                For k = 2 To 511
                    If (ReceiveDataBuffer(k) <> 170) Or (ReceiveDataBuffer(k + 1) <> 85) Then
                        Disk(DiskIDnumber).BadBlockNumber = Disk(DiskIDnumber).BadBlockNumber + 1
                    Else
                        Exit For
                    End If
                Next k
                
                TestFrm.ProgressBar(DiskIDnumber - 1).Value = TestFrm.ProgressBar(DiskIDnumber - 1).Value + Increment
                If result = 0 Then
                    Disk(DiskIDnumber).ErrStatus = op
                    Exit Sub
                End If
             Next j
         Next i
    For i = 1 To 512
        SendDataBuffer(i) = 0
    Next i
    SendDataBuffer(1) = &HFF
    SendDataBuffer(2) = &HFF
    SendDataBuffer(3) = &HFF
    SendDataBuffer(4) = &HFF
 
    result = fnScsi2usb_VendorW(Disk(DiskIDnumber).Handle, Get_FAT, 512, SendDataBuffer(1)) 'command
    result = fnScsi2usb_VendorR(Disk(DiskIDnumber).Handle, GetConfiguration, 0&, 512)
    If result = 0 Then
        Disk(DiskIDnumber).ErrStatus = GetConfiguration
       Exit Sub
    End If
     
    For i = 1 To 512
        SendDataBuffer(i) = 0
    Next i
    result = fnScsi2usb_TestUnitReady(Disk(DiskIDnumber).Handle)
    Disk(DiskIDnumber).ErrStatus = 0
End Sub

Public Sub InterpreteDeviceCodeOP(DiskIDnumber As Integer)
Dim i As Integer, j As Integer
Dim MaxValue As Integer
Dim tmp As Byte
Dim TotalBlock As Integer
Dim M As Integer
Dim ii As Long
    On Error Resume Next
    
'initialize
 tmp = 0
 i = 1
 MaxValue = 0
 TotalBlock = 0
   Do Until (tmp = 170) '0xaa
       DCode(i).DeviceCode = ReceiveDataBuffer(i)
       tmp = DCode(i).DeviceCode
       i = i + 1
   Loop
   Disk(DiskIDnumber).ChipNumber = i - 2
   For j = 1 To Disk(DiskIDnumber).ChipNumber
       Select Case DCode(j).DeviceCode
              Case 230 '0xE6
                   DCode(j).ZoneNumber = 4
                   DCode(j).BlockNumber = 256
                   DCode(j).SectorNumber = 16
                   TotalBlock = TotalBlock + 4 * 256
                   M = M + 8
                   
              Case 115 '0x73
                   DCode(j).ZoneNumber = 4
                   DCode(j).BlockNumber = 256
                   DCode(j).SectorNumber = 32
                   TotalBlock = TotalBlock + 4 * 256
                   M = M + 16
                   
              Case 117 '0x75
                   DCode(j).ZoneNumber = 8
                   DCode(j).BlockNumber = 256
                   DCode(j).SectorNumber = 32
                   TotalBlock = TotalBlock + 8 * 256
                   M = M + 32
                   
              Case 118 '0x76
                   DCode(j).ZoneNumber = 16
                   DCode(j).BlockNumber = 256
                   DCode(j).SectorNumber = 32
                   TotalBlock = TotalBlock + 16 * 256
                   M = M + 64
                   
              Case 121 '0x79
                   DCode(j).ZoneNumber = 32
                   DCode(j).BlockNumber = 256
                   DCode(j).SectorNumber = 32
                   TotalBlock = TotalBlock + 32 * 256
                   M = M + 128
       End Select
       MaxValue = MaxValue + DCode(j).ZoneNumber
   Next j
   Disk(DiskIDnumber).TotalZoneNumber = MaxValue
End Sub

 Public Sub ReadCapacity(DiskIDNUM As Integer)
    Dim result As Integer
    Dim X As String
    On Error Resume Next
    result = fnScsi2usb_ReadCapacity(Disk(DiskIDNUM).Handle)
    If result <> 0 Then
    '1 zero base, 63 mean hidden sectors
       PrintOut result
       Disk(DiskIDNUM).PhysicalDiskCapacity = (Val(tmpResult) - 1) * 512 / 1024 / 1024
       Disk(DiskIDNUM).PhysicalDiskSectors = Val(tmpResult) - 1
    Else
       Disk(DiskIDNUM).ErrStatus = &H9A
    End If
    tmpResult = ""
End Sub
 

 
Public Function HexFormat(Value As Long, num As Integer, ctype As Boolean)
    Dim s As String
    Dim d, i As Integer
    On Error Resume Next
    
    d = Len(Hex(Value))
    If (d < num) Then
        s = ""
        For i = 1 To num - d
            s = s + "0"
        Next i
    ElseIf (d = num) Then
        s = ""
    Else
        s = ""
    End If
    If ctype = True Then
       HexFormat = "0x" & s & Hex(Value)
    Else
       HexFormat = s & Hex(Value)
    End If
End Function



 
 
 
