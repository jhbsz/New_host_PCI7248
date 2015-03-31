Attribute VB_Name = "UtilityMDL"
Option Explicit

Public PortFail As String
Public Sub SetTimer_1ms()
Dim ERR As Integer
ERR = CTR_Setup(card, 1, RATE_GENERATOR, 200, BINTimer)
ERR = CTR_Setup(card, 2, RATE_GENERATOR, 10, BINTimer)

End Sub
Public Sub Timer_1ms(ms As Integer)
Dim result
Dim old_value1
Dim old_value2
Dim i As Integer
Dim ii As Integer
Dim T1 As Long
Dim T2 As Long

 
result = CTR_Read(0, 2, old_value1)

T1 = Timer
   For i = 1 To ms
   
   
            Do
            DoEvents
            result = CTR_Read(0, 2, old_value2)
            T2 = Timer
                If T2 - T1 > 5 Then
                    MsgBox ("PCI7248_Counter Error !!")
                    End
                End If
            Loop Until old_value1 <> old_value2
    
            Do
            DoEvents
            result = CTR_Read(0, 2, old_value2)
                If T2 - T1 > 5 Then
                    MsgBox ("PCI7248_Counter Error !!")
                    End
                End If
            Loop Until old_value1 = old_value2
    
    
    
    Next
     
End Sub

Public Sub Timer_1ms_OpenShort(ms As Integer, pin As Byte)
Dim result
Dim old_value1
Dim old_value2
Dim i As Integer
Dim ii As Integer


 
result = CTR_Read(0, 2, old_value1)

If OSValue(pin) <> OSStandard(pin) Then
    TestResult = "Fail"
    Exit Sub
End If

   For i = 1 To ms
            Do
            DoEvents
            result = CTR_Read(0, 2, old_value2)
            Loop Until old_value1 <> old_value2
    
            Do
            DoEvents
            result = CTR_Read(0, 2, old_value2)
            Loop Until old_value1 = old_value2
    
    
    Next
     

 
End Sub
Sub LabelMenuSu(SlotNo As Byte, TestResult As Byte, PreSlotStatus As Byte)

If PreSlotStatus <> 1 Then
Exit Sub
End If

Select Case SlotNo

Case 0
        
        Select Case TestResult 'rv0
        
        Case 0 'Detect FAIL
                 Tester.Label9 = "Unknow Device "
                 Tester.Label3.BackColor = RGB(255, 0, 0) 'red
                 Tester.Label33.BackColor = RGB(255, 0, 0) 'red
        
        Case 1 'write ok
                 Tester.Label9 = "Write Pass "
                 Tester.Label3.BackColor = RGB(0, 255, 0) 'green
                 Tester.Label33.BackColor = RGB(0, 255, 0) 'green
                 
        Case 2 'speed fail
                 Tester.Label9 = "USB SPEED ERROR "
                 Tester.Label3.BackColor = RGB(255, 0, 0) 'red
                 Tester.Label33.BackColor = RGB(255, 0, 0) 'red
        End Select
    
Case 1

        Select Case TestResult 'rv1
                        
        Case 1 'Erease pass
                 Tester.Label9 = "Erease Pass "
                 Tester.Label4.BackColor = RGB(0, 255, 0) 'green
                 Tester.Label33.BackColor = RGB(0, 255, 0) 'green
                        
        Case 2 'Erease fail
                 Tester.Label9 = "Erease Error "
                 Tester.Label4.BackColor = RGB(255, 0, 0) 'red
                 Tester.Label33.BackColor = RGB(255, 0, 0) 'red
        
        End Select

Case 2
       
        Select Case TestResult 'rv2
                        
        Case 1 'Write pass 1
                 Tester.Label9 = "Write Pass "
                 Tester.Label5.BackColor = RGB(0, 255, 0) 'green
                 Tester.Label33.BackColor = RGB(0, 255, 0) 'green
        Case 2 'Write fail 1
                 Tester.Label9 = "Write Error "
                 Tester.Label5.BackColor = RGB(255, 0, 0) 'red
                 Tester.Label33.BackColor = RGB(255, 0, 0) 'red
        End Select

Case 3
               
        Select Case TestResult 'rv3
                        
        Case 1 'Write pass 2
                 Tester.Label9 = "Write Pass "
                 Tester.Label6.BackColor = RGB(0, 255, 0) 'green
                 Tester.Label33.BackColor = RGB(0, 255, 0) 'green
        Case 2 'Write fail 2
                 Tester.Label9 = "Write Error "
                 Tester.Label6.BackColor = RGB(255, 0, 0) 'red
                 Tester.Label33.BackColor = RGB(255, 0, 0) 'red
        End Select
        
Case 4

        Select Case TestResult 'rv4
        
        Case 1 'read pass
                 Tester.Label9 = "Read Pass "
                 Tester.Label7.BackColor = RGB(0, 255, 0) 'green
                 Tester.Label33.BackColor = RGB(0, 255, 0) 'green
        Case 3 'read fail
                 Tester.Label9 = "Read Error "
                 Tester.Label7.BackColor = RGB(255, 0, 0) 'red
                 Tester.Label33.BackColor = RGB(255, 0, 0) 'red
        End Select

End Select

End Sub
Public Static Sub MsecDelay(Msec As Single)

Dim start As Single
Dim Pause As Single
Dim NextDayFlag As Boolean
Dim WaitTime As Single


'MAX ='86400
If (Timer + Msec) > 86400 Then
    NextDayFlag = True
    WaitTime = (Timer + Msec) - 86400
    start = 0
Else
    NextDayFlag = False
    start = Timer
    WaitTime = Msec
End If


If NextDayFlag Then
    Do
        DoEvents
    Loop Until (Timer <= 2)
End If


Do

    DoEvents
    Pause = Timer - start
    
Loop Until (Pause >= WaitTime)

End Sub

 Public Function GetDeviceNameMulti6710(Vid_PID As String) As String
Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long
Dim i As Integer

MemberIndex = HubPort
'  MemberIndex = 2
'LastDevice = False
'MyDeviceDetected = False
 
  
  ' With HidGuid ' CD-ROM Gerateklassen GUID
  '    .Data1 = &H53F56307
  '    .Data2 = &HB6BF
  '    .Data3 = &H11D0
  '    .Data4(0) = &H94
  '    .Data4(1) = &HF2
  '    .Data4(2) = &H0
  '    .Data4(3) = &HA0
  '    .Data4(4) = &HC9
  '    .Data4(5) = &H1E
  '    .Data4(6) = &HFB
  '    .Data4(7) = &H8B
  ' End With
   
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

'53f56307,0xb6bf,0x11d0,0x94,0xf2,0x00,0xa0,0xc9,0x1e,0xfb,0x8b);

DevicePathName = ""
DeviceInfoSet = SetupDiGetClassDevs _
    (HidGuid, _
    vbNullString, _
    0, _
    (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
    
 Do

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
        
       ' Print #1, DevicePathName
        
 '        HIDHandle = CreateFile _
 '           (DevicePathName, _
 '           GENERIC_READ Or GENERIC_WRITE, _
 '           (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
 '           Security, _
 '           OPEN_EXISTING, _
 '           0&, _
 '           0)
            
        'Set the Size property to the number of bytes in the structure.
 '       DeviceAttributes.Size = LenB(DeviceAttributes)
 '       Result = HidD_GetAttributes _
 '           (HIDHandle, _
 '           DeviceAttributes)
         
'        PID_LIST.AddItem Str(MemberIndex) & "--->" & Hex$(DeviceAttributes.ProductID)
 '       VID_LIST.AddItem Str(MemberIndex) & "--->" & Hex$(DeviceAttributes.VendorID)
        
          'Find out if the device matches the one we're looking for.
 '       If (DeviceAttributes.VendorID = MyVendorID) And _
 '           (DeviceAttributes.ProductID = MyProductID) Then
                'It's the desired device.
 '               Form1.Caption = "FIND DEVICE"
 '               MyDeviceDetected = True
 '       Else
 '               MyDeviceDetected = False
                'If it's not the one we want, close its handle.
  '              Result = CloseHandle _
  '                  (HIDHandle)
  '      End If
        
 
End If
 
  MemberIndex = MemberIndex + 1


 Debug.Print MemberIndex, DevicePathName
 
'If InStr(1, DevicePathName, Vid_PID) Or InStr(1, DevicePathName, "9101") Then
If InStr(1, DevicePathName, "6710") Then       'old
 'If InStr(1, DevicePathName, "058f0o1111b") And InStr(1, DevicePathName, "reader") Then  'Alcor 6331
GetDeviceNameMulti6710 = DevicePathName
'Tester.Print "VID="; Mid(DevicePathName, InStr(1, DevicePathName, "vid"), 8); " ; ";
'Tester.Print "PID="; Mid(DevicePathName, InStr(1, DevicePathName, "pid"), 8)

Else
GetDeviceNameMulti6710 = ""
 
End If

Loop While GetDeviceNameMulti6710 = "" And MemberIndex < 10
'Tester.Print "VID="; DevicePathName
End Function

 Public Function GetDeviceNameMulti9369(Vid_PID As String) As String
Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long
Dim i As Integer

 'MemberIndex = HubPort
'  MemberIndex = 2
'LastDevice = False
'MyDeviceDetected = False
 
  
   With HidGuid ' CD-ROM Gerateklassen GUID
      .Data1 = &H53F56307
      .Data2 = &HB6BF
      .Data3 = &H11D0
      .Data4(0) = &H94
      .Data4(1) = &HF2
      .Data4(2) = &H0
      .Data4(3) = &HA0
      .Data4(4) = &HC9
      .Data4(5) = &H1E
      .Data4(6) = &HFB
      .Data4(7) = &H8B
   End With

'53f56307,0xb6bf,0x11d0,0x94,0xf2,0x00,0xa0,0xc9,0x1e,0xfb,0x8b);

DevicePathName = ""
DeviceInfoSet = SetupDiGetClassDevs _
    (HidGuid, _
    vbNullString, _
    0, _
    (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
    
 Do

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
        
       ' Print #1, DevicePathName
        
 '        HIDHandle = CreateFile _
 '           (DevicePathName, _
 '           GENERIC_READ Or GENERIC_WRITE, _
 '           (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
 '           Security, _
 '           OPEN_EXISTING, _
 '           0&, _
 '           0)
            
        'Set the Size property to the number of bytes in the structure.
 '       DeviceAttributes.Size = LenB(DeviceAttributes)
 '       Result = HidD_GetAttributes _
 '           (HIDHandle, _
 '           DeviceAttributes)
         
'        PID_LIST.AddItem Str(MemberIndex) & "--->" & Hex$(DeviceAttributes.ProductID)
 '       VID_LIST.AddItem Str(MemberIndex) & "--->" & Hex$(DeviceAttributes.VendorID)
        
          'Find out if the device matches the one we're looking for.
 '       If (DeviceAttributes.VendorID = MyVendorID) And _
 '           (DeviceAttributes.ProductID = MyProductID) Then
                'It's the desired device.
 '               Form1.Caption = "FIND DEVICE"
 '               MyDeviceDetected = True
 '       Else
 '               MyDeviceDetected = False
                'If it's not the one we want, close its handle.
  '              Result = CloseHandle _
  '                  (HIDHandle)
  '      End If
        
 
End If
 
  MemberIndex = MemberIndex + 1


 Debug.Print MemberIndex, DevicePathName
 
'If InStr(1, DevicePathName, Vid_PID) Or InStr(1, DevicePathName, "9101") Then
If InStr(1, DevicePathName, "prod_flash_reader&rev_1.00#2004168&") Then       'old
 'If InStr(1, DevicePathName, "058f0o1111b") And InStr(1, DevicePathName, "reader") Then  'Alcor 6331
GetDeviceNameMulti9369 = DevicePathName
'Tester.Print "VID="; Mid(DevicePathName, InStr(1, DevicePathName, "vid"), 8); " ; ";
'Tester.Print "PID="; Mid(DevicePathName, InStr(1, DevicePathName, "pid"), 8)

Else
GetDeviceNameMulti9369 = ""
 
End If

Loop While GetDeviceNameMulti9369 = "" And MemberIndex < 10
'Tester.Print "VID="; DevicePathName
End Function

Public Sub NewLabelMenu(Flow As Byte, Item As String, TestResult As Byte, PreSlotSattus As Byte)

If PreSlotSattus <> 1 Then
    Exit Sub
End If

Select Case Flow
    Case 0      'Enum Device

        If TestResult = 1 Then
            Tester.Label33.BackColor = vbGreen
        Else
            Tester.Label33.BackColor = vbRed
            Tester.Label9 = "Unknow Device Fail"
        End If
        
    Case 1
        If UsbSpeedTestResult = 2 Then
            Tester.Label9 = "USB SPEED ERROR"
            Tester.Label3.BackColor = vbYellow
            Exit Sub
        End If
        
        If UsbSpeedTestResult = GPO_FAIL Then
            Tester.Label9 = "LED/GPO Fail"
            Tester.Label3.BackColor = vbYellow
            Exit Sub
        End If
        
        If TestResult = 1 Then
            Tester.Label3.BackColor = vbGreen
        Else
            Tester.Label9 = Item & " Fail"
            Tester.Label3.BackColor = vbRed
        End If
        
    Case 2
        If TestResult = 1 Then
            Tester.Label4.BackColor = vbGreen
        Else
            Tester.Label9 = Item & " Fail"
            Tester.Label4.BackColor = vbRed
        End If
    
    Case 3
        If TestResult = 1 Then
            Tester.Label5.BackColor = vbGreen
        Else
            Tester.Label9 = Item & " Fail"
            Tester.Label5.BackColor = vbRed
        End If
        
    Case 4
        If TestResult = 1 Then
            Tester.Label6.BackColor = vbGreen
        Else
            Tester.Label9 = Item & " Fail"
            Tester.Label6.BackColor = vbRed
        End If
    
    Case 5
        If TestResult = 1 Then
            Tester.Label7.BackColor = vbGreen
        Else
            Tester.Label9 = Item & " Fail"
            Tester.Label7.BackColor = vbRed
        End If

End Select

End Sub

Sub LabelMenu(SlotNo As Byte, TestResult As Byte, PreSlotStatus As Byte)


If (SlotNo = 0) And (PreSlotStatus = 0) Then  'OS Label
    
    If TestResult = 1 Then
        Tester.Label9 = "Open/Short PASS "
        Tester.Label33.BackColor = RGB(0, 255, 0)   'OS PASS
        Exit Sub
    Else
        Tester.Label9 = "Open/Short FAIL"
        Tester.Label33.BackColor = RGB(255, 0, 0)
        Exit Sub
    End If
    
ElseIf PreSlotStatus <> 1 Then
    Exit Sub
Else
    a = a
End If


Select Case SlotNo

Case 0, 10
        
        
        Select Case TestResult
        
        Case 0
                 Tester.Label9 = "unknow device "
                 Tester.Label33.BackColor = RGB(255, 0, 0)
        
        Case 1
                 Tester.Label3.BackColor = RGB(0, 255, 0)
                 Tester.Label33.BackColor = RGB(0, 255, 0)
                 
        Case 2, 3
                  If UsbSpeedTestResult = 0 Then
                  
                       If SlotNo = 0 Then
                        
                            Tester.Label9 = "SD Fail "
                            Tester.Label3.BackColor = RGB(255, 0, 0)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                        Else
                            Tester.Label9 = "MiniSD Fail "
                            Tester.Label3.BackColor = RGB(255, 255, 0)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                        End If
                            
                        
                  
                  ElseIf UsbSpeedTestResult = 2 Then
        
                        Tester.Label9 = "USB SPEED ERROR "
                        Tester.Label3.BackColor = RGB(255, 255, 0)
                        Tester.Label33.BackColor = RGB(0, 255, 0)
               
                    
                  Else  'GPO_FAIL
                    
                        Tester.Label9 = "GPO or CARD Detect FAIL or SD , MiniSD fail"
                        Tester.Label3.BackColor = RGB(0, 0, 255)
                        Tester.Label33.BackColor = RGB(0, 255, 0)
                   
                  End If
                  
                'For AU6980
                  
                If FlashCapacityError = 2 Then
        
                        Tester.Label9 = "Flash Capacity Error "
                        Tester.Label3.BackColor = RGB(255, 255, 0)
                        Tester.Label33.BackColor = RGB(0, 255, 0)
                        
                   End If
                   
                   
                 'For AU6610
                  If ChipName = "AU6610" Then  '!! do not change program sequace
                  
                  
                        If TestResult = 2 Then
                            Tester.Label9 = "USB SPEED ERROR "
                            Tester.Label3.BackColor = RGB(0, 0, 255)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                       ElseIf TestResult = 3 Then
                           Tester.Label9 = "I2C Bus Error "
                           Tester.Label3.BackColor = RGB(0, 0, 255)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                       End If
                        
                   
                  End If
                  
                  
                    If InStr(ChipName, "AU9520") <> 0 Then '!! do not change program sequace
                  
                  
                        If TestResult = 2 Then
                            Tester.Label9 = "Normal mode: find card fail"
                            Tester.Label3.BackColor = RGB(0, 0, 255)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                       ElseIf TestResult = 3 Then
                           Tester.Label9 = "Normal mode:smart card R/W fail "
                           Tester.Label3.BackColor = RGB(0, 0, 255)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                       End If
                        
                   
                  End If
                  
                    If InStr(ChipName, "AU9520") <> 0 And Len(ChipName) > 10 Then  '!! do not change program sequace
                  
                  
                        If TestResult = 2 Then
                            Tester.Label9 = "1st mode: find card fail"
                            Tester.Label3.BackColor = RGB(0, 0, 255)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                       ElseIf TestResult = 3 Then
                           Tester.Label9 = "1st mode  mode:smart card R/W fail "
                           Tester.Label3.BackColor = RGB(0, 0, 255)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                       End If
                        
                   
                  End If
                   
                  
                    
        End Select
    
Case 1


        Select Case TestResult
        
        Case 0
                Tester.Label9 = "unknow device "
                Tester.Label33.BackColor = RGB(255, 0, 0)
        
        Case 1
                 Tester.Label4.BackColor = RGB(0, 255, 0)
                Tester.Label33.BackColor = RGB(0, 255, 0)
                 
        Case 2, 3
                   If ChipName = "AU6254" Then
                      Tester.Label9 = "USB speed ERROR or  port1 unknow device"
                      Tester.Label4.BackColor = RGB(255, 0, 0)
                      Tester.Label33.BackColor = RGB(0, 255, 0)
                      Exit Sub
                  End If
             
                 
                  If UsbSpeedTestResult = 0 Then
                       Tester.Label9 = "CF Fail "
                        Tester.Label4.BackColor = RGB(255, 0, 0)
                        Tester.Label33.BackColor = RGB(0, 255, 0)
                  
                  ElseIf UsbSpeedTestResult = 2 Then
        
                        Tester.Label9 = "USB SPEED ERROR "
                        Tester.Label4.BackColor = RGB(255, 255, 0)
                       Tester.Label33.BackColor = RGB(0, 255, 0)
               
                    
                  Else  'GPO_FAIL
                    
                       Tester.Label9 = "GPO or CARD Detect FAIL or CF fail"
                       Tester.Label4.BackColor = RGB(0, 0, 255)
                       Tester.Label33.BackColor = RGB(0, 255, 0)
                   
                  End If
                  
                  
                  
                  If ChipName = "AU6610" Then  '!! do not change program sequace
                  
                        Tester.Label9 = "IsoChrous bus error "
                        Tester.Label4.BackColor = RGB(255, 0, 0)
                        Tester.Label33.BackColor = RGB(0, 255, 0)
                  End If
                  
                  
                 
                   
                  
                    
        End Select
          
      
    
    
Case 2, 21
                   '======================================
                   
         Select Case TestResult
        
        Case 0
                 Tester.Label9 = "unknow device "
                 Tester.Label33.BackColor = RGB(255, 0, 0)
        
        Case 1
                 Tester.Label5.BackColor = RGB(0, 255, 0)
                 Tester.Label33.BackColor = RGB(0, 255, 0)
                 
        Case 2, 3
                    If ChipName = "AU6254" Then
                            Tester.Label9 = PortFail
                            Tester.Label5.BackColor = RGB(255, 0, 0)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                             Exit Sub
                     End If
        
        
                  If UsbSpeedTestResult = 0 Then
                  
                  
                  
                  
                       If SlotNo = 2 Then
                            Tester.Label9 = "XD Fail "
                            Tester.Label5.BackColor = RGB(255, 0, 0)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                       Else
                            Tester.Label9 = "SMC Fail "
                            Tester.Label5.BackColor = RGB(255, 0, 255)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                       End If
                        
                  
                  ElseIf UsbSpeedTestResult = 2 Then
        
                        Tester.Label9 = "USB SPEED ERROR "
                        Tester.Label5.BackColor = RGB(255, 255, 0)
                        Tester.Label33.BackColor = RGB(0, 255, 0)
               
                    
                  Else  'GPO_FAIL
                    
                        Tester.Label9 = "GPO or CARD Detect FAIL or XD , SMC fail"
                        Tester.Label5.BackColor = RGB(0, 0, 255)
                        Tester.Label33.BackColor = RGB(0, 255, 0)
                   
                  End If
                  
                  
                     If InStr(ChipName, "AU9520") <> 0 Then    '!! do not change program sequace
                  
                  
                        If TestResult = 2 Then
                            Tester.Label9 = " find card fail"
                            Tester.Label5.BackColor = RGB(0, 0, 255)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                       ElseIf TestResult = 3 Then
                           Tester.Label9 = "smart card R/W fail "
                           Tester.Label5.BackColor = RGB(0, 0, 255)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                       End If
                        
                   
                  End If
                  
                  
                   If InStr(ChipName, "AU9520FLF2") <> 0 And Len(ChipName) > 10 Then '!! do not change program sequace
                  
                  
                        If TestResult = 2 Then
                            Tester.Label9 = "2nd mode: find card fail"
                            Tester.Label3.BackColor = RGB(0, 0, 255)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                       ElseIf TestResult = 3 Then
                           Tester.Label9 = "2nd mode  mode:smart card R/W fail "
                           Tester.Label3.BackColor = RGB(0, 0, 255)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                       End If
                        
                   
                  End If
                    
        End Select
                   
       

Case 3, 31, 32



           Select Case TestResult
        
        Case 0
                 Tester.Label9 = "unknow device "
                 Tester.Label33.BackColor = RGB(255, 0, 0)
        
        Case 1
                 Tester.Label6.BackColor = RGB(0, 255, 0)
                 Tester.Label33.BackColor = RGB(0, 255, 0)
                 
        Case 2, 3
                     
                      If ChipName = "AU6254" Then
                            Tester.Label9 = "SD card R/W fail ot Keyboard fail"
                            Tester.Label6.BackColor = RGB(255, 0, 0)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                            Exit Sub
                      End If
                  
        
                  If UsbSpeedTestResult = 0 Then
                  
                       If SlotNo = 3 Then
                            Tester.Label9 = "MS Fail "
                            Tester.Label6.BackColor = RGB(255, 0, 0)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                       Else
                            Tester.Label9 = "MSPro Fail "
                            Tester.Label6.BackColor = RGB(255, 0, 255)
                            Tester.Label33.BackColor = RGB(0, 255, 0)
                       End If
                        
                  
                  ElseIf UsbSpeedTestResult = 2 Then
        
                       Tester.Label9 = "USB SPEED ERROR "
                        Tester.Label6.BackColor = RGB(255, 255, 0)
                        Tester.Label33.BackColor = RGB(0, 255, 0)
               
                    
                  Else  'GPO_FAIL
                    
                       Tester.Label9 = "GPO or CARD Detect FAIL or MS, MSpro fail"
                       Tester.Label6.BackColor = RGB(0, 0, 255)
                       Tester.Label33.BackColor = RGB(0, 255, 0)
                   
                  End If
                  
                  
                  If GPOFail = 2 Then
                  
                        Tester.Label9 = "GPO fail"
                        Tester.Label6.BackColor = RGB(0, 0, 255)
                        Tester.Label33.BackColor = RGB(0, 255, 0)
                   
                  End If
                  
                  
                  
                  
                  
                  
                    
        End Select
       
        
Case 4, 41
        
        If TestResult <> 0 Then
        
                If TestResult <> 1 Then
                
                    If TestResult <> 11 Then
                    'If Dir("f:stepa01.gif") <> "stepa01.gif" Then
                        Tester.Print "teststep1-Fail"
                        Tester.Label9 = "Mini SD Fail "
                        Tester.Label3.BackColor = RGB(0, 0, 255)
                    Else
                        Tester.Print "USB SPEED ERROR"
                        Tester.Label9 = "USB SPEED ERROR"
                        Tester.Label3.BackColor = RGB(255, 255, 0)
                    End If
                    
                    
                    
                Else
                    Tester.Label7.BackColor = RGB(0, 255, 0)
                End If
                Tester.Label33.BackColor = RGB(0, 255, 0)
        Else
              
              Tester.Label9 = "unknow device "
              Tester.Label33.BackColor = RGB(255, 0, 0)
        End If

Case 5     '
       
        If TestResult <> 1 Then
        Tester.Cls
            'If Dir("f:stepa01.gif") <> "stepa01.gif" Then
            Tester.Print "teststep1-Fail"
            
             
            Tester.Label9 = "SmartCard Fail "
            Tester.Label3.BackColor = RGB(255, 0, 0)
        Else
           Tester.Cls
            Tester.Print "teststep1-pass"
            
            Tester.Label3.BackColor = RGB(0, 255, 0)
        End If
        
Case 51     '
       
        If TestResult <> 1 Then
      '  Tester.Cls
            'If Dir("f:stepa01.gif") <> "stepa01.gif" Then
           ' Tester.Print "teststep1-Fail"
            
             
           ' Tester.Label9 = "SmartCard Fail "
            Tester.Label7.BackColor = RGB(255, 0, 0)
        Else
           'Tester.Cls
            'Tester.Print "teststep1-pass"
            
            Tester.Label7.BackColor = RGB(0, 255, 0)
        End If
        
        
Case 6
       If TestResult <> 1 Then
       Tester.Cls
            Tester.Print "teststep2-Fail"
            
            Tester.Label4.BackColor = RGB(255, 0, 0)
           Tester.Label9 = "CF Fail "
        Else
            Tester.Print "teststep2-pass"
            Tester.Label4.BackColor = RGB(0, 255, 0)
        End If
End Select
    
End Sub
Public Function GetDeviceNameMulti6331(Vid_PID As String) As String
Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long
Dim i As Integer

 'MemberIndex = HubPort
' MemberIndex = 1
'LastDevice = False
'MyDeviceDetected = False
 
  
   With HidGuid ' CD-ROM Gerateklassen GUID
      .Data1 = &H53F56307
      .Data2 = &HB6BF
      .Data3 = &H11D0
      .Data4(0) = &H94
      .Data4(1) = &HF2
      .Data4(2) = &H0
      .Data4(3) = &HA0
      .Data4(4) = &HC9
      .Data4(5) = &H1E
      .Data4(6) = &HFB
      .Data4(7) = &H8B
   End With

'53f56307,0xb6bf,0x11d0,0x94,0xf2,0x00,0xa0,0xc9,0x1e,0xfb,0x8b);

DevicePathName = ""
DeviceInfoSet = SetupDiGetClassDevs _
    (HidGuid, _
    vbNullString, _
    0, _
    (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
    
 Do

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
        
       ' Print #1, DevicePathName
        
 '        HIDHandle = CreateFile _
 '           (DevicePathName, _
 '           GENERIC_READ Or GENERIC_WRITE, _
 '           (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
 '           Security, _
 '           OPEN_EXISTING, _
 '           0&, _
 '           0)
            
        'Set the Size property to the number of bytes in the structure.
 '       DeviceAttributes.Size = LenB(DeviceAttributes)
 '       Result = HidD_GetAttributes _
 '           (HIDHandle, _
 '           DeviceAttributes)
         
'        PID_LIST.AddItem Str(MemberIndex) & "--->" & Hex$(DeviceAttributes.ProductID)
 '       VID_LIST.AddItem Str(MemberIndex) & "--->" & Hex$(DeviceAttributes.VendorID)
        
          'Find out if the device matches the one we're looking for.
 '       If (DeviceAttributes.VendorID = MyVendorID) And _
 '           (DeviceAttributes.ProductID = MyProductID) Then
                'It's the desired device.
 '               Form1.Caption = "FIND DEVICE"
 '               MyDeviceDetected = True
 '       Else
 '               MyDeviceDetected = False
                'If it's not the one we want, close its handle.
  '              Result = CloseHandle _
  '                  (HIDHandle)
  '      End If
        
 
End If
 
  MemberIndex = MemberIndex + 1


 Debug.Print MemberIndex, DevicePathName
 
'If InStr(1, DevicePathName, Vid_PID) Or InStr(1, DevicePathName, "9101") Then
If InStr(1, DevicePathName, "prod_2.0_card_reader") Then       'old
 'If InStr(1, DevicePathName, "058f0o1111b") And InStr(1, DevicePathName, "reader") Then  'Alcor 6331
GetDeviceNameMulti6331 = DevicePathName
'Tester.Print "VID="; Mid(DevicePathName, InStr(1, DevicePathName, "vid"), 8); " ; ";
'Tester.Print "PID="; Mid(DevicePathName, InStr(1, DevicePathName, "pid"), 8)

Else
GetDeviceNameMulti6331 = ""
 
End If

Loop While GetDeviceNameMulti6331 = "" And MemberIndex < 10
'Tester.Print "VID="; DevicePathName
End Function
Public Function GetDeviceNameHub(Vid_PID As String) As String
Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long

 
'LastDevice = False
'MyDeviceDetected = False
 
  
  
    With HidGuid ' CD-ROM Gerateklassen GUID
       .Data1 = &HF18A0E88
       .Data2 = &HC30C
       .Data3 = &H11D0
       .Data4(0) = &H88
       .Data4(1) = &H15
       .Data4(2) = &H0
       .Data4(3) = &HA0
       .Data4(4) = &HC9
       .Data4(5) = &H6
       .Data4(6) = &HBE
       .Data4(7) = &HD8
    End With




DevicePathName = ""
DeviceInfoSet = SetupDiGetClassDevs _
    (HidGuid, _
    vbNullString, _
    0, _
    (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
    
Do
 
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
        
       ' Print #1, DevicePathName
        
 '        HIDHandle = CreateFile _
 '           (DevicePathName, _
 '           GENERIC_READ Or GENERIC_WRITE, _
 '           (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
 '           Security, _
 '           OPEN_EXISTING, _
 '           0&, _
 '           0)
            
        'Set the Size property to the number of bytes in the structure.
 '       DeviceAttributes.Size = LenB(DeviceAttributes)
 '       Result = HidD_GetAttributes _
 '           (HIDHandle, _
 '           DeviceAttributes)
         
'        PID_LIST.AddItem Str(MemberIndex) & "--->" & Hex$(DeviceAttributes.ProductID)
 '       VID_LIST.AddItem Str(MemberIndex) & "--->" & Hex$(DeviceAttributes.VendorID)
        
          'Find out if the device matches the one we're looking for.
 '       If (DeviceAttributes.VendorID = MyVendorID) And _
 '           (DeviceAttributes.ProductID = MyProductID) Then
                'It's the desired device.
 '               Form1.Caption = "FIND DEVICE"
 '               MyDeviceDetected = True
 '       Else
 '               MyDeviceDetected = False
                'If it's not the one we want, close its handle.
  '              Result = CloseHandle _
  '                  (HIDHandle)
  '      End If
        
 
End If
 MemberIndex = MemberIndex + 1
If InStr(1, DevicePathName, Vid_PID) Or InStr(1, DevicePathName, "9101") Then
 
GetDeviceNameHub = DevicePathName

Else
GetDeviceNameHub = ""
 
End If

Loop While GetDeviceNameHub = "" And MemberIndex < 10

Tester.Print "VID="; Mid(DevicePathName, InStr(1, DevicePathName, "vid"), 8); " ; ";
Tester.Print "PID="; Mid(DevicePathName, InStr(1, DevicePathName, "pid"), 8)


End Function

Public Function GetDeviceNameHub_NoReply(Vid_PID As String) As String
Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long

 
'LastDevice = False
'MyDeviceDetected = False
 
  
  
    With HidGuid ' CD-ROM Gerateklassen GUID
       .Data1 = &HF18A0E88
       .Data2 = &HC30C
       .Data3 = &H11D0
       .Data4(0) = &H88
       .Data4(1) = &H15
       .Data4(2) = &H0
       .Data4(3) = &HA0
       .Data4(4) = &HC9
       .Data4(5) = &H6
       .Data4(6) = &HBE
       .Data4(7) = &HD8
    End With




DevicePathName = ""
DeviceInfoSet = SetupDiGetClassDevs _
    (HidGuid, _
    vbNullString, _
    0, _
    (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
    
Do
 
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
        
       ' Print #1, DevicePathName
        
 '        HIDHandle = CreateFile _
 '           (DevicePathName, _
 '           GENERIC_READ Or GENERIC_WRITE, _
 '           (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
 '           Security, _
 '           OPEN_EXISTING, _
 '           0&, _
 '           0)
            
        'Set the Size property to the number of bytes in the structure.
 '       DeviceAttributes.Size = LenB(DeviceAttributes)
 '       Result = HidD_GetAttributes _
 '           (HIDHandle, _
 '           DeviceAttributes)
         
'        PID_LIST.AddItem Str(MemberIndex) & "--->" & Hex$(DeviceAttributes.ProductID)
 '       VID_LIST.AddItem Str(MemberIndex) & "--->" & Hex$(DeviceAttributes.VendorID)
        
          'Find out if the device matches the one we're looking for.
 '       If (DeviceAttributes.VendorID = MyVendorID) And _
 '           (DeviceAttributes.ProductID = MyProductID) Then
                'It's the desired device.
 '               Form1.Caption = "FIND DEVICE"
 '               MyDeviceDetected = True
 '       Else
 '               MyDeviceDetected = False
                'If it's not the one we want, close its handle.
  '              Result = CloseHandle _
  '                  (HIDHandle)
  '      End If
        
 
End If
 MemberIndex = MemberIndex + 1
If InStr(1, DevicePathName, Vid_PID) Or InStr(1, DevicePathName, "9101") Then
 
GetDeviceNameHub_NoReply = DevicePathName

Else
GetDeviceNameHub_NoReply = ""
 
End If

Loop While GetDeviceNameHub_NoReply = "" And MemberIndex < 10

'Tester.Print "VID="; Mid(DevicePathName, InStr(1, DevicePathName, "vid"), 8); " ; ";
'Tester.Print "PID="; Mid(DevicePathName, InStr(1, DevicePathName, "pid"), 8)


End Function

Public Function GetDeviceNameMulti(Vid_PID As String) As String
Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long
Dim i As Integer

 MemberIndex = HubPort
'LastDevice = False
'MyDeviceDetected = False
 
  
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
    
 ' For i = 0 To 3
 
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
        
       ' Print #1, DevicePathName
        
 '        HIDHandle = CreateFile _
 '           (DevicePathName, _
 '           GENERIC_READ Or GENERIC_WRITE, _
 '           (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
 '           Security, _
 '           OPEN_EXISTING, _
 '           0&, _
 '           0)
            
        'Set the Size property to the number of bytes in the structure.
 '       DeviceAttributes.Size = LenB(DeviceAttributes)
 '       Result = HidD_GetAttributes _
 '           (HIDHandle, _
 '           DeviceAttributes)
         
'        PID_LIST.AddItem Str(MemberIndex) & "--->" & Hex$(DeviceAttributes.ProductID)
 '       VID_LIST.AddItem Str(MemberIndex) & "--->" & Hex$(DeviceAttributes.VendorID)
        
          'Find out if the device matches the one we're looking for.
 '       If (DeviceAttributes.VendorID = MyVendorID) And _
 '           (DeviceAttributes.ProductID = MyProductID) Then
                'It's the desired device.
 '               Form1.Caption = "FIND DEVICE"
 '               MyDeviceDetected = True
 '       Else
 '               MyDeviceDetected = False
                'If it's not the one we want, close its handle.
  '              Result = CloseHandle _
  '                  (HIDHandle)
  '      End If
        
 
End If
 
'  MemberIndex = MemberIndex + 1


 'Debug.Print MemberIndex, DevicePathName
 
If InStr(1, DevicePathName, Vid_PID) Or InStr(1, DevicePathName, "9101") Then
 
GetDeviceNameMulti = DevicePathName
'Tester.Print "VID="; Mid(DevicePathName, InStr(1, DevicePathName, "vid"), 8); " ; ";
'Tester.Print "PID="; Mid(DevicePathName, InStr(1, DevicePathName, "pid"), 8)

Else
GetDeviceNameMulti = ""
 
End If

' Next i

End Function

Public Function GetDeviceName(Vid_PID As String) As String
Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long

' MemberIndex = HubPort
'LastDevice = False
'MyDeviceDetected = False
 
  
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
        
    For MemberIndex = 0 To 10
 
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

        If InStr(1, DevicePathName, Vid_PID) Then
            GetDeviceName = DevicePathName
            Tester.Print "VID="; Mid(DevicePathName, InStr(1, DevicePathName, "vid"), 8); " ; ";
            Tester.Print "PID="; Mid(DevicePathName, InStr(1, DevicePathName, "pid"), 8)
            Exit Function
        Else
            GetDeviceName = ""
            Tester.Print "VID= unknow"; " ; ";
            Tester.Print "PID= unknow"
        End If

    Next MemberIndex
    
End Function

Public Function GetDeviceName_NoReply(Vid_PID As String) As String
Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long

' MemberIndex = HubPort
'LastDevice = False
'MyDeviceDetected = False

With HidGuid ' CD-ROM Gerateklassen GUID
    If (Vid_PID <> "pid_6254") Then
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
    Else
        .Data1 = &HF18A0E88
        .Data2 = &HC30C
        .Data3 = &H11D0
        .Data4(0) = &H88
        .Data4(1) = &H15
        .Data4(2) = &H0
        .Data4(3) = &HA0
        .Data4(4) = &HC9
        .Data4(5) = &H6
        .Data4(6) = &HBE
        .Data4(7) = &HD8
    End If
End With

For MemberIndex = 0 To 10

    DevicePathName = ""
    DeviceInfoSet = SetupDiGetClassDevs _
        (HidGuid, _
        vbNullString, _
        0, _
        (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
        
     
    MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)
    If Tester.SkipOtherDevice <> "enter skip number" Then
        result = SetupDiEnumDeviceInterfaces _
            (DeviceInfoSet, _
            0, _
            HidGuid, _
            Tester.SkipOtherDevice, _
            MyDeviceInterfaceData)
    Else
            result = SetupDiEnumDeviceInterfaces _
            (DeviceInfoSet, _
            0, _
            HidGuid, _
            MemberIndex, _
            MyDeviceInterfaceData)
    End If
        
 
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
 
    If InStr(1, DevicePathName, Vid_PID) Or InStr(1, DevicePathName, "9101") Or InStr(1, DevicePathName, "9102") Then
        GetDeviceName_NoReply = DevicePathName
        Exit Function
    Else
        GetDeviceName_NoReply = ""
    End If
Next


End Function

Public Function WaitDevOn(Vid_PID As String) As Integer

Dim PassTime As Long
Dim OldTimer As Long

    OldTimer = Timer
    
    WaitDevOn = 0
    
    Do
        DoEvents
        If GetDeviceName_NoReply(Vid_PID) <> "" Then
            WaitDevOn = 1
        End If
        Call MsecDelay(0.1)
        PassTime = Timer - OldTimer
    Loop Until (WaitDevOn = 1) Or (PassTime >= 3)
    
End Function

Public Function WaitHUBOn(Vid_PID As String) As Integer

Dim PassTime As Long
Dim OldTimer As Long

    OldTimer = Timer
    
    WaitHUBOn = 0
    
    Do
        DoEvents
        If GetDeviceNameHub_NoReply(Vid_PID) <> "" Then
            WaitHUBOn = 1
        End If
        Call MsecDelay(0.1)
        PassTime = Timer - OldTimer
    Loop Until (WaitHUBOn = 1) Or (PassTime >= 4)
    
End Function

Public Function WaitDevOFF(Vid_PID As String) As Integer

Dim PassTime As Long
Dim OldTimer As Long

    OldTimer = Timer
    WaitDevOFF = 0
    'Call MsecDelay(0.1)
    
    Do
        DoEvents
        If GetDeviceName_NoReply(Vid_PID) = "" Then
            WaitDevOFF = 1
        End If
        Call MsecDelay(0.1)
        PassTime = Timer - OldTimer
    Loop Until (WaitDevOFF = 1) Or (PassTime >= 10)
    
End Function

Public Function WaitHUBOFF(Vid_PID As String) As Integer

Dim PassTime As Long
Dim OldTimer As Long

    OldTimer = Timer
    WaitHUBOFF = 0
    'Call MsecDelay(0.1)
    
    Do
        DoEvents
        If GetDeviceNameHub_NoReply(Vid_PID) = "" Then
            WaitHUBOFF = 1
        End If
        Call MsecDelay(0.1)
        PassTime = Timer - OldTimer
    Loop Until (WaitHUBOFF = 1) Or (PassTime >= 10)
    
End Function
