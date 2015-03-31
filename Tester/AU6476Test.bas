Attribute VB_Name = "AU6476Mdl"
Global AU6476Upper As Single
Global AU6476Lower As Single
Public Declare Function fnGetDeviceHandle Lib "BtnStatus.dll" (ByRef InData As Long) As Byte
 
Public Declare Function fnInquiryBtnStatus Lib "BtnStatus.dll" (ByRef OutData As Long) As Byte
 
Public Declare Function fnFreeDeviceHandle Lib "BtnStatus.dll" (ByRef InData As Long) As Byte
'Public Declare Function fnFreeDeviceHandle Lib "BtnStatus.dll" (ByRef InData As Long) As Integer
'Public Declare Function fnGetDeviceHandle Lib "BtnStatus.dll" (ByRef InData As Long) As Integer
'Public Declare Function fnInquiryBtnStatus Lib "BtnStatus.dll" (ByRef InData As Long) As Integer


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

  Public Sub AU6476JLOT5TestSub()
' add XD MS data pin bonding error sorting
 
' 980827  : set overcurrent flow , is test unit ready only
Tester.Print "3.43~3.17  V verify"
Dim TmpChip As String
Dim RomSelector As Byte
               
Call PowerSet2(0, "3.45", "0.5", 1, "3.17", "0.5", 1)
   
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
CardResult = DO_WritePort(card, Channel_P1B, &H0)

 If ChipName = "AU6368A1" Or ChipName = "AU6376" Then
        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
   ' CardResult = DO_WritePort(card, Channel_P1A, &H3E)  ' 1111 1110
End If
                  
 
                  
  
   CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
 
     Call MsecDelay(0.6)
      VidName = "vid_058f"
               
              
  
  
  
  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
             
            
               
 '================================= Test light off =============================
                
            
             
                  TestResult = ""
                  
                   
                  
                    CardResult = DO_WritePort(card, Channel_P1A, &H3D)  ' 0110 0100  only CF open
                   
                     Call MsecDelay(1.5)
                    
                   Call MsecDelay(0.1)
                 
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 '
                 '  R/W test
                 '
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
                'initial return value
                
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
               
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
   
                
                 rv0 = CBWTest_NewOverCurrent2(1, 1, "058f")     ' Test CF at 1st slot
                 Tester.Print "rv0="; rv0
                 
                 If rv0 <> 5 Then
                 GoTo AU6377ALFResult
                 End If
                 
                  CardResult = DO_WritePort(card, Channel_P1A, &HBF)  ' 0110 0100  only CF open
                   
                   Call MsecDelay(0.5)
                    CardResult = DO_WritePort(card, Channel_P1A, &H2D)
                     Call MsecDelay(1.2)
                     ReaderExist = 0
                     
                  If rv0 = 5 Then
                 
                  rv1 = CBWTest_NewOverCurrent2(1, 1, "058f")
                  
                  End If
                  Tester.Print "rv1="; rv1
                 
                 
                 If rv1 <> 6 Then
                 GoTo AU6377ALFResult
                 End If
       
                
       
                 
                
AU6377ALFResult:
                        
                        If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                           
                        ElseIf rv0 = 7 Or rv1 = 7 Or rv0 = 6 Or rv1 = 5 Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "Sorting fail"
                            
                        ElseIf rv0 = 2 Or rv1 = 2 Or rv1 = UNKNOW Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
                               
                   
                            
                        ElseIf rv0 = 5 And rv1 = 6 Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
               CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
                  
               
                    
   End Sub
   
   Public Sub AU6476JLOT6TestSub()
' add XD MS data pin bonding error sorting
 
' 980827  : set overcurrent flow , is test unit ready only
Tester.Print "3.43~3.17  V verify"
Dim TmpChip As String
Dim RomSelector As Byte

AU6476Upper = 4.43
AU6476Lower = 3.17
               
'Call PowerSet2(0, "3.43", "0.5", 1, "3.18", "0.5", 1)
   
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
 
 result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
 
 
                  
  
  ' CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
 
     Call MsecDelay(0.6)
      VidName = "vid_058f"
               
              
  
  
  
  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
             
            
               
 '================================= Test light off =============================
                
            
             
                  TestResult = ""
                  
                   
                  
                    CardResult = DO_WritePort(card, Channel_P1A, &H3D)  ' 0110 0100  only CF open
                   
                     Call MsecDelay(1.5)
                    
                   Call MsecDelay(0.1)
                 
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 '
                 '  R/W test
                 '
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
                'initial return value
                
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
               
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
   
                
                 rv0 = CBWTest_NewOverCurrent3(1, 1, "058f")     ' Test CF at 1st slot
                 Tester.Print "rv0="; rv0
                 
               
                 
        
                      
                
       
                 
                
AU6377ALFResult:
                        
                        If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                           
                        ElseIf rv0 = 5 Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "Sorting fail < lower limit "
                            
                        ElseIf rv0 = 2 Or v0 = 7 Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
                               
                         ElseIf rv0 = 6 Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "Sorting fail > upper limit"
                            
                        ElseIf rv0 = 1 Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
               CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
                  
               
                    
   End Sub
   
Public Sub AU6476KLF20TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
'**********************************************************************************************
'*                             This Sub copy from AU6476DLF24                                 *
'**********************************************************************************************
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
    PCI7248Exist
    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    Call MsecDelay(0.2)
End If
                
  
                  
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         Call MsecDelay(0.3)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
                    
        ' CardResult = DO_WritePort(card, Channel_P1A, &H3F)
        
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
               
                 VidName = "vid_18e3"
                 
                
               
                
                
                'If TmpChip = "AU6476DLF20" Then
                
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 254 And LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                             rv0 = 2
                         End If
          
                'End If
                
                
             
'====================================== Assing R/W test switch =====================================
                 
                 
                   TestResult = ""
                   'If ChipName = "AU6476DLF20" Then
                    'CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0000 0100
                   'End If
              
                    
                   'Call MsecDelay(0.1)
                 
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 '
                 '  R/W test
                 '
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
                'initial return value
                
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
               
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                
                
                ClosePipe
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                    
                    If rv0 <> 1 Then
                       rv0 = 2
                       Tester.Print "SD bus width Fail"
                    End If
                 End If
                 
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                
                rv1 = CBWTest_NewOverCurrent(1, rv0, VidName)   'AU6476E55 OverCurrent enable for all bonding type
                If rv1 = 1 Then
                    rv1 = CBWTest_New(1, rv0, VidName)    ' CF slot
                End If
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                
                If rv2 = 1 And rv3 <> 1 Then
                    Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
                MsecDelay (0.1)
                rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
                If rv4 = 1 Then
                    rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                    
                    If rv4 <> 1 Then
                        rv4 = 2
                        Tester.Print "MS bus width Fail"
                    End If
                
                End If
                   
              
            
                ClosePipe
                Call LabelMenu(3, rv4, rv3)
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
            CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 high, 7th bit of control 1, and pwr off  for HID mode
            'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
            Call MsecDelay(0.2)
          
            CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
            Call MsecDelay(1.3)
            
            'i = 0
            
            'Do
            '    CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                
            '    If (LightOFF <> 252) And (LightOFF <> 254) Then
            '        rv4 = 1
            '    Else
            '        rv4 = 2
            '        Call MsecDelay(0.5)
            '        'Tester.Text26.Text = CInt(Tester.Text26) + 1
            '    End If
            '    i = i + 1
            
            'Loop Until (rv4 = 1) Or (i > 6)
        
        'If rv4 <> 1 Then
        '    Tester.Label9.Caption = "GPO FAIL " & LightOFF
        '    UsbSpeedTestResult = GPO_FAIL
        '    rv4 = 2
        'End If
     
     End If
          
        If rv4 = 1 Then
        ' code begin
                     
                     Tester.Cls
                     Tester.Print "keypress test begin---------------"
                     Dim ReturnValue As Byte
                     
                     DeviceHandle = &HFFFF  'invalid handle initial value
                     
                     ReturnValue = fnGetDeviceHandle(DeviceHandle)
                     Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
                     Tester.Print "DeviceHandle="; DevicehHandle
                     
                        If ReturnValue <> 1 Then
                            rv0 = UNKNOW       '---> HID mode unknow device mode
                            Call LabelMenu(0, rv0, 1)
                            Tester.Label9.Caption = "HID mode unknow device"
                            fnFreeDeviceHandle (DeviceHandle)
                            GoTo AU6377ALFResult
                        End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                     i = 0
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 5 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
                     If ReturnValue <> 10 Then
                     
                      rv1 = 2
                     
                       Call LabelMenu(1, rv1, rv0)
                       Label9.Caption = "KeyPress Fail"
                       
                     End If
                              
                 
           End If
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
                        If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
                        ElseIf rv0 = READ_FAIL Then
                            SDReadFail = SDReadFail + 1
                            TestResult = "SD_RF"
                        ElseIf rv1 = WRITE_FAIL Then
                            CFWriteFail = CFWriteFail + 1
                            TestResult = "CF_WF"
                        ElseIf rv1 = READ_FAIL Then
                            CFReadFail = CFReadFail + 1
                            TestResult = "CF_RF"
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
               
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
    
    Public Sub AU6476JLOT1TestSub()
' add XD MS data pin bonding error sorting
 
' 980827  : set overcurrent flow , is test unit ready only
Tester.Print "3.25 V verify"
Dim TmpChip As String
Dim RomSelector As Byte
               
  Call PowerSet2(0, "3.25", "0.5", 1, "3.25", "0.5", 1)
  If ChipName = "AU6370GLF20" Then
      ChipName = "AU6370DLF20"
  End If
                
                ' open power
 If ChipName = "AU6377ALF24" Or ChipName = "AU6377ALF25" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
                
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
            
                'GPIO control setting
If ChipName = "AU6370BL" Or InStr(ChipName, "AU6375HL") <> 0 Or ChipName = "AU6375CL" Or ChipName = "AU6377ALF21" Or ChipName = "AU6377ALS10" Then
     TmpChip = ChipName
     ChipName = "AU6376"
End If
                
If ChipName = "AU6376ELF22" Or ChipName = "AU6376ILF20" Then
      ChipName = "AU6376"
End If
                
If ChipName = "AU6376JLF20" Then
      TmpChip = ChipName
     ChipName = "AU6376"
End If
                
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
If ChipName = "AU6368A" Then
       CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 0111 1111
           result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
End If
If ChipName = "AU6368A1" Or ChipName = "AU6376" Then
        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    CardResult = DO_WritePort(card, Channel_P1A, &H3E)  ' 1111 1110
End If
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
  End If
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  
 
  
  
  
  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
               
 '================================= Test light off =============================
                
            
             
                  TestResult = ""
                  
                   
                  
                    CardResult = DO_WritePort(card, Channel_P1A, &H3D)  ' 0110 0100  only CF open
                   
     
                    
                   Call MsecDelay(0.1)
                 
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 '
                 '  R/W test
                 '
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
                'initial return value
                
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
               
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
   
                
                 rv0 = CBWTest_NewOverCurrent(1, 1, VidName)     ' Test CF at 1st slot
                 
       
                 Call LabelMenu(0, rv0, 1)
                     
                    If rv0 = 3 Then
                    Tester.Label9.Caption = "Over Current fail"
                    Tester.Label2.Caption = "Over Current fail---"
                    End If
               
                   
                Tester.Print rv0, " \\CF :0 Unknow device, 1 pass ,2   fail, 3 over currnt, 4 preious slot fail"
                
               
                Tester.Print "LBA="; LBA
    
                 
                
AU6377ALFResult:
                        
                        If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = 2 Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
                        ElseIf rv0 = 3 Then
                        
                             XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                            
                            
                        ElseIf rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
               
                  
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
   End Sub
   
    Public Sub AU6476JLOT2TestSub()
' add XD MS data pin bonding error sorting
 
' 980827  : set overcurrent flow , is test unit ready only
Tester.Print "3.20 V verify"
Dim TmpChip As String
Dim RomSelector As Byte
               
  Call PowerSet2(0, "3.20", "0.5", 1, "3.20", "0.5", 1)
  If ChipName = "AU6370GLF20" Then
      ChipName = "AU6370DLF20"
  End If
                
                ' open power
 If ChipName = "AU6377ALF24" Or ChipName = "AU6377ALF25" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
                
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
            
                'GPIO control setting
If ChipName = "AU6370BL" Or InStr(ChipName, "AU6375HL") <> 0 Or ChipName = "AU6375CL" Or ChipName = "AU6377ALF21" Or ChipName = "AU6377ALS10" Then
     TmpChip = ChipName
     ChipName = "AU6376"
End If
                
If ChipName = "AU6376ELF22" Or ChipName = "AU6376ILF20" Then
      ChipName = "AU6376"
End If
                
If ChipName = "AU6376JLF20" Then
      TmpChip = ChipName
     ChipName = "AU6376"
End If
                
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
If ChipName = "AU6368A" Then
       CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 0111 1111
           result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
End If
If ChipName = "AU6368A1" Or ChipName = "AU6376" Then
        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    CardResult = DO_WritePort(card, Channel_P1A, &H3E)  ' 1111 1110
End If
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
  End If
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  
 
  
  
  
  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
               
 '================================= Test light off =============================
                
            
             
                  TestResult = ""
                  
                   
                  
                    CardResult = DO_WritePort(card, Channel_P1A, &H3D)  ' 0110 0100  only CF open
                   
     
                    
                   Call MsecDelay(0.1)
                 
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 '
                 '  R/W test
                 '
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
                'initial return value
                
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
               
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
   
                
                 rv0 = CBWTest_NewOverCurrent(1, 1, VidName)     ' Test CF at 1st slot
                 
       
                 Call LabelMenu(0, rv0, 1)
                     
                    If rv0 = 3 Then
                    Tester.Label9.Caption = "Over Current fail"
                    Tester.Label2.Caption = "Over Current fail---"
                    End If
               
                   
                Tester.Print rv0, " \\CF :0 Unknow device, 1 pass ,2   fail, 3 over currnt, 4 preious slot fail"
                
               
                Tester.Print "LBA="; LBA
    
                 
                
AU6377ALFResult:
                        
                        If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = 2 Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
                        ElseIf rv0 = 3 Then
                        
                             XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                            
                            
                        ElseIf rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
               
                  
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
   End Sub

    Public Sub AU6476JLOT3TestSub()
' add XD MS data pin bonding error sorting
 
' 980827  : set overcurrent flow , is test unit ready only
Tester.Print "3.15 V verify"
Dim TmpChip As String
Dim RomSelector As Byte
               
  Call PowerSet2(0, "3.15", "0.5", 1, "3.15", "0.5", 1)
  If ChipName = "AU6370GLF20" Then
      ChipName = "AU6370DLF20"
  End If
                
                ' open power
 If ChipName = "AU6377ALF24" Or ChipName = "AU6377ALF25" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
                
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
            
                'GPIO control setting
If ChipName = "AU6370BL" Or InStr(ChipName, "AU6375HL") <> 0 Or ChipName = "AU6375CL" Or ChipName = "AU6377ALF21" Or ChipName = "AU6377ALS10" Then
     TmpChip = ChipName
     ChipName = "AU6376"
End If
                
If ChipName = "AU6376ELF22" Or ChipName = "AU6376ILF20" Then
      ChipName = "AU6376"
End If
                
If ChipName = "AU6376JLF20" Then
      TmpChip = ChipName
     ChipName = "AU6376"
End If
                
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
If ChipName = "AU6368A" Then
       CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 0111 1111
           result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
End If
If ChipName = "AU6368A1" Or ChipName = "AU6376" Then
        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    CardResult = DO_WritePort(card, Channel_P1A, &H3E)  ' 1111 1110
End If
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
  End If
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  
 
  
  
  
  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
               
 '================================= Test light off =============================
                
            
             
                  TestResult = ""
                  
                   
                  
                    CardResult = DO_WritePort(card, Channel_P1A, &H3D)  ' 0110 0100  only CF open
                   
     
                    
                   Call MsecDelay(0.1)
                 
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 '
                 '  R/W test
                 '
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
                'initial return value
                
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
               
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
   
                
                 rv0 = CBWTest_NewOverCurrent(1, 1, VidName)     ' Test CF at 1st slot
                 
       
                 Call LabelMenu(0, rv0, 1)
                     
                    If rv0 = 3 Then
                    Tester.Label9.Caption = "Over Current fail"
                    Tester.Label2.Caption = "Over Current fail---"
                    End If
               
                   
                Tester.Print rv0, " \\CF :0 Unknow device, 1 pass ,2   fail, 3 over currnt, 4 preious slot fail"
                
               
                Tester.Print "LBA="; LBA
    
                 
                
AU6377ALFResult:
                        
                        If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = 2 Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
                        ElseIf rv0 = 3 Then
                        
                             XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                            
                            
                        ElseIf rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
               
                  
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
   End Sub
Public Sub AU6476_020222TestSub()

    If ChipName = "AU6476CLF27" Then
        Call AU6476CLF27TestSub
    ElseIf ChipName = "AU6476CLF07" Then
        Call AU6476CLF07TestSub
    ElseIf ChipName = "AU6476VLF26" Then
        Call AU6476VLF26TestSub
    ElseIf ChipName = "AU6476YLF26" Then
        Call AU6476YLF26TestSub
    Else
        TestResult = "Bin2"
    End If
    
End Sub


Public Sub AU6476TestSub()
              If ChipName = "AU6420ALF20" Then
              
                    Call AU6420ALTestSub
                ElseIf ChipName = "AU6420BLF20" Then
            
                    Call AU6420BLTestSub
                
                ElseIf ChipName = "AU6420BLF2A" Then
            
                    Call AU6420BLF2ASub
                    
                ElseIf ChipName = "AU6420BLS2A" Then
            
                    Call AU6420BLS2ASub
                
               ElseIf ChipName = "AU6420BLS20" Then
            
                    Call AU6420BLS20Sub
               
               ElseIf ChipName = "AU6420BLS30" Then
            
                    Call AU6420BLS30Sub
                
               ElseIf ChipName = "AU6420CLF20" Then
            
                    Call AU6420CLTestSub
               ElseIf ChipName = "AU6476BLF20" Then
            
                    Call AU6476BLTestSub
               ElseIf ChipName = "AU6476CLF20" Then
            
                    Call AU6476CLTestSub
               ElseIf ChipName = "AU6476CLF21" Then
            
                    Call AU6476CLF21TestSub
               ElseIf ChipName = "AU6476CLF22" Then
            
                    Call AU6476CLF22TestSub
               ElseIf ChipName = "AU6476CLF24" Then
            
                    Call AU6476CLF24TestSub
                    
                ElseIf ChipName = "AU6476CLF25" Then
            
                    Call AU6476CLF25TestSub
                
                ElseIf ChipName = "AU6476CLF26" Then
            
                    Call AU6476CLF26TestSub
            
                ElseIf ChipName = "AU6476CLF06" Then
            
                    Call AU6476CLF06TestSub
                
               ElseIf ChipName = "AU6476MLF24" Then
            
                    Call AU6476MLF24TestSub
                    
              ElseIf ChipName = "AU6476DLF20" Or ChipName = "AU6476DLF21" Or ChipName = "AU6476DLF22" Or ChipName = "AU6476DLF24" Then
            
                    Call AU6476DLTestSub
              ElseIf ChipName = "AU6476FLF20" Then
            
                    Call AU6476FLTestSub
              ElseIf ChipName = "AU6476FLF21" Then
            
                    Call AU6476FL21TestSub
              ElseIf ChipName = "AU6476FLF24" Then
            
                    Call AU6476FL24TestSub
               ElseIf ChipName = "AU6476ELF20" Then
              
                    Call AU6476ELTestSub
               ElseIf ChipName = "AU6476ELF21" Then
            
                    Call AU6476ELF21TestSub
               ElseIf ChipName = "AU6476ELF24" Then
            
                    Call AU6476ELF24TestSub
                ElseIf ChipName = "AU6476BLF21" Then
            
                    Call AU6476BLF21TestSub
                 ElseIf ChipName = "AU6476BLF22" Then
            
                    Call AU6476BLF22TestSub
                    
                ElseIf ChipName = "AU6476BLF23" Then
            
                    Call AU6476BLF23TestSub
                    
                ElseIf ChipName = "AU6476BLF24" Then
            
                    Call AU6476BLF24TestSub
                    
                ElseIf ChipName = "AU6476BLF25" Then
            
                    Call AU6476BLF25TestSub
                
                ElseIf ChipName = "AU6476BLF26" Then
                
                    Call AU6476BLF26TestSub
                
                ElseIf ChipName = "AU6476WLF05" Then
            
                    Call AU6476WLF05TestSub
                    
                ElseIf ChipName = "AU6476WLF06" Then
            
                    Call AU6476WLF06TestSub
                
                ElseIf ChipName = "AU6476WLF35" Then
            
                    Call AU6476WLF35TestSub
                    
                ElseIf ChipName = "AU6476WLF36" Then
            
                    Call AU6476WLF36TestSub
                
                ElseIf ChipName = "AU6476WLF3E" Then
            
                    Call AU6476WLF3ETestSub
                    
                ElseIf ChipName = "AU6476WLF3F" Then
            
                    Call AU6476WLF3FTestSub
                  
                  ElseIf ChipName = "AU6476LLF25" Then
            
                    Call AU6476LLF25TestSub
                    
                  ElseIf ChipName = "AU6476QLF25" Then
            
                    Call AU6476QLF25TestSub
                    
                     ElseIf ChipName = "AU6476RLF25" Then
            
                    Call AU6476RLF25TestSub
                    
                       ElseIf ChipName = "AU6476RLF26" Then
            
                    Call AU6476RLF26TestSub
                    
                       ElseIf ChipName = "AU6476RLF27" Then
            
                    Call AU6476RLF27TestSub
                    
                 ElseIf ChipName = "AU6476ILF21" Then
            
                    Call AU6476ILF21TestSub
                    
                 ElseIf ChipName = "AU6476ILF22" Then
            
                    Call AU6476ILF22TestSub
                                 
                  
                  ElseIf ChipName = "AU6476ILF24" Then
            
                    Call AU6476ILF24TestSub
                                   
                  ElseIf ChipName = "AU6476JLO10" Then
            
                    Call AU6476JLO10TestSub
                    
                   ElseIf ChipName = "AU6476JLO20" Then
            
                    Call AU6476JLO20TestSub
                    
                   ElseIf ChipName = "AU6476JLOT1" Then
            
                    Call AU6476JLOT1TestSub
                    
                    ElseIf ChipName = "AU6476JLOT2" Then
            
                    Call AU6476JLOT2TestSub
                    
                     ElseIf ChipName = "AU6476JLOT3" Then
            
                    Call AU6476JLOT3TestSub
                    
                       ElseIf ChipName = "AU6476JLOT6" Then
            
                    Call AU6476JLOT6TestSub
                    
                    ElseIf ChipName = "AU6476KLF20" Then
            
                    Call AU6476KLF20TestSub
                    
                  End If
                  
                  
                  
                  
End Sub

