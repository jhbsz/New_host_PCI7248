Attribute VB_Name = "AU6479Test"
Public Sub AU6479TestSub()

    If ChipName = "AU6479ALF20" Then
        Call AU6479ALF20TestSub
    ElseIf ChipName = "AU6479AFF33" Then
        Call AU6479AFF33TestSub
    ElseIf ChipName = "AU6479BLF20" Then
        Call AU6479BLF20TestSub
    ElseIf ChipName = "AU6479BLF21" Then
        Call AU6479BLF21TestSub
    ElseIf ChipName = "AU6479HLF21" Then
        Call AU6479HLF21TestSub
    ElseIf ChipName = "AU6479BLF22" Then
        Call AU6479BLF22TestSub
    ElseIf (ChipName = "AU6479BLF23") Or (ChipName = "AU6479ILF23") Or (ChipName = "AU6479JLF23") Or (ChipName = "AU6479NLF23") Then     'change to 4LUN socket-board
        Call AU6479BLF23TestSub
    ElseIf ChipName = "AU6479BLF02" Then
        Call AU6479BLF02TestSub
    ElseIf ChipName = "AU6479HLF22" Then
        Call AU6479HLF22TestSub
    ElseIf ChipName = "AU6479JLFE3" Then
        Call AU6479JLFE3TestSub
    ElseIf (ChipName = "AU6479KLF23") Or (ChipName = "AU6479KLF24") Then
        Call AU6479KLF24TestSub
    ElseIf (ChipName = "AU6479FLF23") Or (ChipName = "AU6479ULF23") Or (ChipName = "AU6479FLF24") Or (ChipName = "AU6479ULF24") Then
        Call AU6479FLF23TestSub
    ElseIf ChipName = "AU6479FLF03" Then
        Call AU6479FLF03TestSub
    ElseIf ChipName = "AU6479OLF23" Then
        Call AU6479OLF23TestSub
    ElseIf ChipName = "AU6479OLT10" Then
        Call AU6479OLT10TestSub
    ElseIf ChipName = "AU6479OLF03" Then
        Call AU6479OLF03TestSub
    ElseIf ChipName = "AU6479ALT10" Then
        Call AU6479ALT10TestSub
    ElseIf ChipName = "AU6479TLF23" Then
        Call AU6479TLF23TestSub
    ElseIf ChipName = "AU6479BFF23" Then
        Call AU6479BFF23TestSub
    ElseIf ChipName = "AU6479CFF23" Then
        Call AU6479CFF23TestSub
    ElseIf ChipName = "AU6479DFF24" Then
        Call AU6479DFF24TestSub
    ElseIf ChipName = "AU6479WLF22" Then
        Call AU6479WLF22TestSub
    ElseIf ChipName = "AU6479PLF32" Then
        Call AU6479PLF32TestSub
    End If

End Sub

Public Sub AU6479BLF20TestSub()

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

'Call PowerSet2(0, "3.3", "0.7", 1, "3.3", "0.7", 1)

Tester.Print "AU6479BL : Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'CF  (Lun0)
rv2 = 0     'SD0 (Lun1)
rv3 = 0     'XD  (Lun3)
rv4 = 0     'MS2 (Lun4)
rv5 = 0     'MS1 (Lun2)
rv6 = 0     'SD1 (Lun4)

Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '            CF ¡BSD0 ¡B XD ¡BMS2                       CF ¡BMS1 ¡B XD ¡BSD1
                'Condition1(Lun0¡BLun1¡BLun3¡BLun4)         Condition2(Lun0¡BLun2¡BLun3¡BLun4)
'8:ENA    ---               0                                           0
'7:HID    ---               0                                           0
'6:M2INS  ---               0                                           1
'5:SD1CDN ---               1                                           0

'4:MSINS  ---               1                                           0
'3:XDCDN  ---               0                                           0
'2:CFCDN  ---               0                                           0
'1:SD0CDN ---               0                                           1

'                         0x18                                        0x21

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H18)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'If CardResult <> 0 Then
'    MsgBox "Read light off fail"
'    End
'End If
'
'Call MsecDelay(0.2)     'power on time
ChipString = "058f"
                
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOFF
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479BLFResult
End If



Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  CF Card test Lun0
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "CF", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
    Call NewLabelMenu(1, "CF_64K", rv1, rv0)
End If
Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  SD Card test Lun1
'===============================================

rv2 = CBWTest_New_PipeReady(1, rv1, ChipString)
If rv2 <> 1 Then
    Call NewLabelMenu(2, "SD", rv2, rv1)
Else
    rv2 = CBWTest_New_128_Sector_PipeReady(1, rv2)  ' write
    Call NewLabelMenu(2, "SD_64K", rv2, rv1)
    
    If rv2 = 1 Then
        rv2 = Read_Speed2ReadData(LBA, 1, 64)
        If rv2 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "SD is 4 Bit, Speed 100 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv2 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(2, "SD Bus Speed/Width", rv2, rv1)
    End If
    
End If

Tester.Print rv2, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'=======================================================================================
'       SDHC R / W
'=======================================================================================
        
'Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
'OpenPipe
'rv1 = ReInitial(0)
'Call MsecDelay(0.02)
'rv1 = AU6435ForceSDHC(rv0)
'ClosePipe
'
'If rv1 = 1 Then
'    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
'End If
'
'
'If rv1 = 1 Then
'    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
'    If rv1 <> 1 Then
'        rv1 = 2
'        Tester.Print "SD2.0 Mode Fail"
'    End If
'End If
'
'ClosePipe
'
'Call LabelMenu(1, rv1, rv0)
'
'Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
'===============================================
'  XD Card test Lun3
'================================================
rv3 = CBWTest_New_PipeReady(3, rv2, ChipString)
If rv3 <> 1 Then
    Call NewLabelMenu(3, "XD", rv3, rv2)
Else
    rv3 = CBWTest_New_128_Sector_PipeReady(3, rv3)  ' write
    Call NewLabelMenu(3, "XD_64", rv3, rv2)
End If

Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  MSPro MS2 Card test Lun4
'================================================
rv4 = CBWTest_New_PipeReady(4, rv3, ChipString)
If rv4 <> 1 Then
    Call NewLabelMenu(4, "MS2", rv4, rv3)
Else
    rv4 = CBWTest_New_128_Sector_PipeReady(4, rv4)  ' write
    Call NewLabelMenu(4, "MS2_64k", rv4, rv3)
    
    If rv4 = 1 Then
        rv4 = Read_Speed2ReadData(LBA, 4, 64)
        If rv4 = 1 Then
            If (ReadData(25) = &H2) And (ReadData(26) = &H70) And (ReadData(27) = &H1) Then
                Tester.Print "MS is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS BusWidth/Speed Fail"
                rv4 = 3
            End If
        End If
        Call NewLabelMenu(4, "MS2 Bus Width/Speed", rv4, rv3)
    End If
End If
Tester.Print rv4, " \\MSpro_2 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

If rv4 <> 1 Then
    GoTo AU6479BLFResult
End If

ClosePipe
CardResult = DO_WritePort(card, Channel_P1A, &H21)      'Lun2: MS1, Lun4: SD1
Call MsecDelay(0.08)
OpenPipe
Call MsecDelay(0.04)
rv5 = ReInitial(1)
If rv5 = 1 Then
    rv5 = ReInitial(4)
End If
ClosePipe
   
If rv5 <> 1 Then
    Call NewLabelMenu(5, "ReNew", rv5, rv4)
    GoTo AU6479BLFResult
Else
    Call MsecDelay(0.2)
End If
   
'===============================================
'  MSPro MS1 Card test Lun2
'================================================
rv5 = CBWTest_New(2, rv5, ChipString)
If rv5 <> 1 Then
    Call NewLabelMenu(5, "MS", rv5, rv4)
Else
    rv5 = CBWTest_New_128_Sector_PipeReady(2, rv5)  ' write
    Call NewLabelMenu(5, "MS_64K", rv5, rv4)

    If rv5 = 1 Then
        rv5 = Read_Speed2ReadData(LBA, 2, 64)
        If rv5 = 1 Then
            If (ReadData(22) = &H2) And (ReadData(23) = &H70) And (ReadData(24) = &H1) Then
                Tester.Print "MS is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS BusWidth/Speed Fail"
                rv5 = 3
            End If
        End If
        Call NewLabelMenu(5, "MS Bus Width/Speed", rv5, rv4)
    End If
End If
Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
 
'===============================================
'  SD_1 Card test Lun4
'===============================================
rv6 = CBWTest_New_PipeReady(4, rv5, ChipString)
If rv6 <> 1 Then
    Call NewLabelMenu(5, "SD1", rv6, rv5)
Else
    rv6 = CBWTest_New_128_Sector_PipeReady(4, rv6)  ' write
    Call NewLabelMenu(5, "SD1_64K", rv6, rv5)

    If rv6 = 1 Then
        rv6 = Read_Speed2ReadData(LBA, 4, 64)
        If rv6 = 1 Then
            If (ReadData(12) = &HA) Then
                Tester.Print "SD is 4 Bit, Speed 100 MHz"
            'Else
            '    Tester.Print "SD BusWidth/Speed Fail"
            '    rv6 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(5, "SD1 Bus Width/Speed", rv6, rv5)
    End If
End If


Tester.Print rv6, " \\SDXC_1 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


AU6479BLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &H98)   ' Close power
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'CF
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'SD0
        TestResult = "Bin3"
    ElseIf rv3 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv4 <> 1 Then        'MS2
        TestResult = "Bin5"
    ElseIf rv5 <> 1 Then        'MS1
        TestResult = "Bin5"
    ElseIf rv6 <> 1 Then        'SD1
        TestResult = "Bin3"
    ElseIf rv6 * rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub

Public Sub AU6479BLF21TestSub()

'2012/5/3 for S/B: AU6479-GBL 100LQ SOCKET V0.90

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

'Call PowerSet2(0, "3.3", "0.7", 1, "3.3", "0.7", 1)

Tester.Print "AU6479BL : Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'CF  (Lun0)
rv2 = 0     'SD0 (Lun1)
rv3 = 0     'MS1 (Lun2)
rv4 = 0     'XD  (Lun3)
rv5 = 0     'SD1 (Lun4)
rv6 = 0     'MS2 (Lun4)



Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '            CF ¡BSD0 ¡BMS1 ¡BXD  ¡BSD1                       CF ¡BSD0 ¡BMS1 ¡BXD  ¡BMS2(Can't not R/W)
                'Condition1(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)         Condition2(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)
'8:ENA    ---               0                                           0
'7:HID    ---               1                                           1
'6:M2INS  ---               1                                           0
'5:SD1CDN ---               0                                           1

'4:MSINS  ---               0                                           0
'3:XDCDN  ---               0                                           0
'2:CFCDN  ---               0                                           0
'1:SD0CDN ---               0                                           0

'                         0x60                                        0x50

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H60)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'If CardResult <> 0 Then
'    MsgBox "Read light off fail"
'    End
'End If
'
'Call MsecDelay(0.2)     'power on time
ChipString = "058f"
                
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOFF
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479BLFResult
End If



Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  CF Card test Lun0
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "CF", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
    Call NewLabelMenu(1, "CF_64K", rv1, rv0)
End If
Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  SD Card test Lun1
'===============================================

rv2 = CBWTest_New_PipeReady(1, rv1, ChipString)
If rv2 <> 1 Then
    Call NewLabelMenu(2, "SD", rv2, rv1)
Else
    rv2 = CBWTest_New_128_Sector_PipeReady(1, rv2)  ' write
    Call NewLabelMenu(2, "SD_64K", rv2, rv1)
    
    If rv2 = 1 Then
        rv2 = Read_Speed2ReadData(LBA, 1, 64)
        If rv2 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv2 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(2, "SD Bus Speed/Width", rv2, rv1)
    End If
    
End If

Tester.Print rv2, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'=======================================================================================
'       SDHC R / W
'=======================================================================================
        
'Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
'OpenPipe
'rv1 = ReInitial(0)
'Call MsecDelay(0.02)
'rv1 = AU6435ForceSDHC(rv0)
'ClosePipe
'
'If rv1 = 1 Then
'    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
'End If
'
'
'If rv1 = 1 Then
'    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
'    If rv1 <> 1 Then
'        rv1 = 2
'        Tester.Print "SD2.0 Mode Fail"
'    End If
'End If
'
'ClosePipe
'
'Call LabelMenu(1, rv1, rv0)
'
'Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
'===============================================
'  MSPro MS1 Card test Lun2
'================================================
rv3 = CBWTest_New(2, rv2, ChipString)
If rv3 <> 1 Then
    Call NewLabelMenu(3, "MS", rv3, rv2)
Else
    rv3 = CBWTest_New_128_Sector_PipeReady(2, rv3)  ' write
    Call NewLabelMenu(3, "MS_64K", rv3, rv2)

    If rv3 = 1 Then
        rv3 = Read_Speed2ReadData(LBA, 2, 64)
        If rv3 = 1 Then
            If (ReadData(22) = &H2) And (ReadData(23) = &H70) And (ReadData(24) = &H1) Then
                Tester.Print "MS is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS BusWidth/Speed Fail"
                rv3 = 3
            End If
        End If
        Call NewLabelMenu(3, "MS Bus Width/Speed", rv3, rv2)
    End If
End If
Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
                  
'===============================================
'  XD Card test Lun3
'================================================
rv4 = CBWTest_New_PipeReady(3, rv3, ChipString)
If rv4 <> 1 Then
    Call NewLabelMenu(4, "XD", rv3, rv2)
Else
    rv4 = CBWTest_New_128_Sector_PipeReady(3, rv4)  ' write
    Call NewLabelMenu(4, "XD_64", rv3, rv2)
End If

Tester.Print rv4, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  SD1 Card test Lun4
'===============================================
rv5 = CBWTest_New_PipeReady(4, rv4, ChipString)
If rv5 <> 1 Then
    Call NewLabelMenu(5, "SD1", rv5, rv4)
Else
    rv5 = CBWTest_New_128_Sector_PipeReady(4, rv5)   ' write
    Call NewLabelMenu(5, "SD1_64K", rv5, rv4)

    If rv5 = 1 Then
        rv5 = Read_Speed2ReadData(LBA, 4, 64)
        If rv5 = 1 Then
            If (ReadData(19) = &H6) Then
                Tester.Print "SD is 4 Bit, Speed 48 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv5 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(5, "SD1 Bus Width/Speed", rv5, rv4)
    End If
End If

Tester.Print rv5, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


ClosePipe
If rv5 <> 1 Then
    GoTo AU6479BLFResult
End If

'            CF ¡BSD0 ¡BMS1 ¡BXD  ¡BMS2(Can't not R/W)
'Condition2(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)
CardResult = DO_WritePort(card, Channel_P1A, &H50)
Call MsecDelay(0.08)
OpenPipe
Call MsecDelay(0.04)
rv5 = ReInitial(1)
If rv5 = 1 Then
    rv5 = ReInitial(4)
End If
ClosePipe

If rv5 <> 1 Then
    Call NewLabelMenu(5, "ReNew", rv5, rv4)
    GoTo AU6479BLFResult
Else
    Call MsecDelay(0.2)
End If

'===============================================
'  MSPro MS2 Card test Lun4
'================================================
OpenPipe
rv6 = RequestSense(4)
ClosePipe
If rv6 <> 1 Then
    Call NewLabelMenu(6, "MS2", rv6, rv5)
'Else
'    rv4 = CBWTest_New_128_Sector_PipeReady(4, rv4)  ' write
'    Call NewLabelMenu(4, "MS2_64k", rv4, rv3)
'
'    If rv4 = 1 Then
'        rv4 = Read_Speed2ReadData(LBA, 4, 64)
'        If rv4 = 1 Then
'            If (ReadData(25) = &H2) And (ReadData(26) = &H70) And (ReadData(27) = &H1) Then
'                Tester.Print "MS is 4 Bit, Speed 40 MHz"
'            Else
'                Tester.Print "MS BusWidth/Speed Fail"
'                rv4 = 3
'            End If
'        End If
'        Call NewLabelMenu(4, "MS2 Bus Width/Speed", rv4, rv3)
'    End If
End If
Tester.Print rv4, " \\MSpro_2 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'

AU6479BLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'CF
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'SD0
        TestResult = "Bin3"
    ElseIf rv3 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv4 <> 1 Then        'MS2
        TestResult = "Bin5"
    ElseIf rv5 <> 1 Then        'MS1
        TestResult = "Bin5"
    ElseIf rv6 <> 1 Then        'SD1
        TestResult = "Bin3"
    ElseIf rv6 * rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub

Public Sub AU6479BLF22TestSub()

'2012/5/3 for S/B: AU6479-GBL 100LQ SOCKET V0.90
'2012/6/14 for B62 version

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

'Call PowerSet2(0, "3.3", "0.7", 1, "3.3", "0.7", 1)

Tester.Print "AU6479BL : Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'CF  (Lun0)
rv2 = 0     'SD0 (Lun1)
rv3 = 0     'MS1 (Lun2)
rv4 = 0     'XD  (Lun3)
rv5 = 0     'SD1 (Lun4)
rv6 = 0     'MS2 (Lun4)



Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '            CF ¡BSD0 ¡BMS1 ¡BXD  ¡BSD1                       CF ¡BSD0 ¡BMS1 ¡BXD  ¡BMS2(Can't not R/W)
                'Condition1(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)         Condition2(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)
'8:ENA    ---               0                                           0
'7:HID    ---               1                                           1
'6:M2INS  ---               1                                           0
'5:SD1CDN ---               0                                           1

'4:MSINS  ---               0                                           0
'3:XDCDN  ---               0                                           0
'2:CFCDN  ---               0                                           0
'1:SD0CDN ---               0                                           0

'                         0x60                                        0x50

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H60)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'If CardResult <> 0 Then
'    MsgBox "Read light off fail"
'    End
'End If
'
'Call MsecDelay(0.2)     'power on time
ChipString = "058f"
                
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOFF
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479BLFResult
End If



Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  CF Card test Lun0
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "CF", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
    Call NewLabelMenu(1, "CF_64K", rv1, rv0)
End If
Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  SD Card test Lun1
'===============================================

rv2 = CBWTest_New_PipeReady(1, rv1, ChipString)
If rv2 <> 1 Then
    Call NewLabelMenu(2, "SD", rv2, rv1)
Else
    rv2 = CBWTest_New_128_Sector_PipeReady(1, rv2)  ' write
    Call NewLabelMenu(2, "SD_64K", rv2, rv1)
    
    If rv2 = 1 Then
        rv2 = Read_Speed2ReadData(LBA, 1, 64)
        If rv2 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv2 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(2, "SD Bus Speed/Width", rv2, rv1)
    End If
    
End If

Tester.Print rv2, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'=======================================================================================
'       SDHC R / W
'=======================================================================================
        
'Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
'OpenPipe
'rv1 = ReInitial(0)
'Call MsecDelay(0.02)
'rv1 = AU6435ForceSDHC(rv0)
'ClosePipe
'
'If rv1 = 1 Then
'    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
'End If
'
'
'If rv1 = 1 Then
'    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
'    If rv1 <> 1 Then
'        rv1 = 2
'        Tester.Print "SD2.0 Mode Fail"
'    End If
'End If
'
'ClosePipe
'
'Call LabelMenu(1, rv1, rv0)
'
'Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
'===============================================
'  MSPro MS1 Card test Lun2
'================================================
rv3 = CBWTest_New(2, rv2, ChipString)
If rv3 <> 1 Then
    Call NewLabelMenu(3, "MS", rv3, rv2)
Else
    rv3 = CBWTest_New_128_Sector_PipeReady(2, rv3)  ' write
    Call NewLabelMenu(3, "MS_64K", rv3, rv2)

    If rv3 = 1 Then
        rv3 = Read_Speed2ReadData(LBA, 2, 64)
        If rv3 = 1 Then
            If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                Tester.Print "MS is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS BusWidth/Speed Fail"
                rv3 = 3
            End If
        End If
        Call NewLabelMenu(3, "MS Bus Width/Speed", rv3, rv2)
    End If
End If
Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
                  
'===============================================
'  XD Card test Lun3
'================================================
rv4 = CBWTest_New_PipeReady(3, rv3, ChipString)
If rv4 <> 1 Then
    Call NewLabelMenu(4, "XD", rv3, rv2)
Else
    rv4 = CBWTest_New_128_Sector_PipeReady(3, rv4)  ' write
    Call NewLabelMenu(4, "XD_64", rv3, rv2)
End If

Tester.Print rv4, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  SD1 Card test Lun4
'===============================================
rv5 = CBWTest_New_PipeReady(4, rv4, ChipString)
If rv5 <> 1 Then
    Call NewLabelMenu(5, "SD1", rv5, rv4)
Else
    rv5 = CBWTest_New_128_Sector_PipeReady(4, rv5)   ' write
    Call NewLabelMenu(5, "SD1_64K", rv5, rv4)

    If rv5 = 1 Then
        rv5 = Read_Speed2ReadData(LBA, 4, 64)
        If rv5 = 1 Then
            If (ReadData(20) = &H6) Then
                Tester.Print "SD is 4 Bit, Speed 48 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv5 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(5, "SD1 Bus Width/Speed", rv5, rv4)
    End If
End If

Tester.Print rv5, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


ClosePipe
If rv5 <> 1 Then
    GoTo AU6479BLFResult
End If

'            CF ¡BSD0 ¡BMS1 ¡BXD  ¡BMS2(Can't not R/W)
'Condition2(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)
CardResult = DO_WritePort(card, Channel_P1A, &H50)
Call MsecDelay(0.08)
OpenPipe
Call MsecDelay(0.04)
rv5 = ReInitial(1)
If rv5 = 1 Then
    rv5 = ReInitial(4)
End If
ClosePipe

If rv5 <> 1 Then
    Call NewLabelMenu(5, "ReNew", rv5, rv4)
    GoTo AU6479BLFResult
Else
    Call MsecDelay(0.2)
End If

'===============================================
'  MSPro MS2 Card test Lun4
'================================================
OpenPipe
rv6 = RequestSense(4)
ClosePipe
If rv6 <> 1 Then
    Call NewLabelMenu(6, "MS2", rv6, rv5)
'Else
'    rv4 = CBWTest_New_128_Sector_PipeReady(4, rv4)  ' write
'    Call NewLabelMenu(4, "MS2_64k", rv4, rv3)
'
'    If rv4 = 1 Then
'        rv4 = Read_Speed2ReadData(LBA, 4, 64)
'        If rv4 = 1 Then
'            If (ReadData(25) = &H2) And (ReadData(26) = &H70) And (ReadData(27) = &H1) Then
'                Tester.Print "MS is 4 Bit, Speed 40 MHz"
'            Else
'                Tester.Print "MS BusWidth/Speed Fail"
'                rv4 = 3
'            End If
'        End If
'        Call NewLabelMenu(4, "MS2 Bus Width/Speed", rv4, rv3)
'    End If
End If
Tester.Print rv4, " \\MSpro_2 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'

AU6479BLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'CF
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'SD0
        TestResult = "Bin3"
    ElseIf rv3 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv4 <> 1 Then        'MS2
        TestResult = "Bin5"
    ElseIf rv5 <> 1 Then        'MS1
        TestResult = "Bin5"
    ElseIf rv6 <> 1 Then        'SD1
        TestResult = "Bin3"
    ElseIf rv6 * rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub
Public Sub AU6479BLF23TestSub()

'2012/5/3 for S/B: AU6479-GBL 100LQ SOCKET V0.90
'2012/6/14 for B62 version
'2012/6/29 for S/B: AU6479-GBL 100LQ 4LUN SOCKET

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

'Call PowerSet2(0, "3.3", "0.7", 1, "3.3", "0.7", 1)

Tester.Print "AU6479BL : Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'CF  (Lun0)
rv2 = 0     'SD0 (Lun1)
rv3 = 0     'MS1 (Lun2)
rv4 = 0     'XD  (Lun3)
'rv5 = 0     'SD1 (Lun4)
'rv6 = 0     'MS2 (Lun4)



Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '            CF ¡BSD0 ¡BMS1 ¡BXD  ¡BSD1                       CF ¡BSD0 ¡BMS1 ¡BXD  ¡BMS2(Can't R/W)
                'Condition1(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)         Condition2(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)
'8:ENA    ---               0                                           0
'7:HID    ---               1                                           1
'6:M2INS  ---               1                                           0
'5:SD1CDN ---               0                                           1

'4:MSINS  ---               0                                           0
'3:XDCDN  ---               0                                           0
'2:CFCDN  ---               0                                           0
'1:SD0CDN ---               0                                           0

'                         0x60                                        0x50

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H60)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'If CardResult <> 0 Then
'    MsgBox "Read light off fail"
'    End
'End If
'
'Call MsecDelay(0.2)     'power on time
ChipString = "058f"
                
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOFF
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479BLFResult
End If



Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  CF Card test Lun0
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "CF", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
    Call NewLabelMenu(1, "CF_64K", rv1, rv0)
End If
Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  SD Card test Lun1
'===============================================

rv2 = CBWTest_New_PipeReady(1, rv1, ChipString)
If rv2 <> 1 Then
    Call NewLabelMenu(2, "SD", rv2, rv1)
Else
    rv2 = CBWTest_New_128_Sector_PipeReady(1, rv2)  ' write
    Call NewLabelMenu(2, "SD_64K", rv2, rv1)
    
    If rv2 = 1 Then
        rv2 = Read_Speed2ReadData(LBA, 1, 64)
        If rv2 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv2 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(2, "SD Bus Speed/Width", rv2, rv1)
    End If
    
End If

Tester.Print rv2, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'=======================================================================================
'       SDHC R / W
'=======================================================================================
        
'Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
'OpenPipe
'rv1 = ReInitial(0)
'Call MsecDelay(0.02)
'rv1 = AU6435ForceSDHC(rv0)
'ClosePipe
'
'If rv1 = 1 Then
'    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
'End If
'
'
'If rv1 = 1 Then
'    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
'    If rv1 <> 1 Then
'        rv1 = 2
'        Tester.Print "SD2.0 Mode Fail"
'    End If
'End If
'
'ClosePipe
'
'Call LabelMenu(1, rv1, rv0)
'
'Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
'===============================================
'  MSPro MS1 Card test Lun2
'================================================
rv3 = CBWTest_New(2, rv2, ChipString)
If rv3 <> 1 Then
    Call NewLabelMenu(3, "MS", rv3, rv2)
Else
    rv3 = CBWTest_New_128_Sector_PipeReady(2, rv3)  ' write
    Call NewLabelMenu(3, "MS_64K", rv3, rv2)

    If rv3 = 1 Then
        rv3 = Read_Speed2ReadData(LBA, 2, 64)
        If rv3 = 1 Then
            If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                Tester.Print "MS is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS BusWidth/Speed Fail"
                rv3 = 3
            End If
        End If
        Call NewLabelMenu(3, "MS Bus Width/Speed", rv3, rv2)
    End If
End If
Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
                  
'===============================================
'  XD Card test Lun3
'================================================
rv4 = CBWTest_New_PipeReady(3, rv3, ChipString)
If rv4 <> 1 Then
    Call NewLabelMenu(4, "XD", rv4, rv3)
Else
    rv4 = CBWTest_New_128_Sector_PipeReady(3, rv4)  ' write
    Call NewLabelMenu(4, "XD_64", rv3, rv2)
End If

Tester.Print rv4, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  SD1 Card test Lun4
'===============================================
rv5 = 1
'rv5 = CBWTest_New_PipeReady(4, rv4, ChipString)
'If rv5 <> 1 Then
'    Call NewLabelMenu(5, "SD1", rv5, rv4)
'Else
'    rv5 = CBWTest_New_128_Sector_PipeReady(4, rv5)   ' write
'    Call NewLabelMenu(5, "SD1_64K", rv5, rv4)
'
'    If rv5 = 1 Then
'        rv5 = Read_Speed2ReadData(LBA, 4, 64)
'        If rv5 = 1 Then
'            If (ReadData(20) = &H6) Then
'                Tester.Print "SD is 4 Bit, Speed 48 MHz"
'            Else
'                Tester.Print "SD BusWidth/Speed Fail"
'                rv5 = 3
'            End If
'        Else
'            Tester.Print "SD Bus Speed/Width Fail"
'        End If
'        Call NewLabelMenu(5, "SD1 Bus Width/Speed", rv5, rv4)
'    End If
'End If
'
'Tester.Print rv5, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
'
'
'ClosePipe
'If rv5 <> 1 Then
'    GoTo AU6479BLFResult
'End If
'
''            CF ¡BSD0 ¡BMS1 ¡BXD  ¡BMS2(Can't not R/W)
''Condition2(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)
'CardResult = DO_WritePort(card, Channel_P1A, &H50)
'Call MsecDelay(0.08)
'OpenPipe
'Call MsecDelay(0.04)
'rv5 = ReInitial(1)
'If rv5 = 1 Then
'    rv5 = ReInitial(4)
'End If
'ClosePipe
'
'If rv5 <> 1 Then
'    Call NewLabelMenu(5, "ReNew", rv5, rv4)
'    GoTo AU6479BLFResult
'Else
'    Call MsecDelay(0.2)
'End If

'===============================================
'  MSPro MS2 Card test Lun4
'================================================
'OpenPipe
'rv6 = RequestSense(4)
'ClosePipe
'If rv6 <> 1 Then
'    Call NewLabelMenu(6, "MS2", rv6, rv5)
'Else
'    rv4 = CBWTest_New_128_Sector_PipeReady(4, rv4)  ' write
'    Call NewLabelMenu(4, "MS2_64k", rv4, rv3)
'
'    If rv4 = 1 Then
'        rv4 = Read_Speed2ReadData(LBA, 4, 64)
'        If rv4 = 1 Then
'            If (ReadData(25) = &H2) And (ReadData(26) = &H70) And (ReadData(27) = &H1) Then
'                Tester.Print "MS is 4 Bit, Speed 40 MHz"
'            Else
'                Tester.Print "MS BusWidth/Speed Fail"
'                rv4 = 3
'            End If
'        End If
'        Call NewLabelMenu(4, "MS2 Bus Width/Speed", rv4, rv3)
'    End If
'End If
'Tester.Print rv4, " \\MSpro_2 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'

AU6479BLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'CF
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'SD0
        TestResult = "Bin3"
    ElseIf rv3 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv4 <> 1 Then        'MS2
        TestResult = "Bin5"
    'ElseIf rv5 <> 1 Then        'MS1
    '    TestResult = "Bin5"
    'ElseIf rv6 <> 1 Then        'SD1
    '    TestResult = "Bin3"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub

Public Sub AU6479OLF23TestSub()

'This code copy from AU6479BLF23
'2013/3/11 for S/B: AU6479-GOL 48LQ SOCKET

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

Tester.Print "AU6479OL : Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'SD0 (Lun0)
rv2 = 0     'SD1 (Lun1)
rv3 = 0     'MS1 (Lun0)
rv4 = 0     'MS2 (Lun1)
rv5 = 0     'XD  (Lun0)


Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '            SD0 ¡BSD1                       MS ¡BM2                        XD
                'Condition1(Lun0¡BLun1)         Condition2(Lun0¡BLun1)         Condition3(Lun0¡BLun1)
'8:ENA    ---               0                              0                                 0
'7:       ---               1                              1                                 1
'6:M2INS  ---               1                              0                                 1
'5:SD1CDN ---               0                              1                                 1

'4:MSINS  ---               1                              0                                 1
'3:XDCDN  ---               1                              1                                 0
'2:       ---               1                              1                                 1
'1:SD0CDN ---               0                              1                                 1

'                         0x6E                           0x57                              0x7B

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H6E)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'If CardResult <> 0 Then
'    MsgBox "Read light off fail"
'    End
'End If
'
'Call MsecDelay(0.2)     'power on time
ChipString = "pid_6459"
                
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOFF
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479OLFResult
End If

Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  SD0 Card test Lun1
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "SDHC", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
    Call NewLabelMenu(1, "SDHC_64K", rv1, rv0)
    
    If rv1 = 1 Then
        rv1 = Read_Speed2ReadData(LBA, 0, 64)
        If rv1 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv1 = 3
            End If
        Else
            Tester.Print "Read SDHC Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(1, "SD Bus Speed/Width", rv1, rv0)
    End If
    
End If

Tester.Print rv1, " \\SDHC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  SD1 Card test Lun1
'===============================================

rv2 = CBWTest_New_PipeReady(1, rv1, ChipString)
If rv2 <> 1 Then
    Call NewLabelMenu(2, "SD", rv2, rv1)
Else
    rv2 = CBWTest_New_128_Sector_PipeReady(1, rv2)  ' write
    Call NewLabelMenu(2, "SD_64K", rv2, rv1)
    
'    If rv2 = 1 Then
'        rv2 = Read_Speed2ReadData(LBA, 0, 64)
'        If rv2 = 1 Then
'            If (ReadData(20) = &H6) Then
'                Tester.Print "SD is 4 Bit, Speed 48 MHz"
'            Else
'                Tester.Print "SD BusWidth/Speed Fail"
'                rv2 = 3
'            End If
'        Else
'            Tester.Print "Read SD Bus Speed/Width Fail"
'        End If
'        Call NewLabelMenu(2, "SD Bus Speed/Width", rv2, rv1)
'    End If
End If

Tester.Print rv2, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'ClosePipe
If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H57)
    Call MsecDelay(0.04)
    'OpenPipe
    rv3 = ReInitial(0)
    If rv3 = 1 Then
        rv3 = ReInitial(1)
    End If
End If

If rv3 <> 1 Then
    Call NewLabelMenu(3, "ReNew", rv3, rv2)
    ClosePipe
    GoTo AU6479OLFResult
End If
                  
'===============================================
'  MSPro MS1 Card test Lun0
'================================================
rv3 = CBWTest_New_PipeReady(0, rv2, ChipString)
If rv3 <> 1 Then
    Call NewLabelMenu(3, "MS1", rv3, rv2)
Else
    rv3 = CBWTest_New_128_Sector_PipeReady(0, rv3)  ' write
    Call NewLabelMenu(3, "MS_64K", rv3, rv2)

    If rv3 = 1 Then
        rv3 = Read_Speed2ReadData(LBA, 0, 64)
        If rv3 = 1 Then
            If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                Tester.Print "MS is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS BusWidth/Speed Fail"
                rv3 = 3
            End If
        End If
        Call NewLabelMenu(3, "MS Bus Width/Speed", rv3, rv2)
    End If
End If
Tester.Print rv3, " \\MS1 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  MSPro MS2 Card test Lun1
'================================================
rv4 = CBWTest_New_PipeReady(1, rv3, ChipString)
If rv4 <> 1 Then
    Call NewLabelMenu(4, "MS2", rv4, rv3)
Else
    rv4 = CBWTest_New_128_Sector_PipeReady(0, rv4)  ' write
    Call NewLabelMenu(4, "MS_64K", rv4, rv3)

    If rv4 = 1 Then
        rv4 = Read_Speed2ReadData(LBA, 1, 64)
        If rv3 = 1 Then
            If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                Tester.Print "MS2 is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS2 BusWidth/Speed Fail"
                rv4 = 3
            End If
        End If
        Call NewLabelMenu(4, "MS2 Bus Width/Speed", rv4, rv3)
    End If
End If
Tester.Print rv4, " \\MS2 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


ClosePipe
If rv4 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H7B)
    Call MsecDelay(0.04)
    OpenPipe
    Call MsecDelay(0.1)
    rv5 = ReInitial(0)
    If rv5 = 1 Then
        rv5 = ReInitial(1)
    End If
End If

If rv5 <> 1 Then
    Call NewLabelMenu(5, "ReNew", rv5, rv4)
    ClosePipe
    GoTo AU6479OLFResult
End If

'===============================================
'  XD Card test Lun0
'================================================
rv5 = CBWTest_New_PipeReady(0, rv4, ChipString)
If rv5 <> 1 Then
    Call NewLabelMenu(5, "XD", rv4, rv3)
Else
    rv5 = CBWTest_New_128_Sector_PipeReady(0, rv5)  ' write
    Call NewLabelMenu(5, "XD_64", rv3, rv2)
End If
ClosePipe

Tester.Print rv5, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


AU6479OLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'SD0 (SDHC)
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'SD1 (SD)
        TestResult = "Bin3"
    ElseIf rv3 <> 1 Then        'MS1
        TestResult = "Bin5"
    ElseIf rv4 <> 1 Then        'MS2
        TestResult = "Bin5"
    ElseIf rv5 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub

Public Sub AU6479OLT10TestSub()

'This code copy from AU6479BLF23
'2013/3/11 for S/B: AU6479-GOL 48LQ SOCKET
'2013/9/4 purpose to skip M2 test item only for C63 version

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

Tester.Print "AU6479OL : Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'SD0 (Lun0)
rv2 = 0     'SD1 (Lun1)
rv3 = 0     'MS1 (Lun0)
rv4 = 0     'MS2 (Lun1)
rv5 = 0     'XD  (Lun0)


Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '            SD0 ¡BSD1                       MS ¡BM2                        XD
                'Condition1(Lun0¡BLun1)         Condition2(Lun0¡BLun1)         Condition3(Lun0¡BLun1)
'8:ENA    ---               0                              0                                 0
'7:       ---               1                              1                                 1
'6:M2INS  ---               1                              0                                 1
'5:SD1CDN ---               0                              1                                 1

'4:MSINS  ---               1                              0                                 1
'3:XDCDN  ---               1                              1                                 0
'2:       ---               1                              1                                 1
'1:SD0CDN ---               0                              1                                 1

'                         0x6E                           0x57                              0x7B

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H6E)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'If CardResult <> 0 Then
'    MsgBox "Read light off fail"
'    End
'End If
'
'Call MsecDelay(0.2)     'power on time
ChipString = "pid_6459"
                
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOFF
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479OLFResult
End If

Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  SD0 Card test Lun1
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "SDHC", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
    Call NewLabelMenu(1, "SDHC_64K", rv1, rv0)
    
    If rv1 = 1 Then
        rv1 = Read_Speed2ReadData(LBA, 0, 64)
        If rv1 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv1 = 3
            End If
        Else
            Tester.Print "Read SDHC Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(1, "SD Bus Speed/Width", rv1, rv0)
    End If
    
End If

Tester.Print rv1, " \\SDHC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  SD1 Card test Lun1
'===============================================

rv2 = CBWTest_New_PipeReady(1, rv1, ChipString)
If rv2 <> 1 Then
    Call NewLabelMenu(2, "SD", rv2, rv1)
Else
    rv2 = CBWTest_New_128_Sector_PipeReady(1, rv2)  ' write
    Call NewLabelMenu(2, "SD_64K", rv2, rv1)
    
'    If rv2 = 1 Then
'        rv2 = Read_Speed2ReadData(LBA, 0, 64)
'        If rv2 = 1 Then
'            If (ReadData(20) = &H6) Then
'                Tester.Print "SD is 4 Bit, Speed 48 MHz"
'            Else
'                Tester.Print "SD BusWidth/Speed Fail"
'                rv2 = 3
'            End If
'        Else
'            Tester.Print "Read SD Bus Speed/Width Fail"
'        End If
'        Call NewLabelMenu(2, "SD Bus Speed/Width", rv2, rv1)
'    End If
End If

Tester.Print rv2, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'ClosePipe
If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H57)
    Call MsecDelay(0.04)
    'OpenPipe
    rv3 = ReInitial(0)
    If rv3 = 1 Then
        rv3 = ReInitial(1)
    End If
End If

If rv3 <> 1 Then
    Call NewLabelMenu(3, "ReNew", rv3, rv2)
    ClosePipe
    GoTo AU6479OLFResult
End If
                  
'===============================================
'  MSPro MS1 Card test Lun0
'================================================
rv3 = CBWTest_New_PipeReady(0, rv2, ChipString)
If rv3 <> 1 Then
    Call NewLabelMenu(3, "MS1", rv3, rv2)
Else
    rv3 = CBWTest_New_128_Sector_PipeReady(0, rv3)  ' write
    Call NewLabelMenu(3, "MS_64K", rv3, rv2)

    If rv3 = 1 Then
        rv3 = Read_Speed2ReadData(LBA, 0, 64)
        If rv3 = 1 Then
            If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                Tester.Print "MS is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS BusWidth/Speed Fail"
                rv3 = 3
            End If
        End If
        Call NewLabelMenu(3, "MS Bus Width/Speed", rv3, rv2)
    End If
End If
Tester.Print rv3, " \\MS1 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  MSPro MS2 Card test Lun1
'================================================
'rv4 = CBWTest_New_PipeReady(1, rv3, ChipString)
'If rv4 <> 1 Then
'    Call NewLabelMenu(4, "MS2", rv4, rv3)
'Else
'    rv4 = CBWTest_New_128_Sector_PipeReady(0, rv4)  ' write
'    Call NewLabelMenu(4, "MS_64K", rv4, rv3)
'
'    If rv4 = 1 Then
'        rv4 = Read_Speed2ReadData(LBA, 1, 64)
'        If rv3 = 1 Then
'            If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
'                Tester.Print "MS2 is 4 Bit, Speed 40 MHz"
'            Else
'                Tester.Print "MS2 BusWidth/Speed Fail"
'                rv4 = 3
'            End If
'        End If
'        Call NewLabelMenu(4, "MS2 Bus Width/Speed", rv4, rv3)
'    End If
'End If
'Tester.Print rv4, " \\MS2 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

rv4 = 1

ClosePipe
If rv4 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H7B)
    Call MsecDelay(0.04)
    OpenPipe
    Call MsecDelay(0.1)
    rv5 = ReInitial(0)
    If rv5 = 1 Then
        rv5 = ReInitial(1)
    End If
End If

If rv5 <> 1 Then
    Call NewLabelMenu(5, "ReNew", rv5, rv4)
    ClosePipe
    GoTo AU6479OLFResult
End If

'===============================================
'  XD Card test Lun0
'================================================
rv5 = CBWTest_New_PipeReady(0, rv4, ChipString)
If rv5 <> 1 Then
    Call NewLabelMenu(5, "XD", rv4, rv3)
Else
    rv5 = CBWTest_New_128_Sector_PipeReady(0, rv5)  ' write
    Call NewLabelMenu(5, "XD_64", rv3, rv2)
End If
ClosePipe

Tester.Print rv5, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


AU6479OLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'SD0 (SDHC)
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'SD1 (SD)
        TestResult = "Bin3"
    ElseIf rv3 <> 1 Then        'MS1
        TestResult = "Bin5"
    ElseIf rv4 <> 1 Then        'MS2
        TestResult = "Bin5"
    ElseIf rv5 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub

Public Sub AU6479OLF03TestSub()

'This code copy from AU6479BLF23
'2013/3/11 for S/B: AU6479-GOL 48LQ SOCKET

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer
Dim HV_Done_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String


If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If


Routine_Label:


If Not HV_Done_Flag Then
    Call PowerSet2(0, "5.3", "0.5", 1, "5.3", "0.5", 1)
    Call MsecDelay(0.2)
    Tester.Print "AU6479FL : 5.3V Begin Test ..."
    SetSiteStatus (RunHV)
Else
    Call PowerSet2(0, "4.7", "0.5", 1, "4.7", "0.5", 1)
    Call MsecDelay(0.2)
    Tester.Print vbCrLf & "AU6479FL : 4.7V Begin Test ..."
    SetSiteStatus (RunLV)
End If

DetectCount = 0

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'SD0 (Lun0)
rv2 = 0     'SD1 (Lun1)
rv3 = 0     'MS1 (Lun0)
rv4 = 0     'MS2 (Lun1)
rv5 = 0     'XD  (Lun0)


Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '            SD0 ¡BSD1                       MS ¡BM2                        XD
                'Condition1(Lun0¡BLun1)         Condition2(Lun0¡BLun1)         Condition3(Lun0¡BLun1)
'8:ENA    ---               0                              0                                 0
'7:       ---               1                              1                                 1
'6:M2INS  ---               1                              0                                 1
'5:SD1CDN ---               0                              1                                 1

'4:MSINS  ---               1                              0                                 1
'3:XDCDN  ---               1                              1                                 0
'2:       ---               1                              1                                 1
'1:SD0CDN ---               0                              1                                 1

'                         0x6E                           0x57                              0x7B

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H6E)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'If CardResult <> 0 Then
'    MsgBox "Read light off fail"
'    End
'End If
'
'Call MsecDelay(0.2)     'power on time
ChipString = "pid_6459"
                
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOFF
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479OLFResult
End If

Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  SD0 Card test Lun1
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "SDHC", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
    Call NewLabelMenu(1, "SDHC_64K", rv1, rv0)
    
    If rv1 = 1 Then
        rv1 = Read_Speed2ReadData(LBA, 0, 64)
        If rv1 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv1 = 3
            End If
        Else
            Tester.Print "Read SDHC Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(1, "SD Bus Speed/Width", rv1, rv0)
    End If
    
End If

Tester.Print rv1, " \\SDHC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  SD1 Card test Lun1
'===============================================

rv2 = CBWTest_New_PipeReady(1, rv1, ChipString)
If rv2 <> 1 Then
    Call NewLabelMenu(2, "SD", rv2, rv1)
Else
    rv2 = CBWTest_New_128_Sector_PipeReady(1, rv2)  ' write
    Call NewLabelMenu(2, "SD_64K", rv2, rv1)
    
'    If rv2 = 1 Then
'        rv2 = Read_Speed2ReadData(LBA, 0, 64)
'        If rv2 = 1 Then
'            If (ReadData(20) = &H6) Then
'                Tester.Print "SD is 4 Bit, Speed 48 MHz"
'            Else
'                Tester.Print "SD BusWidth/Speed Fail"
'                rv2 = 3
'            End If
'        Else
'            Tester.Print "Read SD Bus Speed/Width Fail"
'        End If
'        Call NewLabelMenu(2, "SD Bus Speed/Width", rv2, rv1)
'    End If
End If

Tester.Print rv2, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'ClosePipe
If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H57)
    Call MsecDelay(0.04)
    'OpenPipe
    rv3 = ReInitial(0)
    If rv3 = 1 Then
        rv3 = ReInitial(1)
    End If
End If

If rv3 <> 1 Then
    Call NewLabelMenu(3, "ReNew", rv3, rv2)
    ClosePipe
    GoTo AU6479OLFResult
End If
                  
'===============================================
'  MSPro MS1 Card test Lun0
'================================================
rv3 = CBWTest_New_PipeReady(0, rv2, ChipString)
If rv3 <> 1 Then
    Call NewLabelMenu(3, "MS1", rv3, rv2)
Else
    rv3 = CBWTest_New_128_Sector_PipeReady(0, rv3)  ' write
    Call NewLabelMenu(3, "MS_64K", rv3, rv2)

    If rv3 = 1 Then
        rv3 = Read_Speed2ReadData(LBA, 0, 64)
        If rv3 = 1 Then
            If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                Tester.Print "MS is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS BusWidth/Speed Fail"
                rv3 = 3
            End If
        End If
        Call NewLabelMenu(3, "MS Bus Width/Speed", rv3, rv2)
    End If
End If
Tester.Print rv3, " \\MS1 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  MSPro MS2 Card test Lun1
'================================================
rv4 = CBWTest_New_PipeReady(1, rv3, ChipString)
If rv4 <> 1 Then
    Call NewLabelMenu(4, "MS2", rv4, rv3)
Else
    rv4 = CBWTest_New_128_Sector_PipeReady(0, rv4)  ' write
    Call NewLabelMenu(4, "MS_64K", rv4, rv3)

    If rv4 = 1 Then
        rv4 = Read_Speed2ReadData(LBA, 1, 64)
        If rv3 = 1 Then
            If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                Tester.Print "MS2 is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS2 BusWidth/Speed Fail"
                rv4 = 3
            End If
        End If
        Call NewLabelMenu(4, "MS2 Bus Width/Speed", rv4, rv3)
    End If
End If
Tester.Print rv4, " \\MS2 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

ClosePipe
If rv4 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H7B)
    Call MsecDelay(0.04)
    OpenPipe
    Call MsecDelay(0.1)
    rv5 = ReInitial(0)
    If rv5 = 1 Then
        rv5 = ReInitial(1)
    End If
End If

If rv5 <> 1 Then
    Call NewLabelMenu(5, "ReNew", rv5, rv4)
    GoTo AU6479OLFResult
End If

'===============================================
'  XD Card test Lun0
'================================================
rv5 = CBWTest_New_PipeReady(0, rv4, ChipString)
If rv5 <> 1 Then
    Call NewLabelMenu(5, "XD", rv4, rv3)
Else
    rv5 = CBWTest_New_128_Sector_PipeReady(0, rv5)  ' write
    Call NewLabelMenu(5, "XD_64", rv3, rv2)
End If
ClosePipe

Tester.Print rv5, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


AU6479OLFResult:
    
    ClosePipe
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    If Not HV_Done_Flag Then
        SetSiteStatus (HVDone)
        Call WaitAnotherSiteDone(HVDone, 4#)
    Else
        SetSiteStatus (LVDone)
        Call WaitAnotherSiteDone(LVDone, 4#)
    End If
    Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
    Call MsecDelay(0.2)
    WaitDevOFF (ChipString)
    SetSiteStatus (SiteUnknow)
    
    If HV_Done_Flag = False Then
        If rv0 <> 1 Then
            HV_Result = "Bin2"
            Tester.Print "HV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 <> 1 Then
            HV_Result = "Fail"
            Tester.Print "HV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 1 Then
            HV_Result = "PASS"
            Tester.Print "HV PASS"
        End If
        
        HV_Done_Flag = True
        GoTo Routine_Label
    Else
        If rv0 <> 1 Then
            LV_Result = "Bin2"
            Tester.Print "LV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 <> 1 Then
            LV_Result = "Fail"
            Tester.Print "LV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 1 Then
            LV_Result = "PASS"
            Tester.Print "LV PASS"
        End If
        
    End If
            
            
    If (HV_Result = "Bin2") And (LV_Result = "Bin2") Then
        TestResult = "Bin2"
    ElseIf (HV_Result <> "PASS") And (LV_Result = "PASS") Then
        TestResult = "Bin3"
    ElseIf (HV_Result = "PASS") And (LV_Result <> "PASS") Then
        TestResult = "Bin4"
    ElseIf (HV_Result <> "PASS") And (LV_Result <> "PASS") Then
        TestResult = "Bin5"
    ElseIf (HV_Result = "PASS") And (LV_Result = "PASS") Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
           
           
'    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
'    WaitDevOFF (ChipString)
'
'    If rv0 <> 1 Then            'Enum
'        UnknowDeviceFail = UnknowDeviceFail + 1
'        TestResult = "Bin2"
'    ElseIf rv1 <> 1 Then        'SD0 (SDHC)
'        TestResult = "Bin3"
'    ElseIf rv2 <> 1 Then        'SD1 (SD)
'        TestResult = "Bin3"
'    ElseIf rv3 <> 1 Then        'MS1
'        TestResult = "Bin5"
'    ElseIf rv4 <> 1 Then        'MS2
'        TestResult = "Bin5"
'    ElseIf rv5 <> 1 Then        'XD
'        TestResult = "Bin4"
'    ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
'        TestResult = "PASS"
'    Else
'        TestResult = "Bin2"
'    End If
    
End Sub

Public Sub AU6479KLF24TestSub()

'2012/10/15 for S/B: AU6479-GKL 80LQ SOCKET V1.00
'2012/10/22 C63 version: pid= "6362"

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

Tester.Print "AU6479KL : Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'CF  (Lun0)
rv2 = 0     'SD0 (Lun1)
rv3 = 0     'XD  (Lun1)
rv4 = 0     'MS  (Lun1)

Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)

'If (ChipName <> "AU6479KLF23") Then
'    Call PowerSet2(0, "5.0", "0.5", 1, "5.0", "0.5", 1)
'End If

                '  CF ¡BSD0       CF ¡BXD        CF ¡BMS
                '(Lun0¡BLun1)   (Lun0¡BLun1)   (Lun0¡BLun1)
'8:ENA    ---         0              0              0
'7:NC     ---         1              1              1
'6:MS2INS ---         1              1              1
'5:SD1CDN ---         1              1              1

'4:MSINS  ---         1              1              0
'3:XDCDN  ---         1              0              1
'2:CFCDN  ---         0              0              0
'1:SD0CDN ---         0              1              1

'                    0x7C           0x79           0x75

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H7F)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'If CardResult <> 0 Then
'    MsgBox "Read light off fail"
'    End
'End If
'
'Call MsecDelay(0.2)     'power on time
If ChipName = "AU6479KLF23" Then
    ChipString = "6361"
Else
    ChipString = "6362"
End If
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)

CardResult = DO_WritePort(card, Channel_P1A, &H7C)
Call MsecDelay(0.2)

CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOFF
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479BLFResult
End If



Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  CF Card test Lun0
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "CF", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
    Call NewLabelMenu(1, "CF_64K", rv1, rv0)
End If
Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  SD Card test Lun1
'===============================================

rv2 = CBWTest_New_PipeReady(1, rv1, ChipString)
If rv2 <> 1 Then
    Call NewLabelMenu(2, "SD", rv2, rv1)
Else
    rv2 = CBWTest_New_128_Sector_PipeReady(1, rv2)  ' write
    Call NewLabelMenu(2, "SD_64K", rv2, rv1)
    
    If rv2 = 1 Then
        rv2 = Read_Speed2ReadData(LBA, 1, 64)
        If rv2 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv2 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(2, "SD Bus Speed/Width", rv2, rv1)
    End If
    
End If

Tester.Print rv2, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'=======================================================================================
'       SDHC R / W
'=======================================================================================
        
'Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
'OpenPipe
'rv1 = ReInitial(0)
'Call MsecDelay(0.02)
'rv1 = AU6435ForceSDHC(rv0)
'ClosePipe
'
'If rv1 = 1 Then
'    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
'End If
'
'
'If rv1 = 1 Then
'    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
'    If rv1 <> 1 Then
'        rv1 = 2
'        Tester.Print "SD2.0 Mode Fail"
'    End If
'End If
'
'ClosePipe
'
'Call LabelMenu(1, rv1, rv0)
'
'Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
                  
'===============================================
'  XD Card test Lun1
'===============================================
CardResult = DO_WritePort(card, Channel_P1A, &H79)
Call MsecDelay(0.04)
rv3 = CBWTest_New_PipeReady(1, rv2, ChipString)
If rv3 <> 1 Then
    Call NewLabelMenu(3, "XD", rv3, rv2)
Else
    rv3 = CBWTest_New_128_Sector_PipeReady(1, rv3)  ' write
    Call NewLabelMenu(3, "XD_64", rv3, rv2)
End If

Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
                  
'===============================================
'  MSPro MS1 Card test Lun1
'================================================
CardResult = DO_WritePort(card, Channel_P1A, &H75)
Call MsecDelay(0.04)
rv4 = CBWTest_New(1, rv3, ChipString)
If rv4 <> 1 Then
    Call NewLabelMenu(4, "MS", rv4, rv3)
Else
    rv4 = CBWTest_New_128_Sector_PipeReady(1, rv4)  ' write
    Call NewLabelMenu(4, "MS_64K", rv4, rv3)

    If rv4 = 1 Then
        rv4 = Read_Speed2ReadData(LBA, 1, 64)
        If rv4 = 1 Then
            If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                Tester.Print "MS is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS BusWidth/Speed Fail"
                rv4 = 3
            End If
        End If
        Call NewLabelMenu(4, "MS Bus Width/Speed", rv4, rv3)
    End If
End If
Tester.Print rv4, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
                  


'===============================================
'  SD1 Card test Lun4
'===============================================
rv5 = 1
'rv5 = CBWTest_New_PipeReady(4, rv4, ChipString)
'If rv5 <> 1 Then
'    Call NewLabelMenu(5, "SD1", rv5, rv4)
'Else
'    rv5 = CBWTest_New_128_Sector_PipeReady(4, rv5)   ' write
'    Call NewLabelMenu(5, "SD1_64K", rv5, rv4)
'
'    If rv5 = 1 Then
'        rv5 = Read_Speed2ReadData(LBA, 4, 64)
'        If rv5 = 1 Then
'            If (ReadData(20) = &H6) Then
'                Tester.Print "SD is 4 Bit, Speed 48 MHz"
'            Else
'                Tester.Print "SD BusWidth/Speed Fail"
'                rv5 = 3
'            End If
'        Else
'            Tester.Print "SD Bus Speed/Width Fail"
'        End If
'        Call NewLabelMenu(5, "SD1 Bus Width/Speed", rv5, rv4)
'    End If
'End If
'
'Tester.Print rv5, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
'
'
'ClosePipe
'If rv5 <> 1 Then
'    GoTo AU6479BLFResult
'End If
'
''            CF ¡BSD0 ¡BMS1 ¡BXD  ¡BMS2(Can't not R/W)
''Condition2(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)
'CardResult = DO_WritePort(card, Channel_P1A, &H50)
'Call MsecDelay(0.08)
'OpenPipe
'Call MsecDelay(0.04)
'rv5 = ReInitial(1)
'If rv5 = 1 Then
'    rv5 = ReInitial(4)
'End If
'ClosePipe
'
'If rv5 <> 1 Then
'    Call NewLabelMenu(5, "ReNew", rv5, rv4)
'    GoTo AU6479BLFResult
'Else
'    Call MsecDelay(0.2)
'End If

'===============================================
'  MSPro MS2 Card test Lun4
'================================================
'OpenPipe
'rv6 = RequestSense(4)
'ClosePipe
'If rv6 <> 1 Then
'    Call NewLabelMenu(6, "MS2", rv6, rv5)
'Else
'    rv4 = CBWTest_New_128_Sector_PipeReady(4, rv4)  ' write
'    Call NewLabelMenu(4, "MS2_64k", rv4, rv3)
'
'    If rv4 = 1 Then
'        rv4 = Read_Speed2ReadData(LBA, 4, 64)
'        If rv4 = 1 Then
'            If (ReadData(25) = &H2) And (ReadData(26) = &H70) And (ReadData(27) = &H1) Then
'                Tester.Print "MS is 4 Bit, Speed 40 MHz"
'            Else
'                Tester.Print "MS BusWidth/Speed Fail"
'                rv4 = 3
'            End If
'        End If
'        Call NewLabelMenu(4, "MS2 Bus Width/Speed", rv4, rv3)
'    End If
'End If
'Tester.Print rv4, " \\MSpro_2 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'

AU6479BLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &HFC)   ' Close power
    Call MsecDelay(0.2)
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'CF
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'SD0
        TestResult = "Bin3"
    ElseIf rv3 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv4 <> 1 Then        'MS
        TestResult = "Bin5"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub

Public Sub AU6479TLF23TestSub()

'This code copy from AU6479KLF24TestSub
'2013/5/12 for S/B: AU6481GQL 80LQ SOCKET
'2013/5/12 C64 version: pid= "6362"

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

Tester.Print "AU6479TL : Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'CF  (Lun0)
rv2 = 0     'SD0 (Lun1)
rv3 = 0     'MS  (Lun2)
rv4 = 0     'XD  (Lun3)
rv5 = 0     'SMC (Lun3)

Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)

'If (ChipName <> "AU6479KLF23") Then
'    Call PowerSet2(0, "5.0", "0.5", 1, "5.0", "0.5", 1)
'End If

' Ctrl_1                    Ctrl_2

'8:ENA                      8:
'7:                         7:
'6:MSINS                    6:
'5:                         5:

'4:XDCDN                    4:
'3:SMCCDN '"H" active       3:
'2:CFCDN                    2:
'1:SDCDN                    1:GPON7


'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H7F)  '(connect CF, SD, MS, XD card)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
    MsgBox "Read light off fail"
    End
End If
'
'Call MsecDelay(0.2)     'power on time

ChipString = "6362"

             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)

CardResult = DO_WritePort(card, Channel_P1A, &H50)
Call MsecDelay(0.2)

CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Or (LightOff <> &HFF) Then
    Tester.Print "LightON="; LightOn
    Tester.Print "LightOFF="; LightOff
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479BLFResult
End If


If LBA = 1 Then
    LBA = 1000
Else
    LBA = LBA + 1
End If
Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  CF Card test Lun0
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "CF", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
    Call NewLabelMenu(1, "CF_64K", rv1, rv0)
End If
Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  SD Card test Lun1
'===============================================

rv2 = CBWTest_New_PipeReady(1, rv1, ChipString)
If rv2 <> 1 Then
    Call NewLabelMenu(2, "SD", rv2, rv1)
Else
    rv2 = CBWTest_New_128_Sector_PipeReady(1, rv2)  ' write
    Call NewLabelMenu(2, "SD_64K", rv2, rv1)
    
    If rv2 = 1 Then
        rv2 = Read_Speed2ReadData(LBA, 1, 64)
        If rv2 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv2 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(2, "SD Bus Speed/Width", rv2, rv1)
    End If
    
End If

Tester.Print rv2, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  MSPro MS2 Card test Lun2
'================================================
rv3 = CBWTest_New_PipeReady(2, rv2, ChipString)
If rv3 <> 1 Then
    Call NewLabelMenu(3, "MS", rv3, rv2)
Else
    rv3 = CBWTest_New_128_Sector_PipeReady(1, rv3)  ' write
    Call NewLabelMenu(3, "MS_64K", rv3, rv2)

    If rv3 = 1 Then
        rv3 = Read_Speed2ReadData(LBA, 1, 64)
        If rv3 = 1 Then
            If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                Tester.Print "MS is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS BusWidth/Speed Fail"
                rv3 = 3
            End If
        End If
        Call NewLabelMenu(3, "MS Bus Width/Speed", rv3, rv2)
    End If
End If
Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
         
                  
'===============================================
'  XD Card test Lun3
'===============================================
rv4 = CBWTest_New_PipeReady(3, rv3, ChipString)
If rv4 <> 1 Then
    Call NewLabelMenu(4, "XD", rv4, rv3)
Else
    rv4 = CBWTest_New_128_Sector_PipeReady(1, rv4)  ' write
    Call NewLabelMenu(4, "XD_64", rv4, rv3)
End If

Tester.Print rv4, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
                  
'===============================================
'  SMC Card test Lun3
'===============================================
'ClosePipe

If rv4 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H5C)
    Call MsecDelay(0.02)
    CardResult = DO_WritePort(card, Channel_P1A, &H54)
    Call MsecDelay(0.02)
    
    rv5 = ReInitial(3)
    rv5 = CBWTest_New_PipeReady(3, rv4, ChipString)
    Call NewLabelMenu(5, "SMC", rv5, rv4)

End If

Tester.Print rv5, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


AU6479BLFResult:

    ClosePipe
    CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
    Call MsecDelay(0.2)
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'CF
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'SD0
        TestResult = "Bin3"
    ElseIf rv3 <> 1 Then        'MS
        TestResult = "Bin5"
    ElseIf rv4 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv5 <> 1 Then        'SMC
        TestResult = "Bin4"
    ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub

Public Sub AU6479BFF23TestSub()

'This code copy from AU6479KLF24TestSub
'2013/7/22 for S/B: AU6479GBF 28QFN SOCKET


Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

Tester.Print "AU6479BF : Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'SD0  (Lun0)
rv2 = 0     'SD1  (Lun1)

Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)



' Ctrl_1                    Ctrl_2

'8:ENA                      8:
'7:                         7:
'6:                         6:
'5:                         5:

'4:                         4:
'3:                         3:
'2:SD1CDN                   2:
'1:SD0CDN                   1:


'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H0)  '(connect SD0, SD1 card)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

'Call MsecDelay(0.2)     'power on time

ChipString = "6459"

             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)

CardResult = DO_WritePort(card, Channel_P1A, &H0)
Call MsecDelay(0.2)


If LBA = 1 Then
    LBA = 1000
Else
    LBA = LBA + 1
End If
Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  SD Card test Lun0
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "SDHC", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
    Call NewLabelMenu(1, "SDHC_64K", rv1, rv0)
    
    If rv1 = 1 Then
        rv1 = Read_Speed2ReadData(LBA, 0, 64)
        If rv1 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv1 = 3
            End If
        Else
            Tester.Print "SDHC Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(1, "SDHC Bus Speed/Width", rv1, rv0)
    End If
    
End If

Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  SD Card test Lun1
'===============================================

rv2 = CBWTest_New_PipeReady(1, rv1, ChipString)
If rv2 <> 1 Then
    Call NewLabelMenu(2, "SD", rv2, rv1)
Else
    rv2 = CBWTest_New_128_Sector_PipeReady(1, rv1)  ' write
    Call NewLabelMenu(2, "SD_64K", rv2, rv1)
End If

Tester.Print rv2, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"



AU6479BLFResult:

    ClosePipe
    CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
    Call MsecDelay(0.2)
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'SD0
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'SD
        TestResult = "Bin4"
    ElseIf rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub

Public Sub AU6479CFF23TestSub()

'This code copy from AU6479BFF23TestSub
'2013/8/06 test MMC

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer
Dim ChipString As String

    OldChipName = ""

    If PCI7248InitFinish = 0 Then
          PCI7248Exist
    End If

    DetectCount = 0

    Tester.Print "AU6479CF : Begin Test ..."

    LBA = LBA + 1
             
    rv0 = 0     'Enum
    rv1 = 0     'SD0  (Lun0)
    rv2 = 0     'LED
    
    Tester.Label3.BackColor = RGB(255, 255, 255)
    Tester.Label4.BackColor = RGB(255, 255, 255)
    Tester.Label5.BackColor = RGB(255, 255, 255)
    Tester.Label6.BackColor = RGB(255, 255, 255)
    Tester.Label7.BackColor = RGB(255, 255, 255)
    Tester.Label8.BackColor = RGB(255, 255, 255)


    ' Ctrl_1                    Ctrl_2
    
    '8:ENA                      8:
    '7:                         7:
    '6:                         6:
    '5:                         5:
    
    '4:                         4:
    '3:                         3:
    '2:                         2:
    '1:SD0CDN                   1: must read low


    '=========================================
    '    POWER on
    '=========================================
    CardResult = DO_WritePort(card, Channel_P1A, &H0)  '(connect SD0, SD1 card)
    
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
    
    'Call MsecDelay(0.2)     'power on time
    
    ChipString = "6366"

             
    '===============================================
    '  Enum Device
    '===============================================
                        
    Call MsecDelay(0.2)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.2)
    
    Call NewLabelMenu(0, "WaitDevice", rv0, 1)
    
    CardResult = DO_WritePort(card, Channel_P1A, &H0)
    Call MsecDelay(0.2)
    
    If LBA = 1 Then
        LBA = 1000
    Else
        LBA = LBA + 1
    End If
    Tester.Print "LBA="; LBA

    '===============================================
    '  SD Card test Lun0
    '===============================================
    
    rv1 = CBWTest_New(0, rv0, ChipString)
    If rv1 <> 1 Then
        Call NewLabelMenu(1, "SDHC", rv1, rv0)
    Else
        rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
        'Call NewLabelMenu(1, "SDHC_64K", rv2, rv1)
        Call NewLabelMenu(1, "SDHC_64K", rv1, rv0)
        If rv1 = 1 Then
            rv1 = Read_Speed2ReadData(LBA, 0, 64)
            'Call NewLabelMenu(1, "SDHC Bus Speed/Width", rv1, rv0) 'no need to test speed
        End If
        
    End If

    Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If CardResult <> 0 Then
        MsgBox "Read light On fail"
        End
    End If

    Call MsecDelay(0.02)
    
    If ((LightOn And &H1) <> 0) Then
        Tester.Print "LightON="; LightOn
        UsbSpeedTestResult = GPO_FAIL
        rv2 = 3
        GoTo AU6479CFFResult
    Else
        rv2 = 1
    End If
    
AU6479CFFResult:

    ClosePipe
    CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
    Call MsecDelay(0.2)
    WaitDevOFF (ChipString)
                
        If rv0 <> 1 Or rv2 <> 1 Then           'Enum fail or LED fail
            UnknowDeviceFail = UnknowDeviceFail + 1
            TestResult = "Bin2"
        ElseIf rv1 <> 1 Then        'SD0
            TestResult = "Bin3"
        ElseIf rv1 * rv0 * rv2 = PASS Then
            TestResult = "PASS"
        Else
            TestResult = "Bin2"
        End If
    
End Sub

Public Sub AU6479DFF24TestSub()

'This code copy from AU6479BFF23TestSub
'2013/8/06 test MMC
'20140/8/29 this bonding must enum device finished first

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer
Dim ChipString As String

    OldChipName = ""

    If PCI7248InitFinish = 0 Then
          PCI7248Exist
    End If

    DetectCount = 0

    Tester.Print "AU6479DF : Begin Test ..."

    LBA = LBA + 1
             
    rv0 = 0     'Enum
    rv1 = 0     'SD0  (Lun0)
    rv2 = 0     'LED
    
    Tester.Label3.BackColor = RGB(255, 255, 255)
    Tester.Label4.BackColor = RGB(255, 255, 255)
    Tester.Label5.BackColor = RGB(255, 255, 255)
    Tester.Label6.BackColor = RGB(255, 255, 255)
    Tester.Label7.BackColor = RGB(255, 255, 255)
    Tester.Label8.BackColor = RGB(255, 255, 255)


    ' Ctrl_1                    Ctrl_2
    
    '8:ENA                      8:
    '7:                         7:
    '6:                         6:
    '5:                         5:
    
    '4:                         4:
    '3:                         3:
    '2:                         2:
    '1:SD0CDN                   1: must read low


    '=========================================
    '    POWER on
    '=========================================

    CardResult = DO_WritePort(card, Channel_P1A, &H7F)  '(connect SD0 card)
    
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
    
    'Call MsecDelay(0.2)     'power on time
    
    ChipString = "6366"

             
    '===============================================
    '  Enum Device
    '===============================================
                        
    Call MsecDelay(0.2)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.2)
    
    Call NewLabelMenu(0, "WaitDevice", rv0, 1)
    
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
    Call MsecDelay(0.2)
    
    If LBA = 1 Then
        LBA = 1000
    Else
        LBA = LBA + 1
    End If
    Tester.Print "LBA="; LBA

    '===============================================
    '  SD Card test Lun0
    '===============================================
    
    rv1 = CBWTest_New(0, rv0, ChipString)
    If rv1 <> 1 Then
        Call NewLabelMenu(1, "SDHC", rv1, rv0)
    Else
        rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
        'Call NewLabelMenu(1, "SDHC_64K", rv2, rv1)
        Call NewLabelMenu(1, "SDHC_64K", rv1, rv0)
        If rv1 = 1 Then
            rv1 = Read_Speed2ReadData(LBA, 0, 64)
            'Call NewLabelMenu(1, "SDHC Bus Speed/Width", rv1, rv0) 'no need to test speed
        End If
        
    End If

    Tester.Print rv1, " \\MMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If CardResult <> 0 Then
        MsgBox "Read light On fail"
        End
    End If

    Call MsecDelay(0.02)
    
    If ((LightOn And &H1) <> 0) Then
        Tester.Print "LightON="; LightOn
        UsbSpeedTestResult = GPO_FAIL
        rv2 = 3
        GoTo AU6479CFFResult
    Else
        rv2 = 1
    End If
    
AU6479CFFResult:

    ClosePipe
    CardResult = DO_WritePort(card, Channel_P1A, &HF0)   ' Close power
    Call MsecDelay(0.2)
    WaitDevOFF (ChipString)
                
        If rv0 <> 1 Or rv2 <> 1 Then           'Enum fail or LED fail
            UnknowDeviceFail = UnknowDeviceFail + 1
            TestResult = "Bin2"
        ElseIf rv1 <> 1 Then        'SD0
            TestResult = "Bin3"
        ElseIf rv1 * rv0 * rv2 = PASS Then
            TestResult = "PASS"
        Else
            TestResult = "Bin2"
        End If
    
End Sub


Public Sub AU6479ALF20TestSub()


'2012/9/9 for S/B: AU6479-AL 128LQ 5LUN SOCKET V1.02
'This copy from AU6479BLF23

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

Tester.Print "AU6479AL : Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'CF  (Lun0)
rv2 = 0     'SD0 (Lun1)
rv3 = 0     'MS1 (Lun2)
rv4 = 0     'XD  (Lun3)
rv5 = 0     'SD1 (Lun4)
rv6 = 0     'MS2 (Lun4)



Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '            CF ¡BSD0 ¡BMS1 ¡BXD  ¡BSD1                       CF ¡BSD0 ¡BMS1 ¡BXD  ¡BMS2
                'Condition1(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)         Condition2(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)
'8:ENA    ---               0                                           0
'7:NC     ---               1                                           1
'6:M2INS  ---               1                                           0
'5:SD1CDN ---               0                                           1

'4:MSINS  ---               0                                           0
'3:XDCDN  ---               0                                           0
'2:CFCDN  ---               0                                           0
'1:SD0CDN ---               0                                           0

'                         0x60                                        0x50

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H7F)
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
    MsgBox "Read light off fail"
    End
End If

Call MsecDelay(0.2)     'power on time
ChipString = "058f"
                
CardResult = DO_WritePort(card, Channel_P1A, &H60)
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.4)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Or ((LightOff And &H1) <> 1) Then
    Tester.Print "LightON="; LightOn
    Tester.Print "LightOFF="; LightOff
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479BLFResult
End If



Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  CF Card test Lun0
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "CF", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
    Call NewLabelMenu(1, "CF_64K", rv1, rv0)
End If
Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  SD Card test Lun1
'===============================================

rv2 = CBWTest_New_PipeReady(1, rv1, ChipString)
If rv2 <> 1 Then
    Call NewLabelMenu(2, "SD", rv2, rv1)
Else
    rv2 = CBWTest_New_128_Sector_PipeReady(1, rv2)  ' write
    Call NewLabelMenu(2, "SD_64K", rv2, rv1)
    
    If rv2 = 1 Then
        rv2 = Read_Speed2ReadData(LBA, 1, 64)
        If rv2 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv2 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(2, "SD Bus Speed/Width", rv2, rv1)
    End If
    
End If

Tester.Print rv2, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

                  
'===============================================
'  MSPro MS1 Card test Lun2
'===============================================
rv3 = CBWTest_New(2, rv2, ChipString)
If rv3 <> 1 Then
    Call NewLabelMenu(3, "MS", rv3, rv2)
Else
    rv3 = CBWTest_New_128_Sector_PipeReady(2, rv3)  ' write
    Call NewLabelMenu(3, "MS_64K", rv3, rv2)

    If rv3 = 1 Then
        rv3 = Read_Speed2ReadData(LBA, 2, 64)
        If rv3 = 1 Then
            If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                Tester.Print "MS is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS BusWidth/Speed Fail"
                rv3 = 3
            End If
        End If
        Call NewLabelMenu(3, "MS Bus Width/Speed", rv3, rv2)
    End If
End If
Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
                  
'===============================================
'  XD Card test Lun3
'===============================================
rv4 = CBWTest_New_PipeReady(3, rv3, ChipString)
If rv4 <> 1 Then
    Call NewLabelMenu(4, "XD", rv4, rv3)
Else
    rv4 = CBWTest_New_128_Sector_PipeReady(3, rv4)  ' write
    Call NewLabelMenu(4, "XD_64", rv3, rv2)
End If

Tester.Print rv4, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  SD1 Card test Lun4
'===============================================
rv5 = CBWTest_New_PipeReady(4, rv4, ChipString)
If rv5 <> 1 Then
    Call NewLabelMenu(5, "SD1", rv5, rv4)
Else
    rv5 = CBWTest_New_128_Sector_PipeReady(4, rv5)   ' write
    Call NewLabelMenu(5, "SD1_64K", rv5, rv4)

    If rv5 = 1 Then
        rv5 = Read_Speed2ReadData(LBA, 4, 64)
        If rv5 = 1 Then
            If (ReadData(20) = &H6) Then
                Tester.Print "SD is 4 Bit, Speed 48 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv5 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(5, "SD1 Bus Width/Speed", rv5, rv4)
    End If
End If

Tester.Print rv5, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


ClosePipe
If rv5 <> 1 Then
    GoTo AU6479BLFResult
End If


'            CF ¡BSD0 ¡BMS1 ¡BXD  ¡BMS2
'Condition2(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)
CardResult = DO_WritePort(card, Channel_P1A, &H50)
Call MsecDelay(0.08)
OpenPipe
Call MsecDelay(0.04)
rv5 = ReInitial(4)
If rv5 <> 1 Then
    rv5 = ReInitial(4)
End If
ClosePipe

If rv5 <> 1 Then
    Call NewLabelMenu(5, "ReNew", rv5, rv4)
    GoTo AU6479BLFResult
Else
    Call MsecDelay(0.2)
End If

'===============================================
'  MS2 Card test Lun4
'===============================================

rv6 = CBWTest_New(4, rv5, ChipString)   ' write
If rv6 <> 1 Then
    Call NewLabelMenu(3, "MS2_4k", rv6, rv5)
Else
    rv6 = CBWTest_New_128_Sector_PipeReady(4, rv6)   ' write
    Call NewLabelMenu(3, "MS2_64K", rv6, rv5)
    
    If rv6 = 1 Then
        rv6 = Read_Speed2ReadData(LBA, 4, 64)
        If rv6 = 1 Then
            If (ReadData(27) = &H2) And (ReadData(28) = &H70) And (ReadData(29) = &H1) Then
                Tester.Print "MS is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS BusWidth/Speed Fail"
                rv6 = 3
            End If
        End If
        Call NewLabelMenu(3, "MS2 Bus Width/Speed", rv6, rv5)
    End If
End If

Tester.Print rv6, " \\MSpro_2 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


AU6479BLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'CF
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'SD0
        TestResult = "Bin3"
    ElseIf rv3 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv4 <> 1 Then        'MS
        TestResult = "Bin5"
    ElseIf rv5 <> 1 Then        'SD1
        TestResult = "Bin3"
    ElseIf rv6 <> 1 Then        'MS2
        TestResult = "Bin5"
    ElseIf rv6 * rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub

Public Sub AU6479ALT10TestSub()


'2012/9/9 for S/B: AU6479-AL 128LQ 5LUN SOCKET V1.02
'This copy from AU6479ALF20
'2013/3/19: purpose to skip M2 test item only for C63 version

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

Tester.Print "AU6479AL : Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'CF  (Lun0)
rv2 = 0     'SD0 (Lun1)
rv3 = 0     'MS1 (Lun2)
rv4 = 0     'XD  (Lun3)
rv5 = 0     'SD1 (Lun4)
'rv6 = 0     'MS2 (Lun4)



Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '            CF ¡BSD0 ¡BMS1 ¡BXD  ¡BSD1                       CF ¡BSD0 ¡BMS1 ¡BXD  ¡BMS2
                'Condition1(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)         Condition2(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)
'8:ENA    ---               0                                           0
'7:NC     ---               1                                           1
'6:M2INS  ---               1                                           0
'5:SD1CDN ---               0                                           1

'4:MSINS  ---               0                                           0
'3:XDCDN  ---               0                                           0
'2:CFCDN  ---               0                                           0
'1:SD0CDN ---               0                                           0

'                         0x60                                        0x50

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H7F)
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
    MsgBox "Read light off fail"
    End
End If

Call MsecDelay(0.2)     'power on time
ChipString = "058f"
                
CardResult = DO_WritePort(card, Channel_P1A, &H60)
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.4)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Or ((LightOff And &H1) <> 1) Then
    Tester.Print "LightON="; LightOn
    Tester.Print "LightOFF="; LightOff
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479BLFResult
End If



Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  CF Card test Lun0
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "CF", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
    Call NewLabelMenu(1, "CF_64K", rv1, rv0)
End If
Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  SD Card test Lun1
'===============================================

rv2 = CBWTest_New_PipeReady(1, rv1, ChipString)
If rv2 <> 1 Then
    Call NewLabelMenu(2, "SD", rv2, rv1)
Else
    rv2 = CBWTest_New_128_Sector_PipeReady(1, rv2)  ' write
    Call NewLabelMenu(2, "SD_64K", rv2, rv1)
    
    If rv2 = 1 Then
        rv2 = Read_Speed2ReadData(LBA, 1, 64)
        If rv2 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv2 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(2, "SD Bus Speed/Width", rv2, rv1)
    End If
    
End If

Tester.Print rv2, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

                  
'===============================================
'  MSPro MS1 Card test Lun2
'===============================================
rv3 = CBWTest_New(2, rv2, ChipString)
If rv3 <> 1 Then
    Call NewLabelMenu(3, "MS", rv3, rv2)
Else
    rv3 = CBWTest_New_128_Sector_PipeReady(2, rv3)  ' write
    Call NewLabelMenu(3, "MS_64K", rv3, rv2)

    If rv3 = 1 Then
        rv3 = Read_Speed2ReadData(LBA, 2, 64)
        If rv3 = 1 Then
            If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                Tester.Print "MS is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS BusWidth/Speed Fail"
                rv3 = 3
            End If
        End If
        Call NewLabelMenu(3, "MS Bus Width/Speed", rv3, rv2)
    End If
End If
Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
                  
'===============================================
'  XD Card test Lun3
'===============================================
rv4 = CBWTest_New_PipeReady(3, rv3, ChipString)
If rv4 <> 1 Then
    Call NewLabelMenu(4, "XD", rv4, rv3)
Else
    rv4 = CBWTest_New_128_Sector_PipeReady(3, rv4)  ' write
    Call NewLabelMenu(4, "XD_64", rv3, rv2)
End If

Tester.Print rv4, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  SD1 Card test Lun4
'===============================================
rv5 = CBWTest_New_PipeReady(4, rv4, ChipString)
If rv5 <> 1 Then
    Call NewLabelMenu(5, "SD1", rv5, rv4)
Else
    rv5 = CBWTest_New_128_Sector_PipeReady(4, rv5)   ' write
    Call NewLabelMenu(5, "SD1_64K", rv5, rv4)

    If rv5 = 1 Then
        rv5 = Read_Speed2ReadData(LBA, 4, 64)
        If rv5 = 1 Then
            If (ReadData(20) = &H6) Then
                Tester.Print "SD is 4 Bit, Speed 48 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv5 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(5, "SD1 Bus Width/Speed", rv5, rv4)
    End If
End If

Tester.Print rv5, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


ClosePipe
'If rv5 <> 1 Then
'    GoTo AU6479BLFResult
'End If


'            CF ¡BSD0 ¡BMS1 ¡BXD  ¡BMS2
'Condition2(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)
'CardResult = DO_WritePort(card, Channel_P1A, &H50)
'Call MsecDelay(0.08)
'OpenPipe
'Call MsecDelay(0.04)
'rv5 = ReInitial(4)
'If rv5 <> 1 Then
'    rv5 = ReInitial(4)
'End If
'ClosePipe

'If rv5 <> 1 Then
'    Call NewLabelMenu(5, "ReNew", rv5, rv4)
'    GoTo AU6479BLFResult
'Else
'    Call MsecDelay(0.2)
'End If

'===============================================
'  MS2 Card test Lun4
'===============================================

'rv6 = CBWTest_New(4, rv5, ChipString)   ' write
'If rv6 <> 1 Then
'    Call NewLabelMenu(3, "MS2_4k", rv6, rv5)
'Else
'    rv6 = CBWTest_New_128_Sector_PipeReady(4, rv6)   ' write
'    Call NewLabelMenu(3, "MS2_64K", rv6, rv5)
'
'    If rv6 = 1 Then
'        rv6 = Read_Speed2ReadData(LBA, 4, 64)
'        If rv6 = 1 Then
'            If (ReadData(27) = &H2) And (ReadData(28) = &H70) And (ReadData(29) = &H1) Then
'                Tester.Print "MS is 4 Bit, Speed 40 MHz"
'            Else
'                Tester.Print "MS BusWidth/Speed Fail"
'                rv6 = 3
'            End If
'        End If
'        Call NewLabelMenu(3, "MS2 Bus Width/Speed", rv6, rv5)
'    End If
'End If
'
'Tester.Print rv6, " \\MSpro_2 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


AU6479BLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'CF
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'SD0
        TestResult = "Bin3"
    ElseIf rv3 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv4 <> 1 Then        'MS
        TestResult = "Bin5"
    ElseIf rv5 <> 1 Then        'SD1
        TestResult = "Bin3"
    ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub

Public Sub AU6479JLFE3TestSub()

'2012/5/3 for S/B: AU6479-GBL 100LQ SOCKET V0.90
'2012/6/14 for B62 version
'2012/6/29 for S/B: AU6479-GBL 100LQ 4LUN SOCKET
'2012/7/25 Enum 3 cycle purpose to sorting some LC unstable IC

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

Tester.Print "AU6479BL : Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'CF  (Lun0)
rv2 = 0     'SD0 (Lun1)
rv3 = 0     'MS1 (Lun2)
rv4 = 0     'XD  (Lun3)
'rv5 = 0     'SD1 (Lun4)
'rv6 = 0     'MS2 (Lun4)



Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '            CF ¡BSD0 ¡BMS1 ¡BXD  ¡BSD1                       CF ¡BSD0 ¡BMS1 ¡BXD  ¡BMS2(Can't not R/W)
                'Condition1(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)         Condition2(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)
'8:ENA    ---               0                                           0
'7:HID    ---               1                                           1
'6:M2INS  ---               1                                           0
'5:SD1CDN ---               0                                           1

'4:MSINS  ---               0                                           0
'3:XDCDN  ---               0                                           0
'2:CFCDN  ---               0                                           0
'1:SD0CDN ---               0                                           0

'                         0x60                                        0x50

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H60)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'If CardResult <> 0 Then
'    MsgBox "Read light off fail"
'    End
'End If
'
'Call MsecDelay(0.2)     'power on time
ChipString = "058f"
                
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString) 'Cycle: 1

If (rv0 = 1) Then
    CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
    rv0 = WaitDevOFF(ChipString)
    If (rv0 = 1) Then
        Call MsecDelay(0.2)
        CardResult = DO_WritePort(card, Channel_P1A, &H60)
        rv0 = WaitDevOn(ChipString) 'Cycle: 2
        If (rv0 = 1) Then
            Call MsecDelay(0.2)
            CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
            rv0 = WaitDevOFF(ChipString)
            If (rv0 = 1) Then
                Call MsecDelay(0.2)
                CardResult = DO_WritePort(card, Channel_P1A, &H60)
                rv0 = WaitDevOn(ChipString) 'Cycle: 3
            End If
        End If
    End If
End If
    
    
Call MsecDelay(0.2)

Call NewLabelMenu(0, "UnknowDevice", rv0, 1)


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOFF
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479BLFResult
End If



Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  CF Card test Lun0
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "CF", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
    Call NewLabelMenu(1, "CF_64K", rv1, rv0)
End If
Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  SD Card test Lun1
'===============================================

rv2 = CBWTest_New_PipeReady(1, rv1, ChipString)
If rv2 <> 1 Then
    Call NewLabelMenu(2, "SD", rv2, rv1)
Else
    rv2 = CBWTest_New_128_Sector_PipeReady(1, rv2)  ' write
    Call NewLabelMenu(2, "SD_64K", rv2, rv1)
    
    If rv2 = 1 Then
        rv2 = Read_Speed2ReadData(LBA, 1, 64)
        If rv2 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv2 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(2, "SD Bus Speed/Width", rv2, rv1)
    End If
    
End If

Tester.Print rv2, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'=======================================================================================
'       SDHC R / W
'=======================================================================================
        
'Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
'OpenPipe
'rv1 = ReInitial(0)
'Call MsecDelay(0.02)
'rv1 = AU6435ForceSDHC(rv0)
'ClosePipe
'
'If rv1 = 1 Then
'    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
'End If
'
'
'If rv1 = 1 Then
'    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
'    If rv1 <> 1 Then
'        rv1 = 2
'        Tester.Print "SD2.0 Mode Fail"
'    End If
'End If
'
'ClosePipe
'
'Call LabelMenu(1, rv1, rv0)
'
'Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
'===============================================
'  MSPro MS1 Card test Lun2
'================================================
rv3 = CBWTest_New(2, rv2, ChipString)
If rv3 <> 1 Then
    Call NewLabelMenu(3, "MS", rv3, rv2)
Else
    rv3 = CBWTest_New_128_Sector_PipeReady(2, rv3)  ' write
    Call NewLabelMenu(3, "MS_64K", rv3, rv2)

    If rv3 = 1 Then
        rv3 = Read_Speed2ReadData(LBA, 2, 64)
        If rv3 = 1 Then
            If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                Tester.Print "MS is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS BusWidth/Speed Fail"
                rv3 = 3
            End If
        End If
        Call NewLabelMenu(3, "MS Bus Width/Speed", rv3, rv2)
    End If
End If
Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
                  
'===============================================
'  XD Card test Lun3
'================================================
rv4 = CBWTest_New_PipeReady(3, rv3, ChipString)
If rv4 <> 1 Then
    Call NewLabelMenu(4, "XD", rv3, rv2)
Else
    rv4 = CBWTest_New_128_Sector_PipeReady(3, rv4)  ' write
    Call NewLabelMenu(4, "XD_64", rv3, rv2)
End If

Tester.Print rv4, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  SD1 Card test Lun4
'===============================================
rv5 = 1
'rv5 = CBWTest_New_PipeReady(4, rv4, ChipString)
'If rv5 <> 1 Then
'    Call NewLabelMenu(5, "SD1", rv5, rv4)
'Else
'    rv5 = CBWTest_New_128_Sector_PipeReady(4, rv5)   ' write
'    Call NewLabelMenu(5, "SD1_64K", rv5, rv4)
'
'    If rv5 = 1 Then
'        rv5 = Read_Speed2ReadData(LBA, 4, 64)
'        If rv5 = 1 Then
'            If (ReadData(20) = &H6) Then
'                Tester.Print "SD is 4 Bit, Speed 48 MHz"
'            Else
'                Tester.Print "SD BusWidth/Speed Fail"
'                rv5 = 3
'            End If
'        Else
'            Tester.Print "SD Bus Speed/Width Fail"
'        End If
'        Call NewLabelMenu(5, "SD1 Bus Width/Speed", rv5, rv4)
'    End If
'End If
'
'Tester.Print rv5, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
'
'
'ClosePipe
'If rv5 <> 1 Then
'    GoTo AU6479BLFResult
'End If
'
''            CF ¡BSD0 ¡BMS1 ¡BXD  ¡BMS2(Can't not R/W)
''Condition2(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)
'CardResult = DO_WritePort(card, Channel_P1A, &H50)
'Call MsecDelay(0.08)
'OpenPipe
'Call MsecDelay(0.04)
'rv5 = ReInitial(1)
'If rv5 = 1 Then
'    rv5 = ReInitial(4)
'End If
'ClosePipe
'
'If rv5 <> 1 Then
'    Call NewLabelMenu(5, "ReNew", rv5, rv4)
'    GoTo AU6479BLFResult
'Else
'    Call MsecDelay(0.2)
'End If

'===============================================
'  MSPro MS2 Card test Lun4
'================================================
'OpenPipe
'rv6 = RequestSense(4)
'ClosePipe
'If rv6 <> 1 Then
'    Call NewLabelMenu(6, "MS2", rv6, rv5)
'Else
'    rv4 = CBWTest_New_128_Sector_PipeReady(4, rv4)  ' write
'    Call NewLabelMenu(4, "MS2_64k", rv4, rv3)
'
'    If rv4 = 1 Then
'        rv4 = Read_Speed2ReadData(LBA, 4, 64)
'        If rv4 = 1 Then
'            If (ReadData(25) = &H2) And (ReadData(26) = &H70) And (ReadData(27) = &H1) Then
'                Tester.Print "MS is 4 Bit, Speed 40 MHz"
'            Else
'                Tester.Print "MS BusWidth/Speed Fail"
'                rv4 = 3
'            End If
'        End If
'        Call NewLabelMenu(4, "MS2 Bus Width/Speed", rv4, rv3)
'    End If
'End If
'Tester.Print rv4, " \\MSpro_2 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'

AU6479BLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'CF
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'SD0
        TestResult = "Bin3"
    ElseIf rv3 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv4 <> 1 Then        'MS2
        TestResult = "Bin5"
    'ElseIf rv5 <> 1 Then        'MS1
    '    TestResult = "Bin5"
    'ElseIf rv6 <> 1 Then        'SD1
    '    TestResult = "Bin3"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub


Public Sub AU6479BLF02TestSub()

'2012/5/3 for S/B: AU6479-GBL 100LQ SOCKET V0.90
'2012/6/14 for B62 version
'2012/6/18 for B62 FT3(HV/LV) test.

Dim TmpLBA As Long
Dim i As Integer
Dim ChipString As String
Dim DetectCount As Integer
Dim HV_Done_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

Routine_Label:

DetectCount = 0

If Not HV_Done_Flag Then
    Call PowerSet2(0, "3.6", "0.5", 1, "3.6", "0.5", 1)
    Call MsecDelay(0.3)
    Tester.Print "AU6479BL : HV Begin Test ..."
Else
    Call PowerSet2(0, "3.1", "0.5", 1, "3.1", "0.5", 1)
    Call MsecDelay(0.4)
    Tester.Print vbCrLf & "AU6479BL : LV Begin Test ..."
End If



OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'CF  (Lun0)
rv2 = 0     'SD0 (Lun1)
rv3 = 0     'MS1 (Lun2)
rv4 = 0     'XD  (Lun3)
rv5 = 0     'SD1 (Lun4)
rv6 = 0     'MS2 (Lun4)



Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '            CF ¡BSD0 ¡BMS1 ¡BXD  ¡BSD1                       CF ¡BSD0 ¡BMS1 ¡BXD  ¡BMS2(Can't not R/W)
                'Condition1(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)         Condition2(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)
'8:ENA    ---               0                                           0
'7:HID    ---               1                                           1
'6:M2INS  ---               1                                           0
'5:SD1CDN ---               0                                           1

'4:MSINS  ---               0                                           0
'3:XDCDN  ---               0                                           0
'2:CFCDN  ---               0                                           0
'1:SD0CDN ---               0                                           0

'                         0x60                                        0x50

'=========================================
'    POWER on
'=========================================

CardResult = DO_WritePort(card, Channel_P1A, &H60)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'If CardResult <> 0 Then
'    MsgBox "Read light off fail"
'    End
'End If
'
'Call MsecDelay(0.2)     'power on time
ChipString = "058f"
                
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOFF
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479BLFResult
End If



Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  CF Card test Lun0
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "CF", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)  ' write
    Call NewLabelMenu(1, "CF_64K", rv1, rv0)
End If
Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  SD Card test Lun1
'===============================================

rv2 = CBWTest_New_PipeReady(1, rv1, ChipString)
If rv2 <> 1 Then
    Call NewLabelMenu(2, "SD", rv2, rv1)
Else
    rv2 = CBWTest_New_128_Sector_PipeReady(1, rv2)  ' write
    Call NewLabelMenu(2, "SD_64K", rv2, rv1)
    
    If rv2 = 1 Then
        rv2 = Read_Speed2ReadData(LBA, 1, 64)
        If rv2 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv2 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(2, "SD Bus Speed/Width", rv2, rv1)
    End If
    
End If

Tester.Print rv2, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'=======================================================================================
'       SDHC R / W
'=======================================================================================
        
'Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
'OpenPipe
'rv1 = ReInitial(0)
'Call MsecDelay(0.02)
'rv1 = AU6435ForceSDHC(rv0)
'ClosePipe
'
'If rv1 = 1 Then
'    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
'End If
'
'
'If rv1 = 1 Then
'    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
'    If rv1 <> 1 Then
'        rv1 = 2
'        Tester.Print "SD2.0 Mode Fail"
'    End If
'End If
'
'ClosePipe
'
'Call LabelMenu(1, rv1, rv0)
'
'Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
'===============================================
'  MSPro MS1 Card test Lun2
'================================================
rv3 = CBWTest_New(2, rv2, ChipString)
If rv3 <> 1 Then
    Call NewLabelMenu(3, "MS", rv3, rv2)
Else
    rv3 = CBWTest_New_128_Sector_PipeReady(2, rv3)  ' write
    Call NewLabelMenu(3, "MS_64K", rv3, rv2)

    If rv3 = 1 Then
        rv3 = Read_Speed2ReadData(LBA, 2, 64)
        If rv3 = 1 Then
            If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                Tester.Print "MS is 4 Bit, Speed 40 MHz"
            Else
                Tester.Print "MS BusWidth/Speed Fail"
                rv3 = 3
            End If
        End If
        Call NewLabelMenu(3, "MS Bus Width/Speed", rv3, rv2)
    End If
End If
Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  
                  
'===============================================
'  XD Card test Lun3
'================================================
rv4 = CBWTest_New_PipeReady(3, rv3, ChipString)
If rv4 <> 1 Then
    Call NewLabelMenu(4, "XD", rv3, rv2)
Else
    rv4 = CBWTest_New_128_Sector_PipeReady(3, rv4)  ' write
    Call NewLabelMenu(4, "XD_64", rv3, rv2)
End If

Tester.Print rv4, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  SD1 Card test Lun4
'===============================================
rv5 = CBWTest_New_PipeReady(4, rv4, ChipString)
If rv5 <> 1 Then
    Call NewLabelMenu(5, "SD1", rv5, rv4)
Else
    rv5 = CBWTest_New_128_Sector_PipeReady(4, rv5)   ' write
    Call NewLabelMenu(5, "SD1_64K", rv5, rv4)

    If rv5 = 1 Then
        rv5 = Read_Speed2ReadData(LBA, 4, 64)
        If rv5 = 1 Then
            If (ReadData(20) = &H6) Then
                Tester.Print "SD is 4 Bit, Speed 48 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv5 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(5, "SD1 Bus Width/Speed", rv5, rv4)
    End If
End If

Tester.Print rv5, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


ClosePipe
If rv5 <> 1 Then
    GoTo AU6479BLFResult
End If

'            CF ¡BSD0 ¡BMS1 ¡BXD  ¡BMS2(Can't not R/W)
'Condition2(Lun0¡BLun1¡BLun2¡BLun3¡BLun4)
CardResult = DO_WritePort(card, Channel_P1A, &H50)
Call MsecDelay(0.08)
OpenPipe
Call MsecDelay(0.04)
rv5 = ReInitial(1)
If rv5 = 1 Then
    rv5 = ReInitial(4)
End If
ClosePipe

If rv5 <> 1 Then
    Call NewLabelMenu(5, "ReNew", rv5, rv4)
    GoTo AU6479BLFResult
Else
    Call MsecDelay(0.2)
End If

'===============================================
'  MSPro MS2 Card test Lun4
'================================================
OpenPipe
rv6 = RequestSense(4)
ClosePipe
If rv6 <> 1 Then
    Call NewLabelMenu(6, "MS2", rv6, rv5)
'Else
'    rv4 = CBWTest_New_128_Sector_PipeReady(4, rv4)  ' write
'    Call NewLabelMenu(4, "MS2_64k", rv4, rv3)
'
'    If rv4 = 1 Then
'        rv4 = Read_Speed2ReadData(LBA, 4, 64)
'        If rv4 = 1 Then
'            If (ReadData(25) = &H2) And (ReadData(26) = &H70) And (ReadData(27) = &H1) Then
'                Tester.Print "MS is 4 Bit, Speed 40 MHz"
'            Else
'                Tester.Print "MS BusWidth/Speed Fail"
'                rv4 = 3
'            End If
'        End If
'        Call NewLabelMenu(4, "MS2 Bus Width/Speed", rv4, rv3)
'    End If
End If
Tester.Print rv4, " \\MSpro_2 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'

AU6479BLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
    Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
    WaitDevOFF (ChipString)
    WaitDevOFF (ChipString)
    Call MsecDelay(0.3)
                
    If HV_Done_Flag = False Then
        If rv0 <> 1 Then
            HV_Result = "Bin2"
            Tester.Print "HV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 * rv6 <> 1 Then
            HV_Result = "Fail"
            Tester.Print "HV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 * rv6 = 1 Then
            HV_Result = "PASS"
            Tester.Print "HV PASS"
        End If
        
        HV_Done_Flag = True
        GoTo Routine_Label
    Else
        If rv0 <> 1 Then
            LV_Result = "Bin2"
            Tester.Print "LV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 * rv6 <> 1 Then
            LV_Result = "Fail"
            Tester.Print "LV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 * rv6 = 1 Then
            LV_Result = "PASS"
            Tester.Print "LV PASS"
        End If
        
    End If
            
            
    If (HV_Result = "Bin2") And (LV_Result = "Bin2") Then
        TestResult = "Bin2"
    ElseIf (HV_Result <> "PASS") And (LV_Result = "PASS") Then
        TestResult = "Bin3"
    ElseIf (HV_Result = "PASS") And (LV_Result <> "PASS") Then
        TestResult = "Bin4"
    ElseIf (HV_Result <> "PASS") And (LV_Result <> "PASS") Then
        TestResult = "Bin5"
    ElseIf (HV_Result = "PASS") And (LV_Result = "PASS") Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
                
'    If rv0 <> 1 Then            'Enum
'        UnknowDeviceFail = UnknowDeviceFail + 1
'        TestResult = "Bin2"
'    ElseIf rv1 <> 1 Then        'CF
'        TestResult = "Bin3"
'    ElseIf rv2 <> 1 Then        'SD0
'        TestResult = "Bin3"
'    ElseIf rv3 <> 1 Then        'XD
'        TestResult = "Bin4"
'    ElseIf rv4 <> 1 Then        'MS2
'        TestResult = "Bin5"
'    ElseIf rv5 <> 1 Then        'MS1
'        TestResult = "Bin5"
'    ElseIf rv6 <> 1 Then        'SD1
'        TestResult = "Bin3"
'    ElseIf rv6 * rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
'        TestResult = "PASS"
'    Else
'        TestResult = "Bin2"
'    End If
'
End Sub

Public Sub AU6479HLF21TestSub()

'2012/5/16 Use S/B:AU6435R-GHL

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

Call PowerSet2(0, "5.0", "0.5", 1, "5.0", "0.5", 1)

Tester.Print "AU6479HL Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'SD0 (Lun0)
rv2 = 0     'XD (Lun0)
rv3 = 0     'MS (Lun0)
rv4 = 0     'NBMD


Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                

' Ctrl_1            Ctrl_2

'8:                 8:
'7:                 7:
'6:                 6:
'5:                 5:

'4:XDCDN            4:
'3:MSINS            3:
'2:SDCDN            2:
'1:ENA              1:GPON7

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &HFC)      'SDCDN¡BENA On
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'If CardResult <> 0 Then
'    MsgBox "Read light off fail"
'    End
'End If
'
'Call MsecDelay(0.2)     'power on time
ChipString = "058f"
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOFF
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479HLFResult
End If

LBA = LBA + 1
Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  SD Card test
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "SD", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)    ' write
    Call NewLabelMenu(1, "SD_64K", rv1, rv0)
    
    If rv1 = 1 Then
        rv1 = Read_Speed2ReadData(LBA, 0, 64)
        If rv1 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv1 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(1, "SD Bus Speed/Width", rv1, rv0)
    End If
    
End If

Tester.Print rv1, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'=======================================================================================
'       SDHC R / W
'=======================================================================================
        
'Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
'OpenPipe
'rv1 = ReInitial(0)
'Call MsecDelay(0.02)
'rv1 = AU6435ForceSDHC(rv0)
'ClosePipe
'
'If rv1 = 1 Then
'    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
'End If
'
'
'If rv1 = 1 Then
'    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
'    If rv1 <> 1 Then
'        rv1 = 2
'        Tester.Print "SD2.0 Mode Fail"
'    End If
'End If
'
'ClosePipe
'
'Call LabelMenu(1, rv1, rv0)
'
'Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  

'===============================================
'  XD Card test
'===============================================
If rv1 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &HF6)  'XD¡BENA On
    rv2 = ReInitial(0)
    Call MsecDelay(0.1)
End If

If rv2 <> 1 Then
    Call NewLabelMenu(2, "Init", rv2, rv1)
Else
    rv2 = CBWTest_New_PipeReady(0, rv2, ChipString)
    
    If rv2 <> 1 Then
        Call NewLabelMenu(2, "XD", rv2, rv1)
    Else
        rv2 = CBWTest_New_128_Sector_PipeReady(0, rv2)    ' write
        Call NewLabelMenu(2, "XD_64", rv2, rv1)
    End If
    
End If



Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  MSPro MS1 Card test
'===============================================
If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &HFA)  'MS¡BENA On
    rv3 = ReInitial(0)
    Call MsecDelay(0.1)
End If

If rv3 <> 1 Then
    Call NewLabelMenu(3, "Init", rv3, rv2)
Else
    rv3 = CBWTest_New_PipeReady(0, rv3, ChipString)
    
    If rv3 <> 1 Then
        Call NewLabelMenu(3, "MS", rv3, rv2)
    Else
        rv3 = CBWTest_New_128_Sector_PipeReady(0, rv3)   ' write
        Call NewLabelMenu(3, "MS_64K", rv3, rv2)
    
        If rv3 = 1 Then
            rv3 = Read_Speed2ReadData(LBA, 0, 64)
            If rv3 = 1 Then
                If (ReadData(22) = &H2) And (ReadData(23) = &H70) And (ReadData(24) = &H1) Then
                    Tester.Print "MS is 4 Bit, Speed 40 MHz"
                Else
                    Tester.Print "MS BusWidth/Speed Fail"
                    rv3 = 3
                End If
            End If
            Call NewLabelMenu(3, "MS Bus Width/Speed", rv3, rv2)
        End If
    End If

End If


Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


AU6479HLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &H1)   ' Close power
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'SD
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv3 <> 1 Then        'MS
        TestResult = "Bin5"
    ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub

Public Sub AU6479HLF22TestSub()

'2012/5/16 Use S/B:AU6435R-GHL
'2012/6/14 for B62 version

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

Call PowerSet2(0, "5.0", "0.5", 1, "5.0", "0.5", 1)

Tester.Print "AU6479HL Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'SD0 (Lun0)
rv2 = 0     'XD (Lun0)
rv3 = 0     'MS (Lun0)
rv4 = 0     'NBMD


Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                

' Ctrl_1            Ctrl_2

'8:                 8:
'7:                 7:
'6:                 6:
'5:                 5:

'4:XDCDN            4:
'3:MSINS            3:
'2:SDCDN            2:
'1:ENA              1:GPON7

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &HFC)      'SDCDN¡BENA On
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'If CardResult <> 0 Then
'    MsgBox "Read light off fail"
'    End
'End If
'
'Call MsecDelay(0.2)     'power on time
ChipString = "058f"
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOFF
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479HLFResult
End If

LBA = LBA + 1
Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  SD Card test
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "SD", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)    ' write
    Call NewLabelMenu(1, "SD_64K", rv1, rv0)
    
    If rv1 = 1 Then
        rv1 = Read_Speed2ReadData(LBA, 0, 64)
        If rv1 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv1 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(1, "SD Bus Speed/Width", rv1, rv0)
    End If
    
End If

Tester.Print rv1, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'=======================================================================================
'       SDHC R / W
'=======================================================================================
        
'Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
'OpenPipe
'rv1 = ReInitial(0)
'Call MsecDelay(0.02)
'rv1 = AU6435ForceSDHC(rv0)
'ClosePipe
'
'If rv1 = 1 Then
'    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
'End If
'
'
'If rv1 = 1 Then
'    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
'    If rv1 <> 1 Then
'        rv1 = 2
'        Tester.Print "SD2.0 Mode Fail"
'    End If
'End If
'
'ClosePipe
'
'Call LabelMenu(1, rv1, rv0)
'
'Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  

'===============================================
'  XD Card test
'===============================================
If rv1 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &HF6)  'XD¡BENA On
    rv2 = ReInitial(0)
    Call MsecDelay(0.1)
End If

If rv2 <> 1 Then
    Call NewLabelMenu(2, "Init", rv2, rv1)
Else
    rv2 = CBWTest_New_PipeReady(0, rv2, ChipString)
    
    If rv2 <> 1 Then
        Call NewLabelMenu(2, "XD", rv2, rv1)
    Else
        rv2 = CBWTest_New_128_Sector_PipeReady(0, rv2)    ' write
        Call NewLabelMenu(2, "XD_64", rv2, rv1)
    End If
    
End If



Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  MSPro MS1 Card test
'===============================================
If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &HFA)  'MS¡BENA On
    rv3 = ReInitial(0)
    Call MsecDelay(0.1)
End If

If rv3 <> 1 Then
    Call NewLabelMenu(3, "Init", rv3, rv2)
Else
    rv3 = CBWTest_New_PipeReady(0, rv3, ChipString)
    
    If rv3 <> 1 Then
        Call NewLabelMenu(3, "MS", rv3, rv2)
    Else
        rv3 = CBWTest_New_128_Sector_PipeReady(0, rv3)   ' write
        Call NewLabelMenu(3, "MS_64K", rv3, rv2)
    
        If rv3 = 1 Then
            rv3 = Read_Speed2ReadData(LBA, 0, 64)
            If rv3 = 1 Then
                If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                    Tester.Print "MS is 4 Bit, Speed 40 MHz"
                Else
                    Tester.Print "MS BusWidth/Speed Fail"
                    rv3 = 3
                End If
            End If
            Call NewLabelMenu(3, "MS Bus Width/Speed", rv3, rv2)
        End If
    End If

End If


Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


AU6479HLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &H1)   ' Close power
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'SD
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv3 <> 1 Then        'MS
        TestResult = "Bin5"
    ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub

Public Sub AU6479WLF22TestSub()

'2012/5/16 Use S/B:AU6435R-GHL
'2014/1/22 This code copy from AU6479HLF22TestSub, Add NBMD test for AU6479C64-GWL


Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

Tester.Print "AU6479WL Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'SD0 (Lun0)
rv2 = 0     'XD (Lun0)
rv3 = 0     'MS (Lun0)
rv4 = 0     'NBMD


Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                

' Ctrl_1            Ctrl_2

'8:                 8:
'7:                 7:
'6:                 6:
'5:                 5:

'4:XDCDN            4:
'3:MSINS            3:
'2:SDCDN            2:
'1:ENA              1:GPON7

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &HFC)      'SDCDN¡BENA On
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'If CardResult <> 0 Then
'    MsgBox "Read light off fail"
'    End
'End If
'
'Call MsecDelay(0.2)     'power on time
ChipString = "6366"
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOFF
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "GPO", rv1, rv0)
    GoTo AU6479HLFResult
End If

If LBA < 100 Then
    LBA = 5000
End If

LBA = LBA + 1
Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  SD Card test
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "SD", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)    ' write
    Call NewLabelMenu(1, "SD_64K", rv1, rv0)
    
    If rv1 = 1 Then
        rv1 = Read_Speed2ReadData(LBA, 0, 64)
        If rv1 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv1 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(1, "SD Bus Speed/Width", rv1, rv0)
    End If
    
End If

Tester.Print rv1, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'=======================================================================================
'       SDHC R / W
'=======================================================================================
        
'Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
'OpenPipe
'rv1 = ReInitial(0)
'Call MsecDelay(0.02)
'rv1 = AU6435ForceSDHC(rv0)
'ClosePipe
'
'If rv1 = 1 Then
'    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
'End If
'
'
'If rv1 = 1 Then
'    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
'    If rv1 <> 1 Then
'        rv1 = 2
'        Tester.Print "SD2.0 Mode Fail"
'    End If
'End If
'
'ClosePipe
'
'Call LabelMenu(1, rv1, rv0)
'
'Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  

'===============================================
'  XD Card test
'===============================================
If rv1 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &HF4)  'SD¡BXD¡BENA On
    Call MsecDelay(0.04)
    CardResult = DO_WritePort(card, Channel_P1A, &HF6)  'XD¡BENA On
    rv2 = ReInitial(0)
End If

If rv2 <> 1 Then
    Call NewLabelMenu(2, "Init", rv2, rv1)
Else
    rv2 = CBWTest_New_PipeReady(0, rv2, ChipString)
    
    If rv2 <> 1 Then
        Call NewLabelMenu(2, "XD", rv2, rv1)
    Else
        rv2 = CBWTest_New_128_Sector_PipeReady(0, rv2)    ' write
        Call NewLabelMenu(2, "XD_64", rv2, rv1)
    End If
    
End If



Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  MSPro MS1 Card test
'===============================================
If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &HF2)  'XD¡BMS¡BENA On
    Call MsecDelay(0.04)
    CardResult = DO_WritePort(card, Channel_P1A, &HFA)  'MS¡BENA On
    rv3 = ReInitial(0)
End If

If rv3 <> 1 Then
    Call NewLabelMenu(3, "Init", rv3, rv2)
Else
    rv3 = CBWTest_New_PipeReady(0, rv3, ChipString)
    
    If rv3 <> 1 Then
        Call NewLabelMenu(3, "MS", rv3, rv2)
    Else
        rv3 = CBWTest_New_128_Sector_PipeReady(0, rv3)   ' write
        Call NewLabelMenu(3, "MS_64K", rv3, rv2)
    
        If rv3 = 1 Then
            rv3 = Read_Speed2ReadData(LBA, 0, 64)
            If rv3 = 1 Then
                If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                    Tester.Print "MS is 4 Bit, Speed 40 MHz"
                Else
                    Tester.Print "MS BusWidth/Speed Fail"
                    rv3 = 3
                End If
            End If
            Call NewLabelMenu(3, "MS Bus Width/Speed", rv3, rv2)
        End If
    End If

End If


Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


If rv3 = 1 Then

    CardResult = DO_WritePort(card, Channel_P1A, &H7)   'Ena(L)¡BSDCDN(H)¡BXDCDN(H)¡BMSIN(H)
    WaitDevOFF (ChipString)
    If GetDeviceName_NoReply(ChipString) = "" Then
        rv4 = 1
    Else
        rv4 = 0
    End If
End If


AU6479HLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &H1)   ' Close power
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'SD
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv3 <> 1 Then        'MS
        TestResult = "Bin5"
    ElseIf rv4 <> 1 Then        'NBMD
        TestResult = "Bin5"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub

Public Sub AU6479FLF23TestSub()

'2013/2/21 Use S/B:AU6479GFL 48LQ SOCKET
'2013/2/21 for C63,C64 version

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer
Dim k As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

Tester.Print "AU6479FL Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'SD0 (Lun0)
rv2 = 0     'CF (Lun0)
rv3 = 0     'XD (Lun0)
rv4 = 0     'MS (Lun0)

Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                

' Ctrl_1            Ctrl_2

'8:ENA              8:
'7:                 7:
'6:                 6:
'5:                 5:

'4:MSINS            4:
'3:XDCDN            3:
'2:CFCDN            2:
'1:SDCDN            1:GPON7

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H7F)      'ENA On
Call MsecDelay(0.1)
CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
    MsgBox "Read light off fail"
    End
End If


CardResult = DO_WritePort(card, Channel_P1A, &H7E)      'SDCDN¡BENA On
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

ChipString = "pid_6366"
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)

If rv0 <> 1 Then
    GoTo AU6479FLFResult
End If

CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Or ((LightOff And &H1) <> 1) Then
    Tester.Print "LightON="; LightOn
    Tester.Print "LightOFF="; LightOff
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "LED", rv1, rv0)
    GoTo AU6479FLFResult
End If

If LBA < 1000 Then
    LBA = 1000
Else
    LBA = LBA + 1
End If

Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  SD Card test
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)

' 20131030 v24 do not loop
If ((ChipName = "AU6479ULF23") Or (ChipName = "AU6479FLF23")) Then
    If rv1 <> 1 Then
        For k = 0 To 9
            rv1 = CBWTest_New(0, rv0, ChipString)
            If rv1 = 1 Then
                Exit For
            End If
        Next
    End If
End If

If rv1 <> 1 Then
    Call NewLabelMenu(1, "SD", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)    ' write
    Call NewLabelMenu(1, "SD_64K", rv1, rv0)
    
    If rv1 = 1 Then
        rv1 = Read_Speed2ReadData(LBA, 0, 64)
        If rv1 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv1 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(1, "SD Bus Speed/Width", rv1, rv0)
    End If
    
End If

Tester.Print rv1, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  CF Card test
'===============================================
If rv1 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H7D)  'CF¡BENA On
    rv2 = ReInitial(0)
    Call MsecDelay(0.1)
End If

If rv2 <> 1 Then
    Call NewLabelMenu(2, "Init", rv2, rv1)
Else
    rv2 = CBWTest_New_PipeReady(0, rv2, ChipString)
    
    If rv2 <> 1 Then
        Call NewLabelMenu(2, "CF", rv2, rv1)
    Else
        rv2 = CBWTest_New_128_Sector_PipeReady(0, rv2)    ' write
        Call NewLabelMenu(2, "CF_64", rv2, rv1)
    End If
    
End If

Tester.Print rv2, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  XD Card test
'===============================================
If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H7B)  'XD¡BENA On
    rv3 = ReInitial(0)
    Call MsecDelay(0.1)
End If

If rv3 <> 1 Then
    Call NewLabelMenu(3, "Init", rv3, rv2)
Else
    rv3 = CBWTest_New_PipeReady(0, rv3, ChipString)
    
    If rv3 <> 1 Then
        Call NewLabelMenu(3, "XD", rv3, rv2)
    Else
        rv3 = CBWTest_New_128_Sector_PipeReady(0, rv3)    ' write
        Call NewLabelMenu(3, "XD_64", rv3, rv2)
    End If
    
End If

Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  MSPro Card test
'===============================================
If rv3 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H77)  'MS¡BENA On
    rv4 = ReInitial(0)
    Call MsecDelay(0.1)
End If

If rv4 <> 1 Then
    Call NewLabelMenu(4, "Init", rv4, rv3)
Else
    rv4 = CBWTest_New_PipeReady(0, rv4, ChipString)
    
    If rv4 <> 1 Then
        Call NewLabelMenu(3, "MS", rv4, rv3)
    Else
        rv4 = CBWTest_New_128_Sector_PipeReady(0, rv4)   ' write
        Call NewLabelMenu(3, "MS_64K", rv4, rv3)
    
        If rv4 = 1 Then
            rv4 = Read_Speed2ReadData(LBA, 0, 64)
            If rv4 = 1 Then
                If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                    Tester.Print "MS is 4 Bit, Speed 40 MHz"
                Else
                    Tester.Print "MS BusWidth/Speed Fail"
                    rv4 = 3
                End If
            End If
            Call NewLabelMenu(4, "MS Bus Width/Speed", rv4, rv3)
        End If
    End If

End If

Tester.Print rv4, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


AU6479FLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
    WaitDevOFF (ChipString)
                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'SD
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'CF
        TestResult = "Bin3"
    ElseIf rv3 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv4 <> 1 Then        'MS
        TestResult = "Bin5"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub

Public Sub AU6479PLF32TestSub()

'2014/6/25 Use S/B:AU6479GPL 48LQFP SOCKET V1
'2014/6/25 for C64 version

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer
Dim k As Integer

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If


'==========================================================
'    Start Open Shot test
'==========================================================

OS_Result = 0
rv5 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)

MsecDelay (0.3)

OpenShortTest_SkipZero

rv5 = OS_Result

If rv5 <> 1 Then  'OS Fail
    GoTo AU6479PLFResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)

DetectCount = 0


'==========================================================
'    Start Function test
'==========================================================
Tester.Print "AU6479PL Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1

'rv5 'open short
rv0 = 0     'Enum
rv1 = 0     'SDHC (Lun0)
rv2 = 0     'XD (Lun0)
rv3 = 0     'MS (Lun0)
rv4 = 0     'MMC (Lun0)

Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                

' Ctrl_1            Ctrl_2

'8:ENA              8:
'7:                 7:
'6:XDCDN            6:
'5:GPIO3            5:

'4:XTALSEL          4:
'3:MMC              3:
'2:MSINS            2:
'1:SDCDN            1:GPON7

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H7F)      'ENA On

CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
    MsgBox "Read light Off fail"
    End
End If

CardResult = DO_WritePort(card, Channel_P1A, &H7E)      'SDCDN¡BENA On
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

ChipString = "pid_6366"
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)

If rv0 <> 1 Then
    GoTo AU6479PLFResult
End If

CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Or ((LightOff And &H1) <> 1) Then
    Tester.Print "LightON="; LightOn
    Tester.Print "LightOFF="; LightOff
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "LED", rv1, rv0)
    GoTo AU6479PLFResult
End If

If LBA < 1000 Then
    LBA = 5000
Else
    LBA = LBA + 1
End If

Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  SD Card test
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)

If rv1 <> 1 Then
    Call NewLabelMenu(1, "SD", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)    ' write
    Call NewLabelMenu(1, "SD_64K", rv1, rv0)

    If rv1 = 1 Then
        rv1 = Read_Speed2ReadData(LBA, 0, 64)
        If rv1 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv1 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(1, "SD Bus Speed/Width", rv1, rv0)
    End If

End If

Tester.Print rv1, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  XD Card test
'===============================================
If rv1 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H5F)  'XD¡BENA On
    rv2 = ReInitial(0)
    Call MsecDelay(0.1)
End If

If rv2 <> 1 Then
    Call NewLabelMenu(2, "Init", rv2, rv1)
Else
    rv2 = CBWTest_New_PipeReady(0, rv2, ChipString)

    If rv2 <> 1 Then
        Call NewLabelMenu(2, "XD", rv2, rv1)
    Else
        rv2 = CBWTest_New_128_Sector_PipeReady(0, rv2)    ' write
        Call NewLabelMenu(2, "XD_64", rv2, rv1)
    End If

End If

Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
'rv2 = 1

'===============================================
'  MSPro Card test
'===============================================
If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H7D)  'MS¡BENA On
    rv3 = ReInitial(0)
    Call MsecDelay(0.1)
End If

If rv3 <> 1 Then
    Call NewLabelMenu(3, "Init", rv3, rv2)
Else
    rv3 = CBWTest_New_PipeReady(0, rv3, ChipString)
    
    If rv3 <> 1 Then
        Call NewLabelMenu(3, "MS", rv3, rv2)
    Else
        rv3 = CBWTest_New_128_Sector_PipeReady(0, rv3)   ' write
        Call NewLabelMenu(3, "MS_64K", rv3, rv2)
    
        If rv3 = 1 Then
            rv3 = Read_Speed2ReadData(LBA, 0, 64)
            If rv3 = 1 Then
                If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                    Tester.Print "MS is 4 Bit, Speed 40 MHz"
                Else
                    Tester.Print "MS BusWidth/Speed Fail"
                    rv3 = 3
                End If
            End If
            Call NewLabelMenu(3, "MS Bus Width/Speed", rv3, rv2)
        End If
    End If

End If

Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  MMC Card test
'===============================================
If rv3 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H7B)  'MMC¡BENA On
    rv4 = ReInitial(0)
    Call MsecDelay(0.1)
End If

If rv4 <> 1 Then
    Call NewLabelMenu(4, "Init", rv4, rv3)
Else
    If (CBWTest_New_PipeReady(0, rv4, ChipString) <> 1) Then
        rv4 = 1
    Else
        rv4 = 2
    End If

    Call NewLabelMenu(4, "MMC", rv4, rv3)

End If

Tester.Print rv4, " \\MMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

AU6479PLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
    WaitDevOFF (ChipString)
                
    If (rv0 <> 1) Or (rv5 <> 1) Then           'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'SDHC
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'XD
        TestResult = "Bin4"
    ElseIf rv3 <> 1 Then        'MS
        TestResult = "Bin4"
    ElseIf rv4 <> 1 Then        'MMC
        TestResult = "Bin5"
    ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If

    
End Sub

Public Sub AU6479AFF33TestSub()

'2014/5/19 Use S/B:AU6479GAF 28QFN SOCKET V1
'2014/5/19 for C64 version

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer
Dim k As Integer

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

'==========================================================
'    Start Open Shot test
'==========================================================

OS_Result = 0
rv2 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_SkipZero

rv2 = OS_Result

If OS_Result <> 1 Then  'OS Fail
    GoTo AU6479AFFResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)



DetectCount = 0

Tester.Print "AU6479AF Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'SD0 (Lun0)
'rv2 = 0     'OpenShort
ChipString = "pid_6366"

Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                

' Ctrl_1            Ctrl_2

'8:ENA              8:
'7:                 7:
'6:                 6:
'5:                 5:

'4:                 4:
'3:                 3:
'2:                 2:
'1:SDCDN            1:GPON7

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H7F)      'ENA On
Call MsecDelay(0.1)
CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
    MsgBox "Read light off fail"
    End
End If


CardResult = DO_WritePort(card, Channel_P1A, &H7E)      'SDCDN¡BENA On
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If


             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)

If rv0 <> 1 Then
    GoTo AU6479AFFResult
End If

CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Or ((LightOff And &H1) <> 1) Then
    Tester.Print "LightON="; LightOn
    Tester.Print "LightOFF="; LightOff
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "LED", rv1, rv0)
    GoTo AU6479AFFResult
End If

If LBA < 1000 Then
    LBA = 1000
Else
    LBA = LBA + 1
End If

Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  SD Card test
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)

' 20131030 v24 do not loop
If ((ChipName = "AU6479ULF23") Or (ChipName = "AU6479FLF23")) Then
    If rv1 <> 1 Then
        For k = 0 To 9
            rv1 = CBWTest_New(0, rv0, ChipString)
            If rv1 = 1 Then
                Exit For
            End If
        Next
    End If
End If

If rv1 <> 1 Then
    Call NewLabelMenu(1, "SD", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)    ' write
    Call NewLabelMenu(1, "SD_64K", rv1, rv0)
    
    If rv1 = 1 Then
        rv1 = Read_Speed2ReadData(LBA, 0, 64)
        If rv1 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv1 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(1, "SD Bus Speed/Width", rv1, rv0)
    End If
    
End If

Tester.Print rv1, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


AU6479AFFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
    WaitDevOFF (ChipString)
                
    If rv2 <> 1 Then
        TestResult = "Bin2"
    ElseIf rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin4"
    ElseIf rv1 <> 1 Then        'SD
        TestResult = "Bin3"
    ElseIf rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin4"
    End If
    
End Sub

Public Sub AU6479FLF03TestSub()

'2013/2/21 Use S/B:AU6479GFL 48LQ SOCKET
'2013/4/26 for C63,C64 FT3


Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer
Dim HV_Done_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String


If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If

Dim ChipString As String

Routine_Label:


If Not HV_Done_Flag Then
    Call PowerSet2(0, "5.25", "0.3", 1, "5.25", "0.3", 1)
    Call MsecDelay(0.1)
    Tester.Print "AU6438BS : 5.25V Begin Test ..."
    SetSiteStatus (RunHV)
Else
    Call PowerSet2(0, "4.4", "0.3", 1, "4.4", "0.3", 1)
    Call MsecDelay(0.1)
    Tester.Print vbCrLf & "AU6438BS : 4.7V Begin Test ..."
    SetSiteStatus (RunFV)
End If
DetectCount = 0


OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'SD0 (Lun0)
rv2 = 0     'CF (Lun0)
rv3 = 0     'XD (Lun0)
rv4 = 0     'MS (Lun0)

Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                

' Ctrl_1            Ctrl_2

'8:ENA              8:
'7:                 7:
'6:                 6:
'5:                 5:

'4:MSINS            4:
'3:XDCDN            3:
'2:CFCDN            2:
'1:SDCDN            1:GPON7

'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &H7F)      'ENA On
'Call MsecDelay(0.4)
'CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
    MsgBox "Read light off fail"
    End
End If


CardResult = DO_WritePort(card, Channel_P1A, &H7E)      'SDCDN¡BENA On
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

ChipString = "pid_6366"
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)

If rv0 <> 1 Then
    GoTo AU6479FLFResult
End If

CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

Call MsecDelay(0.02)

If ((LightOn And &H1) <> 0) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOff
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 3
    Call NewLabelMenu(1, "LED", rv1, rv0)
    GoTo AU6479FLFResult
End If

If LBA < 1000 Then
    LBA = 1000
Else
    LBA = LBA + 1
End If

Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  SD Card test
'===============================================

rv1 = CBWTest_New(0, rv0, ChipString)
If rv1 <> 1 Then
    Call NewLabelMenu(1, "SD", rv1, rv0)
Else
    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)    ' write
    Call NewLabelMenu(1, "SD_64K", rv1, rv0)
    
    If rv1 = 1 Then
        rv1 = Read_Speed2ReadData(LBA, 0, 64)
        If rv1 = 1 Then
            If (ReadData(14) = &HA) Then
                Tester.Print "DDR Mode, Speed 50 MHz"
            ElseIf (ReadData(14) = &H9) Then
                Tester.Print "SDR Mode, Speed 120 MHz"
            ElseIf (ReadData(14) = &H8) Then
                Tester.Print "SDR Mode, Speed 100 MHz"
            ElseIf (ReadData(14) = &H7) Then
                Tester.Print "SDR Mode, Speed 80 MHz"
            Else
                Tester.Print "SD BusWidth/Speed Fail"
                rv1 = 3
            End If
        Else
            Tester.Print "SD Bus Speed/Width Fail"
        End If
        Call NewLabelMenu(1, "SD Bus Speed/Width", rv1, rv0)
    End If
    
End If

Tester.Print rv1, " \\SDXC_0 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  CF Card test
'===============================================
If rv1 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H7D)  'CF¡BENA On
    rv2 = ReInitial(0)
    Call MsecDelay(0.1)
End If

If rv2 <> 1 Then
    Call NewLabelMenu(2, "Init", rv2, rv1)
Else
    rv2 = CBWTest_New_PipeReady(0, rv2, ChipString)
    
    If rv2 <> 1 Then
        Call NewLabelMenu(2, "CF", rv2, rv1)
    Else
        rv2 = CBWTest_New_128_Sector_PipeReady(0, rv2)    ' write
        Call NewLabelMenu(2, "CF_64", rv2, rv1)
    End If
    
End If

Tester.Print rv2, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  XD Card test
'===============================================
If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H7B)  'XD¡BENA On
    rv3 = ReInitial(0)
    Call MsecDelay(0.1)
End If

If rv3 <> 1 Then
    Call NewLabelMenu(3, "Init", rv3, rv2)
Else
    rv3 = CBWTest_New_PipeReady(0, rv3, ChipString)
    
    If rv3 <> 1 Then
        Call NewLabelMenu(3, "XD", rv3, rv2)
    Else
        rv3 = CBWTest_New_128_Sector_PipeReady(0, rv3)    ' write
        Call NewLabelMenu(3, "XD_64", rv3, rv2)
    End If
    
End If

Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  MSPro Card test
'===============================================
If rv3 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H77)  'MS¡BENA On
    rv4 = ReInitial(0)
    Call MsecDelay(0.1)
End If

If rv4 <> 1 Then
    Call NewLabelMenu(4, "Init", rv4, rv3)
Else
    rv4 = CBWTest_New_PipeReady(0, rv4, ChipString)
    
    If rv4 <> 1 Then
        Call NewLabelMenu(3, "MS", rv4, rv3)
    Else
        rv4 = CBWTest_New_128_Sector_PipeReady(0, rv4)   ' write
        Call NewLabelMenu(3, "MS_64K", rv4, rv3)
    
        If rv4 = 1 Then
            rv4 = Read_Speed2ReadData(LBA, 0, 64)
            If rv4 = 1 Then
                If (ReadData(24) = &H2) And (ReadData(25) = &H70) And (ReadData(26) = &H1) Then
                    Tester.Print "MS is 4 Bit, Speed 40 MHz"
                Else
                    Tester.Print "MS BusWidth/Speed Fail"
                    rv4 = 3
                End If
            End If
            Call NewLabelMenu(4, "MS Bus Width/Speed", rv4, rv3)
        End If
    End If

End If

Tester.Print rv4, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


AU6479FLFResult:
                
    ClosePipe
    CardResult = DO_WritePort(card, Channel_P1A, &H80)
    If Not HV_Done_Flag Then
        SetSiteStatus (HVDone)
        Call WaitAnotherSiteDone(HVDone, 4#)
    Else
        SetSiteStatus (LVDone)
        Call WaitAnotherSiteDone(LVDone, 4#)
    End If
    Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
    Call MsecDelay(0.3)
    WaitDevOFF (ChipString)
    SetSiteStatus (SiteUnknow)
    
    If HV_Done_Flag = False Then
        If rv0 <> 1 Then
            HV_Result = "Bin2"
            Tester.Print "HV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 <> 1 Then
            HV_Result = "Fail"
            Tester.Print "HV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
            HV_Result = "PASS"
            Tester.Print "HV PASS"
        End If
        
        HV_Done_Flag = True
        GoTo Routine_Label
    Else
        If rv0 <> 1 Then
            LV_Result = "Bin2"
            Tester.Print "LV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 <> 1 Then
            LV_Result = "Fail"
            Tester.Print "LV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
            LV_Result = "PASS"
            Tester.Print "LV PASS"
        End If
        
    End If
            
            
    If (HV_Result = "Bin2") And (LV_Result = "Bin2") Then
        TestResult = "Bin2"
    ElseIf (HV_Result <> "PASS") And (LV_Result = "PASS") Then
        TestResult = "Bin3"
    ElseIf (HV_Result = "PASS") And (LV_Result <> "PASS") Then
        TestResult = "Bin4"
    ElseIf (HV_Result <> "PASS") And (LV_Result <> "PASS") Then
        TestResult = "Bin5"
    ElseIf (HV_Result = "PASS") And (LV_Result = "PASS") Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
    
'    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
'    WaitDevOFF (ChipString)
'
'    If rv0 <> 1 Then            'Enum
'        UnknowDeviceFail = UnknowDeviceFail + 1
'        TestResult = "Bin2"
'    ElseIf rv1 <> 1 Then        'SD
'        TestResult = "Bin3"
'    ElseIf rv2 <> 1 Then        'CF
'        TestResult = "Bin3"
'    ElseIf rv3 <> 1 Then        'XD
'        TestResult = "Bin4"
'    ElseIf rv4 <> 1 Then        'MS
'        TestResult = "Bin5"
'    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
'        TestResult = "PASS"
'    Else
'        TestResult = "Bin2"
'    End If
    
End Sub
