Attribute VB_Name = "AU7510MDL"
Option Explicit
'========= AU7510
Public Const WM_USER = &H400

Public Const WM_FT_USB_UNKNOW_FAIL_AU7510 = WM_USER + &H50
Public Const WM_FT_CHIP_UNKNOW_FAIL_AU7510 = WM_USER + &H75
Public Const WM_FT_RW_READY_AU7510 = WM_USER + &H100
 
Public Const WM_FT_SCAN_START_AU7510 = WM_USER + &H200
Public Const WM_FT_MP_START_AU7510 = WM_USER + &H250
Public Const WM_FT_RW_START_AU7510 = WM_USER + &H300
 
Public Const WM_FT_HOST_FAIL_AU7510 = WM_USER + &H350
Public Const WM_FT_SCAN_DONE_AU7510 = WM_USER + &H400
 
Public Const WM_FT_MP_FAIL_AU7510 = WM_USER + &H450
Public Const WM_FT_MP_DONE_AU7510 = WM_USER + &H500
 
Public Const WM_FT_WRITE_PROTECT_FAIL_AU7510 = WM_USER + &H550
Public Const WM_FT_RW_CE_FAIL_AU7510 = WM_USER + &H600
Public Const WM_FT_RW_ROM_FAIL_AU7510 = WM_USER + &H650
Public Const WM_FT_RW_RAM_FAIL_AU7510 = WM_USER + &H700
Public Const WM_FT_SATA_FAIL_AU7510 = WM_USER + &H750
Public Const WM_FT_GPIO_FAIL_AU7510 = WM_USER + &H760
Public Const WM_FT_PASS_AU7510 = WM_USER + &H800

Global card_number As Integer
Global card_type As Integer
Global ch_num As Integer
Global ch_range As Integer
Global ch_ref As Integer

Global ch_cnt As Integer
Global gnBuffer(4000) As Integer
Global gnBufferWMA(4000) As Integer
Global count1 As Long
Global Card2214 As Integer
Global RecordMode As Integer
Global RecordModeWMA As Integer
Global RecordModeMP3 As Integer

Global Card2214Exist As Integer
Public TmpChip As String
Public OldLBa As Long

'================= for Audio mode =================
Global ADCch_num(0 To 1) As Integer
Global ADCch_range(0 To 1) As Integer
Global ADCch_ref(0 To 1) As Integer
Global ADCBuffer(0 To 11999) As Integer
Public bStop As Boolean
'==================================================

Global MP3MatchIndex As Byte
Global WMAMatchIndex As Byte
Global MP3PatternIndex As Byte
Global WMAPatternIndex As Byte
Global PatternSlope(0 To 5999) As Long
Global RawDataSlope(0 To 5999) As Long

Global PatternAcc(0 To 5999) As Long
Global RawDataAcc(0 To 5999) As Long
Global SineCurve(0 To 1, 0 To 5999) As Long
Global SineCurve2(0 To 1, 0 To 5999) As Long

Global MP3RawData(0 To 1, 0 To 5999) As Integer
Global MP3Pattern(0 To 20, 0 To 1, 0 To 5999) As Integer

Global WMARawData(0 To 1, 0 To 5999) As Integer
Global WMAPattern(0 To 10, 0 To 1, 0 To 5999) As Integer

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
Public MP3Data_9(0 To 100) As Integer   '=============2007.06.14 CHEYENNE CHANG
Public MP3Data_BL(0 To 100) As Integer
Public MP3Data_BL1(0 To 100) As Integer
Public MP3Data_CW1(0 To 100) As Integer '=============2007.06.14 CHEYENNE CHANG
Public MP3Data_10(0 To 100) As Integer
Public MP3Data_11(0 To 100) As Integer
Public MP3Data_12(0 To 100) As Integer
Public MP3Data_13(0 To 100) As Integer
Public MP3Data_15(0 To 100) As Integer
Public MP3Data_16(0 To 100) As Integer
Public MP3Data_161(0 To 100) As Integer
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

 Public Function DACGetData2ChannelUpTriggleAU3152HL() As Integer
 
Dim result As Long
Dim status As Byte
Dim i As Long, k As Long
Dim ReTrgCnt As Integer, MCnt As Integer, PostCnt As Integer
Dim TimerSrc As Integer
Dim cobTrigModeListIndex As Integer
Dim cobTrigSrcListIndex As Integer
Dim cobTrigPolListIndex As Integer
Dim cobClkSrcListIndex As Integer
Dim ErrCount As Integer
Dim OldCount As Integer
Dim CoutinueCount As Integer
Dim cobTrigLevelText As String
Dim cobTrigL_LevelText As String
Dim gnClkDiv As Integer
Dim Id As Integer
Dim startPos As Long
Dim OldTimer
Dim ADCPostCnt As Integer
Dim ADCdma_size As Long
  
   ' dma2205.Hide
  ADCPostCnt = 50 'post trigger 500
  ADCdma_size = 100 ' scan count 2000-- 80000
  ch_cnt = 2
  ADCch_num(0) = 0  ' channel 0
  ADCch_range(0) = 2  ' BiPolar 1.25V, 1 base
 
  ADCch_ref(0) = 1    ' Differentail
  
  
  ADCch_num(1) = 1  ' channel 0
  ADCch_range(1) = 2  ' BiPolar 1.25V, 1 base
  
  ADCch_ref(0) = 1    ' Differentail
  
  '=================================
  '  DMA setting
  '=================================
  
      '=================================
  '  DMA setting
  '=================================
  
        cobTrigModeListIndex = 0
        cobTrigSrcListIndex = 0   ' sourece from software
        cobTrigPolListIndex = 1   '
        cobClkSrcListIndex = 0
      '  cobTrigPolListIndex = 1   ' because the layout has error at right-left way
        cobTrigLevelText = "156"
        cobTrigL_LevelText = "20"  ' because the layout has error
        gnClkDiv = 5000
   
       
    Call MsecDelay(0.5)   'for AU3150 unstable chip
  
   If Card2214Exist = 0 Then
        Card2214 = D2K_Register_Card(DAQ_2214, 0)
        
          For i = 0 To 1 Step 1
                result = D2K_AI_CH_Config(Card2214, ADCch_num(i), ADCch_range(i) Or (ADCch_ref(i) * 256))
          Next
          
   
          
         
    End If
   
   
  For i = 0 To 199
  ADCBuffer(i) = 0
  Next i
  
   
        result = D2K_AI_Config(Card2214, 0, (CLng(cobTrigModeListIndex) * 8) Or CLng(cobTrigSrcListIndex) Or (CLng(cobTrigPolListIndex) * 4096), ADCPostCnt, 0, 1, 1)
          If result <> 0 Then
            MsgBox "2214 configure  fail 1"
            Exit Function
          End If
   
    
 
    result = D2K_AI_ContBufferSetup(Card2214, ADCBuffer(0), ADCdma_size * ch_cnt, Id)
'   Tester.Print "result2 ="; result; "Id="; Id
  If result <> 0 Then
            MsgBox "2214 configure  fail 2"
            Exit Function
          End If
   

   
    result = D2K_AIO_Config(Card2214, cobClkSrcListIndex, (CLng(cobTrigPolListIndex) * 256) Or CH0ATRIG, CLng(cobTrigLevelText), CLng(cobTrigL_LevelText))
            If result <> 0 Then
            MsgBox "2214 configure fail 3"
            Exit Function
            End If
  'Tester.Print "result4="; result
  
  result = D2K_AI_ContReadMultiChannels(Card2214, ch_cnt, ADCch_num(0), Id, ADCdma_size, gnClkDiv * ch_cnt, gnClkDiv, ASYNCH_OP)
       If result <> 0 Then
            MsgBox "2214 configure fail 4"
            Exit Function
          End If
   
 '  Tester.Print "read AD begin ="; result
   Card2214Exist = 1
  
  
  
  
  status = 0
  bStop = 0
  OldTimer = Timer
  While status = 0 And bStop = 0 And (Timer - OldTimer < 10)
     DoEvents
    result = D2K_AI_AsyncCheck(Card2214, status, count1)
    
  Wend
  result = D2K_AI_AsyncClear(Card2214, startPos, count1)
  'Tester.Print "result5 ="; result;
   MPTester.Print "read AD end "
  
 End Function
Public Function DACGetData1Channel() As Integer
'On Error Resume Next
  Dim result As Long
  Dim status As Byte
  Dim i As Long, k As Long
  Dim ReTrgCnt As Integer, MCnt As Integer, PostCnt As Integer
  Dim TimerSrc As Integer
  Dim cobTrigModeListIndex As Integer
  Dim cobTrigSrcListIndex As Integer
  Dim cobTrigPolListIndex As Integer
  Dim cobClkSrcListIndex As Integer
  Dim ErrCount As Integer
  Dim OldCount As Integer
  Dim CoutinueCount As Integer
  Dim cobTrigLevelText As String
  Dim cobTrigL_LevelText As String
  Dim gnClkDiv As Integer
  Dim Id As Integer
  Dim startPos As Long
  Dim OldTimer
  Dim ADCPostCnt As Integer
  Dim ADCdma_size As Long
  
   ' dma2205.Hide
  ADCPostCnt = 1 ' post trigger 500
  ADCdma_size = 100 ' scan count 2000-- 80000
  ch_cnt = 1    ' 1 channel
  ADCch_num(0) = 0  ' channel 0
  'ADCch_range(0) = 2  ' BiPolar 5V, 1 base
   ADCch_range(0) = 2  ' BiPolar 0.5V, 1 base
 'ADCch_range(0) = 4
  ADCch_ref(0) = 1    ' singel end
  
  
  
  '=================================
  '  DMA setting
  '=================================
  
        cobTrigModeListIndex = 0 ' post trig
        cobTrigSrcListIndex = 0  ' software
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 1
        cobTrigLevelText = "200"
        cobTrigL_LevelText = "128"
        gnClkDiv = 5000
   
       
    Call MsecDelay(0.5)   'for AU3150 unstable chip
    

   If Card2214Exist = 0 Then
        Card2214 = D2K_Register_Card(DAQ_2214, 0)
        
          For i = 0 To 0 Step 1
                result = D2K_AI_CH_Config(Card2214, ADCch_num(i), ADCch_range(i) Or (ADCch_ref(i) * 256))
          Next
          
   
          
         
    End If
   
   
  For i = 0 To 99
  ADCBuffer(i) = 0
  Next i
  
   
        result = D2K_AI_Config(Card2214, 0, (CLng(cobTrigModeListIndex) * 8) Or CLng(cobTrigSrcListIndex) Or (CLng(cobTrigPolListIndex) * 4096), ADCPostCnt, 0, 1, 1)
          If result <> 0 Then
            MsgBox "2214 configure fail 1"
            Exit Function
          End If
   
   
  
   result = D2K_AI_ContBufferSetup(Card2214, ADCBuffer(0), ADCdma_size * ch_cnt, Id)
'   Tester.Print "result2 ="; result; "Id="; Id
   
    result = D2K_AIO_Config(Card2214, cobClkSrcListIndex, (CLng(cobTrigPolListIndex) * 256) Or CH0ATRIG, CLng(cobTrigLevelText), CLng(cobTrigL_LevelText))
            If result <> 0 Then
            MsgBox "2214 configure fail 2"
            Exit Function
            End If
  'Tester.Print "result4="; result
  result = D2K_AI_ContReadMultiChannels(Card2214, ch_cnt, ADCch_num(0), Id, ADCdma_size, gnClkDiv * ch_cnt, gnClkDiv, ASYNCH_OP)
     
 '  Tester.Print "read AD begin ="; result
   Card2214Exist = 1
  
  
  
  
  status = 0
  bStop = 0
  OldTimer = Timer
  While status = 0 And bStop = 0 And (Timer - OldTimer < 10)
     DoEvents
    result = D2K_AI_AsyncCheck(Card2214, status, count1)
    
  Wend
  result = D2K_AI_AsyncClear(Card2214, startPos, count1)
  MPTester.Print "result5 ="; result;
  MPTester.Print "read AD end "
  
 End Function


 
Public Sub AU7510A43ALF20TestSub()
'add unload driver function
 If PCI7248InitFinish = 0 Then
       PCI7248Exist
 End If
 Dim OldTimer
 Dim PassTime
 Dim rt2
 Dim LightOn
 Dim mMsg As MSG
 Dim LedCount As Byte
 
   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================
 
 'AU7510 do not have filter driver
 
 
     
                     


NewChipFlag = 0
If OldChipName <> ChipName Then

' reset program

                        winHwnd = FindWindow(vbNullString, "FT")
 
                         If winHwnd <> 0 Then
                           Do
                           rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                           Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "FT")
                           Loop While winHwnd <> 0
                         End If

           ' FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & chipname & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
           ' FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & chipname & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
           ' FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & chipname & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

 
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail


 '====================================
 '  Fix Card
 '====================================
 
 If NewChipFlag = 1 Or FindWindow(vbNullString, "FT") = 0 Then
    
 '==============================================================
' when begin  scan + MP
'===============================================================
  
 
   
       '  power on
     '  cardresult = DO_WritePort(card, Channel_P1A, &HFF)
     '  Call PowerSet(3)   ' close power to disable chip
     '  Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU7510
       
 
        OldTimer = Timer
        AlcorMPMessage = 0
        
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_RW_READY_AU7510 Or PassTime > 30 _
              Or AlcorMPMessage = WM_FT_USB_UNKNOW_FAIL_AU7510 _
              Or AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3, scan fail
        '====================================================
        '
       
          '(1) TIme out : scan again ; shut down and restart
          '(2) usb fail : scan
          '(3) scan fail: delay and recan, after scan two times fail then restart PC
               
      '========== initial stage
             '(1) time out fail stage : restart
            ' If PassTime > 30 Then
             
             '   MsgBox "FT program no response"
             '   Exit Sub
           ' End If
                ' initial time out fail
               If PassTime > 30 Then    'usb issue so when time out , we let restart PC
               
               'restart PC
                        MPTester.TestResultLab = "Bin3:MP Scan Fail"
                        TestResult = "Bin3"
                        MPTester.Print "Scan Fail"
                        ResetFlag = 1
               
               End If
             
             '(2) initial fail :usb fail stage
      
             If AlcorMPMessage = WM_FT_USB_UNKNOW_FAIL_AU7510 Or AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 Then
               
                
SCAN_LABEL:
                Call StartScan_Click_AU7510  ' begin scan
                
               
                OldTimer = Timer
                AlcorMPMessage = 0
        
                Do
                   ' DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                         AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                    DispatchMessage mMsg
                 End If
                    PassTime = Timer - OldTimer
                Loop Until AlcorMPMessage = WM_FT_SCAN_DONE_AU7510 _
                Or AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 _
                Or PassTime > 30 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
       
                
                
                
                MPTester.Print "MP work time="; PassTime
                MPTester.MPText.Text = Hex(AlcorMPMessage)
                
                
                '(21) time out : restart PC
                  If PassTime > 30 Then  'usb issue so when time out , we let restart PC
                        MPTester.TestResultLab = "Bin3:MP Scan Fail"
                        TestResult = "Bin3"
                        MPTester.Print "Scan Fail"
                        ResetFlag = 1
                  End If
                  
                 '(22) usb fail : rescan
                  If AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 Then
                  
                    ScanContFail = ScanContFail + 1
                    If ScanContFail > 3 Then
                        
                        MPTester.TestResultLab = "Bin3:MP Scan Fail"
                        TestResult = "Bin3"
                        MPTester.Print "Scan Fail"
             
                      Exit Sub
                    End If
                    
                    Call MsecDelay(5#)
                     GoTo SCAN_LABEL
                End If
                
                
                '(23) pass
                   
                
                   If AlcorMPMessage = WM_FT_SCAN_DONE_AU7510 Then
                  
                       ScanContFail = 0
                    
                        
                        MPTester.TestResultLab = "SCAN PASS"
                       
                        MPTester.Print "Scan PASS"
             
                   
                   End If
          
               
        End If
 End If
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_RW_READY_AU7510 Or MPFlag = 1 Or ContFail >= 5 Or MPTester.Check1.Value = 1 Then
           
         MPFlag = 1
             ScanContFail = 0
       
             
            ' Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU7510
   
             
              OldTimer = Timer
              AlcorMPMessage = 0
        
                Do
                   ' DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                         AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                    DispatchMessage mMsg
                 End If
                    PassTime = Timer - OldTimer
                Loop Until AlcorMPMessage = WM_FT_MP_DONE_AU7510 _
                Or AlcorMPMessage = WM_FT_MP_FAIL_AU7510 _
                Or PassTime > 90 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               '(31) MP time out fail, close FT program
                If PassTime > 90 Then   ' this is chip issue, so do not reset PC
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    ' (1)
                  
                 
                    Exit Sub
                End If
                
                '(32) MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL_AU7510 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    If MPContFail > 5 Then
                      
                    ResetFlag = 1
                    End If
                 
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
               
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_DONE_AU7510 Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPFlag = 0
                    MPTester.Print "MP PASS"
                End If
                
              
                
      End If
  

                        
 '=========================================
 '    POWER on
 '=========================================
 
 
      
        
        
        MPTester.Print "wait for RW Tester Ready"
        OldTimer = Timer
        AlcorMPMessage = 0
        Do
            'DoEvents
             If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
             End If
        
             PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_RW_READY_AU7510 Or PassTime > 15 _
        Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        MPTester.Print "RW Ready Time="; PassTime
        
       If PassTime > 15 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
          
            Exit Sub
       End If
         
 Dim ADCFlag As Byte
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU7510
        ADCFlag = 0
        Do
             If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
             
            PassTime = Timer - OldTimer
            If ADCFlag = 0 Then
                MPTester.TestResultLab = "ADC begin"
         '   Call MsecDelay(1)
            
             Call DACGetData2ChannelUpTriggleAU3152HL
             ADCFlag = 1
            End If
        Loop Until AlcorMPMessage = WM_FT_CHIP_UNKNOW_FAIL_AU7510 _
              Or AlcorMPMessage = WM_FT_SATA_FAIL_AU7510 _
              Or AlcorMPMessage = WM_FT_RW_RAM_FAIL_AU7510 _
               Or AlcorMPMessage = WM_FT_RW_ROM_FAIL_AU7510 _
               Or AlcorMPMessage = WM_FT_RW_CE_FAIL_AU7510 _
                 Or AlcorMPMessage = WM_FT_PASS_AU7510 _
                 Or AlcorMPMessage = WM_FT_WRITE_PROTECT_FAIL_AU7510 _
                 Or AlcorMPMessage = WM_FT_GPIO_FAIL_AU7510 _
              Or PassTime > 20 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
          MPTester.Print "RW work Time="; PassTime
          MPTester.MPText.Text = Hex(AlcorMPMessage)
        
        
        '===========================================================
        '  RW Time Out Fail
        '===========================================================
        
        If PassTime > 20 Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Time Out Fail"
   
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_CHIP_UNKNOW_FAIL_AU7510
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_SATA_FAIL_AU7510
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_CE_FAIL_AU7510, WM_FT_WRITE_PROTECT_FAIL_AU7510, WM_FT_GPIO_FAIL_AU7510
             TestResult = "Bin3"
             If AlcorMPMessage = WM_FT_RW_CE_FAIL_AU7510 Then
               MPTester.TestResultLab = "Bin3:RW FAIL "
             ElseIf AlcorMPMessage = WM_FT_WRITE_PROTECT_FAIL_AU7510 Then
               MPTester.TestResultLab = "Bin3:WRITE PROTECT FAIL "
             End If
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL_AU7510
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL_AU7510
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_PASS_AU7510
        
 Dim AU7510NoDramUpLimit As Integer
 Dim AU7510WithDramUpLimit As Integer
 Dim i As Long
 Dim TmpHSum As Single
 Dim TmpHAvg As Single
 Dim TmpLSum As Single
 Dim TmpLAvg As Single
 
 AU7510WithDramUpLimit = 220
 AU7510NoDramUpLimit = 220
              ' to measure currennt
 'Dim i As Integer
        
 '====================================
 '   A2D begin
 '====================================
   


 
' dma2205.Show
' dma2205.ShowDataAU6476 0, 100
' If NewChipFlag <> 1 Then
 Dim TmpH As Single
 Dim TmpL As Single
 Dim TmpDiff As Single
 Dim TmpDiffSum As Single
 Dim TmpDiffAvg As Single
 Dim Interval As Single
 Dim DiffCounter As Integer
 Interval = 5

 For i = 0 To 99
     
     
     TmpH = CSng(ADCBuffer(i * 2) + 1) / CSng(32768) * 5# ' ma
     TmpL = CSng(CSng(ADCBuffer(i * 2 + 1)) + 1) / CSng(32768) * 5#
     TmpHSum = TmpHSum + TmpH
     TmpLSum = TmpLSum + TmpL
     TmpDiff = (TmpH - TmpL) * 1000
    ' 'debug.print i; "H"; ADCBuffer(i * 2); TmpH
     ''debug.print i; "L"; ADCBuffer(i * 2 + 1); TmpL
     ' 'debug.print i; TmpH; TmpL; TmpDiff
      If TmpDiff > 0 Then
        DiffCounter = DiffCounter + 1
        TmpDiffSum = TmpDiffSum + TmpDiff
       If TmpDiff > AU7510NoDramUpLimit Or TmpH < 3 Then
                  TestResult = "Bin3"
                    MPTester.TestResultLab = "current:" & CStr(TmpDiff)
                  
                   ContFail = ContFail + 1
        Exit Sub
     End If
     End If
 
 Next i
      'debug.print "1avg"; TmpHSum * 0.01; TmpLSum * 0.01;
      If DiffCounter <> 0 Then
        TmpDiffAvg = CSng(TmpDiffSum / DiffCounter)
        'debug.print "Diffavg"; TmpDiffAvg
        If TmpDiffAvg > 220 Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "current:" & CStr(TmpDiff)
            ContFail = ContFail + 1
            Exit Sub
        End If
          
      End If
' End If
        
        
               
          '     For LedCount = 1 To 20
          '     Call MsecDelay(0.1)
          '     cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
          '      If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
               
          '       Exit For
          '     End If
           '    Next LedCount
                 
           '       MPTester.Print "light="; LightOn
           '      If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
                     MPTester.TestResultLab = "PASS "
                     
                      MPTester.Print "Current="; (TmpHSum - TmpLSum) * 1000
                     TestResult = "PASS"
                     ContFail = 0
           '     Else
                 
           '       TestResult = "Bin3"
           '       MPTester.TestResultLab = "Bin3:LED FAIL "
              
           '    End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
        If TestResult <> "PASS" Then
            Call StartScan_Click_AU7510  ' begin scan
        End If
        'Call PowerSet(1500)
        
         
                            
End Sub

Public Sub AU7510A41ALF20TestSub()
'add unload driver function
 If PCI7248InitFinish = 0 Then
       PCI7248Exist
 End If
 Dim OldTimer
 Dim PassTime
 Dim rt2
 Dim LightOn
 Dim mMsg As MSG
 Dim LedCount As Byte
 
   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================
 
 'AU7510 do not have filter driver
 
 
     
                     


NewChipFlag = 0
If OldChipName <> ChipName Then

' reset program

                        winHwnd = FindWindow(vbNullString, "FT")
 
                         If winHwnd <> 0 Then
                           Do
                           rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                           Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "FT")
                           Loop While winHwnd <> 0
                         End If

           ' FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & chipname & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
           ' FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & chipname & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
           ' FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & chipname & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

 
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail


 '====================================
 '  Fix Card
 '====================================
 
 If NewChipFlag = 1 Or FindWindow(vbNullString, "FT") = 0 Then
    
 '==============================================================
' when begin  scan + MP
'===============================================================
  
 
   
       '  power on
     '  cardresult = DO_WritePort(card, Channel_P1A, &HFF)
     '  Call PowerSet(3)   ' close power to disable chip
     '  Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU7510
       
 
        OldTimer = Timer
        AlcorMPMessage = 0
        
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_RW_READY_AU7510 Or PassTime > 30 _
              Or AlcorMPMessage = WM_FT_USB_UNKNOW_FAIL_AU7510 _
              Or AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3, scan fail
        '====================================================
        '
       
          '(1) TIme out : scan again ; shut down and restart
          '(2) usb fail : scan
          '(3) scan fail: delay and recan, after scan two times fail then restart PC
               
      '========== initial stage
             '(1) time out fail stage : restart
            ' If PassTime > 30 Then
             
             '   MsgBox "FT program no response"
             '   Exit Sub
           ' End If
                ' initial time out fail
               If PassTime > 30 Then    'usb issue so when time out , we let restart PC
               
               'restart PC
                        MPTester.TestResultLab = "Bin3:MP Scan Fail"
                        TestResult = "Bin3"
                        MPTester.Print "Scan Fail"
                        ResetFlag = 1
               
               End If
             
             '(2) initial fail :usb fail stage
      
             If AlcorMPMessage = WM_FT_USB_UNKNOW_FAIL_AU7510 Or AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 Then
               
                
SCAN_LABEL:
                Call StartScan_Click_AU7510  ' begin scan
                
               
                OldTimer = Timer
                AlcorMPMessage = 0
        
                Do
                   ' DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                         AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                    DispatchMessage mMsg
                 End If
                    PassTime = Timer - OldTimer
                Loop Until AlcorMPMessage = WM_FT_SCAN_DONE_AU7510 _
                Or AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 _
                Or PassTime > 30 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
       
                
                
                
                MPTester.Print "MP work time="; PassTime
                MPTester.MPText.Text = Hex(AlcorMPMessage)
                
                
                '(21) time out : restart PC
                  If PassTime > 30 Then  'usb issue so when time out , we let restart PC
                        MPTester.TestResultLab = "Bin3:MP Scan Fail"
                        TestResult = "Bin3"
                        MPTester.Print "Scan Fail"
                        ResetFlag = 1
                  End If
                  
                 '(22) usb fail : rescan
                  If AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 Then
                  
                    ScanContFail = ScanContFail + 1
                    If ScanContFail > 3 Then
                        
                        MPTester.TestResultLab = "Bin3:MP Scan Fail"
                        TestResult = "Bin3"
                        MPTester.Print "Scan Fail"
             
                      Exit Sub
                    End If
                    
                    Call MsecDelay(5#)
                     GoTo SCAN_LABEL
                End If
                
                
                '(23) pass
                   
                
                   If AlcorMPMessage = WM_FT_SCAN_DONE_AU7510 Then
                  
                       ScanContFail = 0
                    
                        
                        MPTester.TestResultLab = "SCAN PASS"
                       
                        MPTester.Print "Scan PASS"
             
                   
                   End If
          
               
        End If
 End If
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_RW_READY_AU7510 Or MPFlag = 1 Or ContFail >= 5 Or MPTester.Check1.Value = 1 Then
           
         MPFlag = 1
             ScanContFail = 0
       
             
            ' Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU7510
   
             
              OldTimer = Timer
              AlcorMPMessage = 0
        
                Do
                   ' DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                         AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                    DispatchMessage mMsg
                 End If
                    PassTime = Timer - OldTimer
                Loop Until AlcorMPMessage = WM_FT_MP_DONE_AU7510 _
                Or AlcorMPMessage = WM_FT_MP_FAIL_AU7510 _
                Or PassTime > 90 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               '(31) MP time out fail, close FT program
                If PassTime > 90 Then   ' this is chip issue, so do not reset PC
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    ' (1)
                  
                 
                    Exit Sub
                End If
                
                '(32) MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL_AU7510 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    If MPContFail > 5 Then
                      
                    ResetFlag = 1
                    End If
                 
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
               
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_DONE_AU7510 Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPFlag = 0
                    MPTester.Print "MP PASS"
                End If
                
              
                
      End If
  

                        
 '=========================================
 '    POWER on
 '=========================================
 
 
      
        
        
        MPTester.Print "wait for RW Tester Ready"
        OldTimer = Timer
        AlcorMPMessage = 0
        Do
            'DoEvents
             If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
             End If
        
             PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_RW_READY_AU7510 Or PassTime > 15 _
        Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        MPTester.Print "RW Ready Time="; PassTime
        
       If PassTime > 15 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
          
            Exit Sub
       End If
         
 Dim ADCFlag As Byte
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU7510
        ADCFlag = 0
        Do
             If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
             
            PassTime = Timer - OldTimer
            If ADCFlag = 0 Then
                MPTester.TestResultLab = "ADC begin"
         '   Call MsecDelay(1)
            
             Call DACGetData2ChannelUpTriggleAU3152HL
             ADCFlag = 1
            End If
        Loop Until AlcorMPMessage = WM_FT_CHIP_UNKNOW_FAIL_AU7510 _
              Or AlcorMPMessage = WM_FT_SATA_FAIL_AU7510 _
              Or AlcorMPMessage = WM_FT_RW_RAM_FAIL_AU7510 _
               Or AlcorMPMessage = WM_FT_RW_ROM_FAIL_AU7510 _
               Or AlcorMPMessage = WM_FT_RW_CE_FAIL_AU7510 _
                 Or AlcorMPMessage = WM_FT_PASS_AU7510 _
                 Or AlcorMPMessage = WM_FT_WRITE_PROTECT_FAIL_AU7510 _
                 Or AlcorMPMessage = WM_FT_GPIO_FAIL_AU7510 _
              Or PassTime > 20 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
          MPTester.Print "RW work Time="; PassTime
          MPTester.MPText.Text = Hex(AlcorMPMessage)
        
        
        '===========================================================
        '  RW Time Out Fail
        '===========================================================
        
        If PassTime > 20 Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Time Out Fail"
   
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_CHIP_UNKNOW_FAIL_AU7510
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_SATA_FAIL_AU7510
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_CE_FAIL_AU7510, WM_FT_WRITE_PROTECT_FAIL_AU7510, WM_FT_GPIO_FAIL_AU7510
             TestResult = "Bin3"
             If AlcorMPMessage = WM_FT_RW_CE_FAIL_AU7510 Then
               MPTester.TestResultLab = "Bin3:RW FAIL "
             ElseIf AlcorMPMessage = WM_FT_WRITE_PROTECT_FAIL_AU7510 Then
               MPTester.TestResultLab = "Bin3:WRITE PROTECT FAIL "
             End If
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL_AU7510
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL_AU7510
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_PASS_AU7510
        
 Dim AU7510NoDramUpLimit As Integer
 Dim AU7510WithDramUpLimit As Integer
 Dim i As Long
 Dim TmpHSum As Single
 Dim TmpHAvg As Single
 Dim TmpLSum As Single
 Dim TmpLAvg As Single
 
 AU7510WithDramUpLimit = 220
 AU7510NoDramUpLimit = 220
              ' to measure currennt
 'Dim i As Integer
        
 '====================================
 '   A2D begin
 '====================================
   


 
' dma2205.Show
' dma2205.ShowDataAU6476 0, 100
' If NewChipFlag <> 1 Then
 Dim TmpH As Single
 Dim TmpL As Single
 Dim TmpDiff As Single
 Dim TmpDiffSum As Single
 Dim TmpDiffAvg As Single
 Dim Interval As Single
 Dim DiffCounter As Integer
 Interval = 5

 For i = 0 To 99
     
     
     TmpH = CSng(ADCBuffer(i * 2) + 1) / CSng(32768) * 5# ' ma
     TmpL = CSng(CSng(ADCBuffer(i * 2 + 1)) + 1) / CSng(32768) * 5#
     TmpHSum = TmpHSum + TmpH
     TmpLSum = TmpLSum + TmpL
     TmpDiff = (TmpH - TmpL) * 1000
    ' 'debug.print i; "H"; ADCBuffer(i * 2); TmpH
     ''debug.print i; "L"; ADCBuffer(i * 2 + 1); TmpL
     ' 'debug.print i; TmpH; TmpL; TmpDiff
      If TmpDiff > 0 Then
        DiffCounter = DiffCounter + 1
        TmpDiffSum = TmpDiffSum + TmpDiff
       If TmpDiff > AU7510NoDramUpLimit Or TmpH < 3 Then
                  TestResult = "Bin3"
                    MPTester.TestResultLab = "current:" & CStr(TmpDiff)
                  
                   ContFail = ContFail + 1
        Exit Sub
     End If
     End If
 
 Next i
      'debug.print "1avg"; TmpHSum * 0.01; TmpLSum * 0.01;
      If DiffCounter <> 0 Then
        TmpDiffAvg = CSng(TmpDiffSum / DiffCounter)
        'debug.print "Diffavg"; TmpDiffAvg
        If TmpDiffAvg > 220 Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "current:" & CStr(TmpDiff)
            ContFail = ContFail + 1
            Exit Sub
        End If
          
      End If
' End If
        
        
               
          '     For LedCount = 1 To 20
          '     Call MsecDelay(0.1)
          '     cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
          '      If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
               
          '       Exit For
          '     End If
           '    Next LedCount
                 
           '       MPTester.Print "light="; LightOn
           '      If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
                     MPTester.TestResultLab = "PASS "
                     
                      MPTester.Print "Current="; (TmpHSum - TmpLSum) * 1000
                     TestResult = "PASS"
                     ContFail = 0
           '     Else
                 
           '       TestResult = "Bin3"
           '       MPTester.TestResultLab = "Bin3:LED FAIL "
              
           '    End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
        If TestResult <> "PASS" Then
            Call StartScan_Click_AU7510  ' begin scan
        End If
        'Call PowerSet(1500)
        
         
                            
End Sub
Public Sub AU7511A45AGF20TestSub()
'add unload driver function
 If PCI7248InitFinish = 0 Then
       PCI7248Exist
 End If
 Dim OldTimer
 Dim PassTime
 Dim rt2
 Dim LightOn
 Dim mMsg As MSG
 Dim LedCount As Byte
 
   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================
 
 'AU7510 do not have filter driver
 
 
     
                     


NewChipFlag = 0
If OldChipName <> ChipName Then

' reset program

                        winHwnd = FindWindow(vbNullString, "FT - 7510 - MAG")
 
                         If winHwnd <> 0 Then
                           Do
                           rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                           Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "FT - 7510 - MAG")
                           Loop While winHwnd <> 0
                         End If

           ' FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & chipname & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
           ' FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & chipname & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
           ' FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & chipname & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

 
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail


 '====================================
 '  Fix Card
 '====================================
 
 If NewChipFlag = 1 Or FindWindow(vbNullString, "FT - 7510 - MAG") = 0 Then
    
 '==============================================================
' when begin  scan + MP
'===============================================================
  
 
   
       '  power on
     '  cardresult = DO_WritePort(card, Channel_P1A, &HFF)
     '  Call PowerSet(3)   ' close power to disable chip
     '  Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU7510MAG
       
 
        OldTimer = Timer
        AlcorMPMessage = 0
        
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_RW_READY_AU7510 Or PassTime > 30 _
              Or AlcorMPMessage = WM_FT_USB_UNKNOW_FAIL_AU7510 _
              Or AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3, scan fail
        '====================================================
        '
       
          '(1) TIme out : scan again ; shut down and restart
          '(2) usb fail : scan
          '(3) scan fail: delay and recan, after scan two times fail then restart PC
               
      '========== initial stage
             '(1) time out fail stage : restart
            ' If PassTime > 30 Then
             
             '   MsgBox "FT program no response"
             '   Exit Sub
           ' End If
                ' initial time out fail
               If PassTime > 30 Then    'usb issue so when time out , we let restart PC
               
               'restart PC
                        MPTester.TestResultLab = "Bin3:MP Scan Fail"
                        TestResult = "Bin3"
                        MPTester.Print "Scan Fail"
                        ResetFlag = 1
               
               End If
             
             '(2) initial fail :usb fail stage
      
             If AlcorMPMessage = WM_FT_USB_UNKNOW_FAIL_AU7510 Or AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 Then
               
                
SCAN_LABEL:
                Call StartScan_Click_AU7510MAG  ' begin scan
                
               
                OldTimer = Timer
                AlcorMPMessage = 0
        
                Do
                   ' DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                         AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                    DispatchMessage mMsg
                 End If
                    PassTime = Timer - OldTimer
                Loop Until AlcorMPMessage = WM_FT_SCAN_DONE_AU7510 _
                Or AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 _
                Or PassTime > 30 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
       
                
                
                
                MPTester.Print "MP work time="; PassTime
                MPTester.MPText.Text = Hex(AlcorMPMessage)
                
                
                '(21) time out : restart PC
                  If PassTime > 30 Then  'usb issue so when time out , we let restart PC
                        MPTester.TestResultLab = "Bin3:MP Scan Fail"
                        TestResult = "Bin3"
                        MPTester.Print "Scan Fail"
                        ResetFlag = 1
                  End If
                  
                 '(22) usb fail : rescan
                  If AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 Then
                  
                    ScanContFail = ScanContFail + 1
                    If ScanContFail > 3 Then
                        
                        MPTester.TestResultLab = "Bin3:MP Scan Fail"
                        TestResult = "Bin3"
                        MPTester.Print "Scan Fail"
             
                      Exit Sub
                    End If
                    
                    Call MsecDelay(5#)
                     GoTo SCAN_LABEL
                End If
                
                
                '(23) pass
                   
                
                   If AlcorMPMessage = WM_FT_SCAN_DONE_AU7510 Then
                  
                       ScanContFail = 0
                    
                        
                        MPTester.TestResultLab = "SCAN PASS"
                       
                        MPTester.Print "Scan PASS"
             
                   
                   End If
          
               
        End If
 End If
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_RW_READY_AU7510 Or MPFlag = 1 Or ContFail >= 5 Or MPTester.Check1.Value = 1 Then
           
         MPFlag = 1
             ScanContFail = 0
       
             
            ' Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU7510MAG
   
             
              OldTimer = Timer
              AlcorMPMessage = 0
        
                Do
                   ' DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                         AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                    DispatchMessage mMsg
                 End If
                    PassTime = Timer - OldTimer
                Loop Until AlcorMPMessage = WM_FT_MP_DONE_AU7510 _
                Or AlcorMPMessage = WM_FT_MP_FAIL_AU7510 _
                Or PassTime > 90 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               '(31) MP time out fail, close FT program
                If PassTime > 90 Then   ' this is chip issue, so do not reset PC
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    ' (1)
                  
                 
                    Exit Sub
                End If
                
                '(32) MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL_AU7510 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    If MPContFail > 5 Then
                      
                    ResetFlag = 1
                    End If
                 
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
               
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_DONE_AU7510 Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPFlag = 0
                    MPTester.Print "MP PASS"
                End If
                
              
                
      End If
  

                        
 '=========================================
 '    POWER on
 '=========================================
 
 
      
        
        
        MPTester.Print "wait for RW Tester Ready"
        OldTimer = Timer
        AlcorMPMessage = 0
        Do
            'DoEvents
             If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
             End If
        
             PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_RW_READY_AU7510 Or PassTime > 15 _
        Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        MPTester.Print "RW Ready Time="; PassTime
        
       If PassTime > 15 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
          
            Exit Sub
       End If
         
 Dim ADCFlag As Byte
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU7510MAG
        ADCFlag = 0
        Do
             If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
             
            PassTime = Timer - OldTimer
            If ADCFlag = 0 Then
                MPTester.TestResultLab = "ADC begin"
         '   Call MsecDelay(1)
            
             Call DACGetData2ChannelUpTriggleAU3152HL
             ADCFlag = 1
            End If
        Loop Until AlcorMPMessage = WM_FT_CHIP_UNKNOW_FAIL_AU7510 _
              Or AlcorMPMessage = WM_FT_SATA_FAIL_AU7510 _
              Or AlcorMPMessage = WM_FT_RW_RAM_FAIL_AU7510 _
               Or AlcorMPMessage = WM_FT_RW_ROM_FAIL_AU7510 _
               Or AlcorMPMessage = WM_FT_RW_CE_FAIL_AU7510 _
                 Or AlcorMPMessage = WM_FT_PASS_AU7510 _
                 Or AlcorMPMessage = WM_FT_WRITE_PROTECT_FAIL_AU7510 _
                 Or AlcorMPMessage = WM_FT_GPIO_FAIL_AU7510 _
              Or PassTime > 20 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
          MPTester.Print "RW work Time="; PassTime
          MPTester.MPText.Text = Hex(AlcorMPMessage)
        
        
        '===========================================================
        '  RW Time Out Fail
        '===========================================================
        
        If PassTime > 20 Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Time Out Fail"
   
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_CHIP_UNKNOW_FAIL_AU7510
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_SATA_FAIL_AU7510
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_CE_FAIL_AU7510, WM_FT_WRITE_PROTECT_FAIL_AU7510, WM_FT_GPIO_FAIL_AU7510
             TestResult = "Bin3"
             If AlcorMPMessage = WM_FT_RW_CE_FAIL_AU7510 Then
               MPTester.TestResultLab = "Bin3:RW FAIL "
             ElseIf AlcorMPMessage = WM_FT_WRITE_PROTECT_FAIL_AU7510 Then
               MPTester.TestResultLab = "Bin3:WRITE PROTECT FAIL "
             End If
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL_AU7510
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL_AU7510
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_PASS_AU7510
        
 Dim AU7510NoDramUpLimit As Integer
 Dim AU7510WithDramUpLimit As Integer
 Dim i As Long
 Dim TmpHSum As Single
 Dim TmpHAvg As Single
 Dim TmpLSum As Single
 Dim TmpLAvg As Single
 
 AU7510WithDramUpLimit = 220
 AU7510NoDramUpLimit = 220
              ' to measure currennt
 'Dim i As Integer
        
 '====================================
 '   A2D begin
 '====================================
   


 
' dma2205.Show
' dma2205.ShowDataAU6476 0, 100
' If NewChipFlag <> 1 Then
 Dim TmpH As Single
 Dim TmpL As Single
 Dim TmpDiff As Single
 Dim TmpDiffSum As Single
 Dim TmpDiffAvg As Single
 Dim Interval As Single
 Dim DiffCounter As Integer
 Interval = 5

 For i = 0 To 99
     
     
     TmpH = CSng(ADCBuffer(i * 2) + 1) / CSng(32768) * 5# ' ma
     TmpL = CSng(CSng(ADCBuffer(i * 2 + 1)) + 1) / CSng(32768) * 5#
     TmpHSum = TmpHSum + TmpH
     TmpLSum = TmpLSum + TmpL
     TmpDiff = (TmpH - TmpL) * 1000
    ' 'debug.print i; "H"; ADCBuffer(i * 2); TmpH
     ''debug.print i; "L"; ADCBuffer(i * 2 + 1); TmpL
     ' 'debug.print i; TmpH; TmpL; TmpDiff
      If TmpDiff > 0 Then
        DiffCounter = DiffCounter + 1
        TmpDiffSum = TmpDiffSum + TmpDiff
       If TmpDiff > AU7510NoDramUpLimit Or TmpH < 3 Then
                  TestResult = "Bin3"
                    MPTester.TestResultLab = "current:" & CStr(TmpDiff)
                  
                   ContFail = ContFail + 1
        Exit Sub
     End If
     End If
 
 Next i
      'debug.print "1avg"; TmpHSum * 0.01; TmpLSum * 0.01;
      If DiffCounter <> 0 Then
        TmpDiffAvg = CSng(TmpDiffSum / DiffCounter)
        'debug.print "Diffavg"; TmpDiffAvg
        If TmpDiffAvg > 220 Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "current:" & CStr(TmpDiff)
            ContFail = ContFail + 1
            Exit Sub
        End If
          
      End If
' End If
        
        
               
          '     For LedCount = 1 To 20
          '     Call MsecDelay(0.1)
          '     cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
          '      If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
               
          '       Exit For
          '     End If
           '    Next LedCount
                 
           '       MPTester.Print "light="; LightOn
           '      If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
                     MPTester.TestResultLab = "PASS "
                     
                      MPTester.Print "Current="; (TmpHSum - TmpLSum) * 1000
                     TestResult = "PASS"
                     ContFail = 0
           '     Else
                 
           '       TestResult = "Bin3"
           '       MPTester.TestResultLab = "Bin3:LED FAIL "
              
           '    End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
        If TestResult <> "PASS" Then
            Call StartScan_Click_AU7510MAG  ' begin scan
        End If
        'Call PowerSet(1500)
        
         
                            
End Sub
Public Sub AU7510A45BGF20TestSub()
'add unload driver function
 If PCI7248InitFinish = 0 Then
       PCI7248Exist
 End If
 Dim OldTimer
 Dim PassTime
 Dim rt2
 Dim LightOn
 Dim mMsg As MSG
 Dim LedCount As Byte
 
   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================
 
 'AU7510 do not have filter driver
 
 
     
                     


NewChipFlag = 0
If OldChipName <> ChipName Then

' reset program

                        winHwnd = FindWindow(vbNullString, "FT - 7510 - MBG")
 
                         If winHwnd <> 0 Then
                           Do
                           rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                           Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "FT - 7510 - MBG")
                           Loop While winHwnd <> 0
                         End If

           ' FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & chipname & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
           ' FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & chipname & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
           ' FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & chipname & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

 
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail


 '====================================
 '  Fix Card
 '====================================
 
 If NewChipFlag = 1 Or FindWindow(vbNullString, "FT - 7510 - MBG") = 0 Then
    
 '==============================================================
' when begin  scan + MP
'===============================================================
  
 
   
       '  power on
     '  cardresult = DO_WritePort(card, Channel_P1A, &HFF)
     '  Call PowerSet(3)   ' close power to disable chip
     '  Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU7510MBG
       
 
        OldTimer = Timer
        AlcorMPMessage = 0
        
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_RW_READY_AU7510 Or PassTime > 30 _
              Or AlcorMPMessage = WM_FT_USB_UNKNOW_FAIL_AU7510 _
              Or AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3, scan fail
        '====================================================
        '
       
          '(1) TIme out : scan again ; shut down and restart
          '(2) usb fail : scan
          '(3) scan fail: delay and recan, after scan two times fail then restart PC
               
      '========== initial stage
             '(1) time out fail stage : restart
            ' If PassTime > 30 Then
             
             '   MsgBox "FT program no response"
             '   Exit Sub
           ' End If
                ' initial time out fail
               If PassTime > 30 Then    'usb issue so when time out , we let restart PC
               
               'restart PC
                        MPTester.TestResultLab = "Bin3:MP Scan Fail"
                        TestResult = "Bin3"
                        MPTester.Print "Scan Fail"
                        ResetFlag = 1
               
               End If
             
             '(2) initial fail :usb fail stage
      
             If AlcorMPMessage = WM_FT_USB_UNKNOW_FAIL_AU7510 Or AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 Then
               
                
SCAN_LABEL:
                Call StartScan_Click_AU7510MBG  ' begin scan
                
               
                OldTimer = Timer
                AlcorMPMessage = 0
        
                Do
                   ' DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                         AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                    DispatchMessage mMsg
                 End If
                    PassTime = Timer - OldTimer
                Loop Until AlcorMPMessage = WM_FT_SCAN_DONE_AU7510 _
                Or AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 _
                Or PassTime > 30 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
       
                
                
                
                MPTester.Print "MP work time="; PassTime
                MPTester.MPText.Text = Hex(AlcorMPMessage)
                
                
                '(21) time out : restart PC
                  If PassTime > 30 Then  'usb issue so when time out , we let restart PC
                        MPTester.TestResultLab = "Bin3:MP Scan Fail"
                        TestResult = "Bin3"
                        MPTester.Print "Scan Fail"
                        ResetFlag = 1
                  End If
                  
                 '(22) usb fail : rescan
                  If AlcorMPMessage = WM_FT_HOST_FAIL_AU7510 Then
                  
                    ScanContFail = ScanContFail + 1
                    If ScanContFail > 3 Then
                        
                        MPTester.TestResultLab = "Bin3:MP Scan Fail"
                        TestResult = "Bin3"
                        MPTester.Print "Scan Fail"
             
                      Exit Sub
                    End If
                    
                    Call MsecDelay(5#)
                     GoTo SCAN_LABEL
                End If
                
                
                '(23) pass
                   
                
                   If AlcorMPMessage = WM_FT_SCAN_DONE_AU7510 Then
                  
                       ScanContFail = 0
                    
                        
                        MPTester.TestResultLab = "SCAN PASS"
                       
                        MPTester.Print "Scan PASS"
             
                   
                   End If
          
               
        End If
 End If
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_RW_READY_AU7510 Or MPFlag = 1 Or ContFail >= 5 Or MPTester.Check1.Value = 1 Then
           
         MPFlag = 1
             ScanContFail = 0
       
             
            ' Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU7510MBG
   
             
              OldTimer = Timer
              AlcorMPMessage = 0
        
                Do
                   ' DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                         AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                    DispatchMessage mMsg
                 End If
                    PassTime = Timer - OldTimer
                Loop Until AlcorMPMessage = WM_FT_MP_DONE_AU7510 _
                Or AlcorMPMessage = WM_FT_MP_FAIL_AU7510 _
                Or PassTime > 90 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               '(31) MP time out fail, close FT program
                If PassTime > 90 Then   ' this is chip issue, so do not reset PC
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    ' (1)
                  
                 
                    Exit Sub
                End If
                
                '(32) MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL_AU7510 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    If MPContFail > 5 Then
                      
                    ResetFlag = 1
                    End If
                 
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
               
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_DONE_AU7510 Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPFlag = 0
                    MPTester.Print "MP PASS"
                End If
                
              
                
      End If
  

                        
 '=========================================
 '    POWER on
 '=========================================
 
 
      
        
        
        MPTester.Print "wait for RW Tester Ready"
        OldTimer = Timer
        AlcorMPMessage = 0
        Do
            'DoEvents
             If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
             End If
        
             PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_RW_READY_AU7510 Or PassTime > 15 _
        Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        MPTester.Print "RW Ready Time="; PassTime
        
       If PassTime > 15 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
          
            Exit Sub
       End If
         
 Dim ADCFlag As Byte
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU7510MBG
        ADCFlag = 0
        Do
             If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
             
            PassTime = Timer - OldTimer
            If ADCFlag = 0 Then
                MPTester.TestResultLab = "ADC begin"
         '   Call MsecDelay(1)
            
             Call DACGetData2ChannelUpTriggleAU3152HL
             ADCFlag = 1
            End If
        Loop Until AlcorMPMessage = WM_FT_CHIP_UNKNOW_FAIL_AU7510 _
              Or AlcorMPMessage = WM_FT_SATA_FAIL_AU7510 _
              Or AlcorMPMessage = WM_FT_RW_RAM_FAIL_AU7510 _
               Or AlcorMPMessage = WM_FT_RW_ROM_FAIL_AU7510 _
               Or AlcorMPMessage = WM_FT_RW_CE_FAIL_AU7510 _
                 Or AlcorMPMessage = WM_FT_PASS_AU7510 _
                 Or AlcorMPMessage = WM_FT_WRITE_PROTECT_FAIL_AU7510 _
                 Or AlcorMPMessage = WM_FT_GPIO_FAIL_AU7510 _
              Or PassTime > 20 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
          MPTester.Print "RW work Time="; PassTime
          MPTester.MPText.Text = Hex(AlcorMPMessage)
        
        
        '===========================================================
        '  RW Time Out Fail
        '===========================================================
        
        If PassTime > 20 Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Time Out Fail"
   
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_CHIP_UNKNOW_FAIL_AU7510
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_SATA_FAIL_AU7510
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_CE_FAIL_AU7510, WM_FT_WRITE_PROTECT_FAIL_AU7510, WM_FT_GPIO_FAIL_AU7510
             TestResult = "Bin3"
             If AlcorMPMessage = WM_FT_RW_CE_FAIL_AU7510 Then
               MPTester.TestResultLab = "Bin3:RW FAIL "
             ElseIf AlcorMPMessage = WM_FT_WRITE_PROTECT_FAIL_AU7510 Then
               MPTester.TestResultLab = "Bin3:WRITE PROTECT FAIL "
             End If
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL_AU7510
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL_AU7510
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_PASS_AU7510
        
 Dim AU7510NoDramUpLimit As Integer
 Dim AU7510WithDramUpLimit As Integer
 Dim i As Long
 Dim TmpHSum As Single
 Dim TmpHAvg As Single
 Dim TmpLSum As Single
 Dim TmpLAvg As Single
 
 AU7510WithDramUpLimit = 220
 AU7510NoDramUpLimit = 220
              ' to measure currennt
 'Dim i As Integer
        
 '====================================
 '   A2D begin
 '====================================
   


 
' dma2205.Show
' dma2205.ShowDataAU6476 0, 100
' If NewChipFlag <> 1 Then
 Dim TmpH As Single
 Dim TmpL As Single
 Dim TmpDiff As Single
 Dim TmpDiffSum As Single
 Dim TmpDiffAvg As Single
 Dim Interval As Single
 Dim DiffCounter As Integer
 Interval = 5

 For i = 0 To 99
     
     
     TmpH = CSng(ADCBuffer(i * 2) + 1) / CSng(32768) * 5# ' ma
     TmpL = CSng(CSng(ADCBuffer(i * 2 + 1)) + 1) / CSng(32768) * 5#
     TmpHSum = TmpHSum + TmpH
     TmpLSum = TmpLSum + TmpL
     TmpDiff = (TmpH - TmpL) * 1000
    ' 'debug.print i; "H"; ADCBuffer(i * 2); TmpH
     ''debug.print i; "L"; ADCBuffer(i * 2 + 1); TmpL
     ' 'debug.print i; TmpH; TmpL; TmpDiff
      If TmpDiff > 0 Then
        DiffCounter = DiffCounter + 1
        TmpDiffSum = TmpDiffSum + TmpDiff
       If TmpDiff > AU7510NoDramUpLimit Or TmpH < 3 Then
                  TestResult = "Bin3"
                    MPTester.TestResultLab = "current:" & CStr(TmpDiff)
                  
                   ContFail = ContFail + 1
        Exit Sub
     End If
     End If
 
 Next i
      'debug.print "1avg"; TmpHSum * 0.01; TmpLSum * 0.01;
      If DiffCounter <> 0 Then
        TmpDiffAvg = CSng(TmpDiffSum / DiffCounter)
        'debug.print "Diffavg"; TmpDiffAvg
        If TmpDiffAvg > 220 Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "current:" & CStr(TmpDiff)
            ContFail = ContFail + 1
            Exit Sub
        End If
          
      End If
' End If
        
        
               
          '     For LedCount = 1 To 20
          '     Call MsecDelay(0.1)
          '     cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
          '      If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
               
          '       Exit For
          '     End If
           '    Next LedCount
                 
           '       MPTester.Print "light="; LightOn
           '      If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
                     MPTester.TestResultLab = "PASS "
                     
                      MPTester.Print "Current="; (TmpHSum - TmpLSum) * 1000
                     TestResult = "PASS"
                     ContFail = 0
           '     Else
                 
           '       TestResult = "Bin3"
           '       MPTester.TestResultLab = "Bin3:LED FAIL "
              
           '    End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
        If TestResult <> "PASS" Then
            Call StartScan_Click_AU7510MBG  ' begin scan
        End If
        'Call PowerSet(1500)
        
         
                            
End Sub
Public Sub LoadMP_Click_AU7510()
 Dim TimePass
Dim rt2
' find window
 

 
 
 
winHwnd = FindWindow(vbNullString, "FT")
 
' run program
If winHwnd = 0 Then

Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AU7510FT\" & ChipName & "\SSD_FT.exe", "", "", SW_SHOW)
'Call ShellExecute(0, "open", App.Path & "\AlcorMP_698x_PD\AlcorMP.exe", "", "", SH_SHOW)
 
'Call ShellExecute(Me.hwnd, "open", App.Path & "\AlcorMP.exe", "", "", SH_SHOW)
 
End If

SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
 End Sub
Public Sub LoadMP_Click_AU7510MAG()
 Dim TimePass
 Dim rt2
' find window
 

 
 
 
winHwnd = FindWindow(vbNullString, "FT - 7510 - MAG")
 
' run program
If winHwnd = 0 Then

Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AU7510FT\" & ChipName & "\SSD_FT-MAG.exe", "", "", SW_SHOW)
'Call ShellExecute(0, "open", App.Path & "\AlcorMP_698x_PD\AlcorMP.exe", "", "", SH_SHOW)
 
'Call ShellExecute(Me.hwnd, "open", App.Path & "\AlcorMP.exe", "", "", SH_SHOW)
 
End If

SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
 End Sub
 Public Sub LoadMP_Click_AU7510MBG()
 Dim TimePass
 Dim rt2
' find window
 

 
 
 
winHwnd = FindWindow(vbNullString, "FT - 7510 - MBG")
 
' run program
If winHwnd = 0 Then

Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AU7510FT\" & ChipName & "\SSD_FT-MBG.exe", "", "", SW_SHOW)
'Call ShellExecute(0, "open", App.Path & "\AlcorMP_698x_PD\AlcorMP.exe", "", "", SH_SHOW)
 
'Call ShellExecute(Me.hwnd, "open", App.Path & "\AlcorMP.exe", "", "", SH_SHOW)
 
End If

SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
 End Sub
Public Sub LoadMP_Click_AU7510A43()
 Dim TimePass
Dim rt2
' find window
 

 
 
 
winHwnd = FindWindow(vbNullString, "FT")
 
' run program
If winHwnd = 0 Then

Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AU7510A43FT\" & ChipName & "\SSD_FT.exe", "", "", SW_SHOW)
'Call ShellExecute(0, "open", App.Path & "\AlcorMP_698x_PD\AlcorMP.exe", "", "", SH_SHOW)
 
'Call ShellExecute(Me.hwnd, "open", App.Path & "\AlcorMP.exe", "", "", SH_SHOW)
 
End If

SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
 End Sub
 
 Public Sub StartScan_Click_AU7510()
 Dim rt2
    winHwnd = FindWindow(vbNullString, "FT")
    'debug.print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_SCAN_START_AU7510, 0&, 0&)
 End Sub
 Public Sub StartScan_Click_AU7510MAG()
 Dim rt2
    winHwnd = FindWindow(vbNullString, "FT - 7510 - MAG")
    'debug.print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_SCAN_START_AU7510, 0&, 0&)
 End Sub
 Public Sub StartScan_Click_AU7510MBG()
 Dim rt2
    winHwnd = FindWindow(vbNullString, "FT - 7510 - MBG")
    'debug.print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_SCAN_START_AU7510, 0&, 0&)
 End Sub
 Public Sub StartMP_Click_AU7510MAG()
 Dim rt2
    winHwnd = FindWindow(vbNullString, "FT - 7510 - MAG")
    'debug.print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_MP_START_AU7510, 0&, 0&)
 End Sub
 Public Sub StartMP_Click_AU7510MBG()
 Dim rt2
    winHwnd = FindWindow(vbNullString, "FT - 7510 - MBG")
    'debug.print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_MP_START_AU7510, 0&, 0&)
 End Sub
 Public Sub StartMP_Click_AU7510()
 Dim rt2
    winHwnd = FindWindow(vbNullString, "FT")
    'debug.print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_MP_START_AU7510, 0&, 0&)
 End Sub
Public Sub StartRWTest_Click_AU7510()
Dim rt2
    winHwnd = FindWindow(vbNullString, "FT")
    'debug.print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_RW_START_AU7510, 0&, 0&)
End Sub

Public Sub StartRWTest_Click_AU7510MAG()
Dim rt2
    winHwnd = FindWindow(vbNullString, "FT - 7510 - MAG")
    'debug.print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_RW_START_AU7510, 0&, 0&)
End Sub

Public Sub StartRWTest_Click_AU7510MBG()
Dim rt2
    winHwnd = FindWindow(vbNullString, "FT - 7510 - MBG")
    'debug.print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_RW_START_AU7510, 0&, 0&)
End Sub
