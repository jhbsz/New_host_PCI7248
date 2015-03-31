Attribute VB_Name = "AU9368MDL"

Public Function CBWTest_NewAU9368Sorting(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_NewAU9368Sorting = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_NewAU9368Sorting = 0
    If LBA > 25 * 1024 Then
        LBA = 0
    End If
    '========================================
     TmpString = ""
    If ReaderExist = 0 Then
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      CBWTest_NewAU9368Sorting = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
   
    
    
    CBWTest_NewAU9368Sorting = 1
        
    
    End Function

Public Sub AU9368ALFTest()


                If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                End If
                
                
               

                LBA = LBA + 1
                
                OldLBa = LBA
               
                FailPosition = 11
               
                
                   
              Call MsecDelay(1.2)
                rv0 = CBWTest_NewAU9368Sorting(0, 1, "vid_058f")
                
                
                
                If rv0 = 1 Then
                
                  rv0 = 0
                  ReaderExist = 0
                 'MOS low
                 CardResult = DO_WritePort(card, Channel_P1CL, &HF)
                 Call MsecDelay(0.8)
                 CardResult = DO_WritePort(card, Channel_P1CL, &H0)
                
                 Call MsecDelay(1.2)
                
                 ClosePipe
                 rv0 = CBWTest_New(0, 1, "vid_058f")
                 ClosePipe
                
                
                End If
                
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New(1, rv0, "vid_058f")
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
                rv2 = CBWTest_New(2, rv1, "vid_058f")
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
               
                rv3 = CBWTest_New_ALPS(2, rv2, "vid_058f")
                 ClosePipe
                   LBA = OldLBa
                 If rv3 = 1 Then
                    FailPosition = 12
                   LBA = LBA + CLng(80000)
                  rv3 = CBWTest_New_ALPS(2, rv2, "vid_058f")
                   ClosePipe
                 End If
                  LBA = OldLBa + 1
                 If rv3 = 1 Then
                   FailPosition = 13
                  LBA = LBA + CLng(160000)
                  rv3 = CBWTest_New_ALPS(2, rv2, "vid_058f")
                   ClosePipe
                End If
                
                If rv3 = 1 Then
                rv3 = CBWTest_New(3, rv2, "vid_058f")
                 ClosePipe
                End If
                
                LBA = OldLBa
                Call LabelMenu(3, rv3, rv2)
                ClosePipe
                
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                Tester.Print "FailPosition="; FailPosition
                
                'If rv0 = 1 And rv1 = 1 And rv2 = 1 And rv3 = 1 Then
                   ' TestResult = "PASS"
                'End If
                
                
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
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                
   End Sub
   
   Public Sub AU9368BTest()


                LBA = LBA + 1
                
                OldLBa = LBA
               
                FailPosition = 11
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New(1, rv0, "vid_058f")
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
                rv2 = CBWTest_New(2, rv1, "vid_058f")
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
               
                rv3 = CBWTest_New_ALPS(2, rv2, "vid_058f")
                 ClosePipe
                   LBA = OldLBa
                 If rv3 = 1 Then
                    FailPosition = 12
                   LBA = LBA + CLng(80000)
                  rv3 = CBWTest_New_ALPS(2, rv2, "vid_058f")
                   ClosePipe
                 End If
                  LBA = OldLBa + 1
                 If rv3 = 1 Then
                   FailPosition = 13
                  LBA = LBA + CLng(160000)
                  rv3 = CBWTest_New_ALPS(2, rv2, "vid_058f")
                   ClosePipe
                End If
                
                If rv3 = 1 Then
                rv3 = CBWTest_New(3, rv2, "vid_058f")
                 ClosePipe
                End If
                
                LBA = OldLBa
                Call LabelMenu(3, rv3, rv2)
                ClosePipe
                
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                Tester.Print "FailPosition="; FailPosition
                
                'If rv0 = 1 And rv1 = 1 And rv2 = 1 And rv3 = 1 Then
                   ' TestResult = "PASS"
                'End If
                
                
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
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                
   End Sub
