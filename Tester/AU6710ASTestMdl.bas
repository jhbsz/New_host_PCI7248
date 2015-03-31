Attribute VB_Name = "AU6710ASTestMdl"
Public Sub AU6710ASTest()
  If PCI7248InitFinish = 0 Then
                          PCI7248Exist
                          
                          
        End If
           ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)

              CardResult = DO_WritePort(card, Channel_P1A, &H0)
               Call MsecDelay(3.4)
              
                HubPort = 0
                ReaderExist = 0
                rv0 = 0
                ClosePipe
                    'rv1 = AU6610Test
                rv0 = AU6710_GetDevice(0, 1, "6710")
                ClosePipe
                 Call LabelMenu(0, rv0, 1)
                
                 HubPort = 1
                ReaderExist = 0
                rv1 = 0
                ClosePipe
                    'rv1 = AU6610Test
                rv1 = AU6710_GetDevice(0, rv0, "6710")
                ClosePipe
                
                 Call LabelMenu(1, rv1, rv0)
Tester.Print " 1st AU6710 : "; rv0
Tester.Print " 2nd AU6710 : "; rv1


                     If rv0 * rv1 = 1 Then
             
                
                         rv2 = 1
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                         If LightOFF <> 254 Then
                             
                            rv2 = 2
                         End If
                     Else
                        rv2 = 4
                     End If
                        
                Call LabelMenu(2, rv2, rv1)
                
        
                
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
                ElseIf rv2 * rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub


Public Function AU6710_GetDevice(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 1024
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        AU6710_GetDevice = 4
        Exit Function
    End If
    '========================================
   
    AU6710_GetDevice = 0
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
            TmpString = GetDeviceNameMulti6710(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      AU6710_GetDevice = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
 
     
    AU6710_GetDevice = 1
        
    
    End Function
