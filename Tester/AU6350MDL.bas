Attribute VB_Name = "AU6350MDL"
Public Sub AU6350KLF21TestSub()
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                
             
              
              
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on  test anotehr card
                '=========================================
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
            '     CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                 Call MsecDelay(1.2)   'power on time
                 
            
                  
                
                     
              
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                  
                  
                 If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect On Fail"
                    End
                 End If
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     
                       Call MsecDelay(0.1)
                           
                           
                      ChipString = "6366"
                      ClosePipe
                      HubPort = 1
                      ReaderExist = 0
                      rv0 = CBWTest_New_AU6350CF(0, 1, ChipString)
                      
                      
                      If rv0 = 1 Then
                         rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                        If rv0 <> 1 Then
                           rv0 = 2
                           Tester.Print "SD bus width fail"
                        End If
                      End If
                      
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                                   
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        rv6 = 2
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                    End If
                         Call MsecDelay(0.01)
                     
                   
                     
                    
                       
                    
                     
                    
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
                
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H5E)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                 
                
                 Call MsecDelay(0.1)
              
                 
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                
                  Call MsecDelay(0.1)
                 
      
                   OpenPipe
                    rv5 = ReInitial(0)
                    ClosePipe
                
                ClosePipe
                rv5 = CBWTest_New(0, rv0, ChipString)
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               
           
           
                    ClosePipe
                      HubPort = 0
                      ReaderExist = 0
                      
                      
                      rv7 = CBWTest_New_AU6350CF(0, rv5, "6335")
                      ClosePipe
                      Tester.Print rv7, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                       Call LabelMenu(4, rv7, rv5)
           
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
           
                If LightOff <> &HFC Then
                  rv7 = 2
                  Tester.Print rv7, "GPO fail"
             
                End If
           
           
           
                
               '   CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
                           
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Or rv7 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Or rv7 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv0 * rv7 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6350KLF22TestSub()
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                
             
              
              
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on  test anotehr card
                '=========================================
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
            '     CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                 Call MsecDelay(1.2)   'power on time
                 
            
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                  
                  
                 If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect On Fail"
                    End
                 End If
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     
                       Call MsecDelay(0.1)
                           
                           
                      ChipString = "6366"
                      ClosePipe
                      HubPort = 1
                      ReaderExist = 0
                      rv0 = CBWTest_New_AU6350CF(0, 1, ChipString)
                      
                      
                      If rv0 = 1 Then
                         rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                        If rv0 <> 1 Then
                           rv0 = 2
                           Tester.Print "SD bus width fail"
                        End If
                      End If
                      
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                                   
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        rv6 = 2
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                    End If
                         Call MsecDelay(0.01)
                     
                   
                     
                    
                       
                    
                     
                    
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
                
              
                '2010/5/28 purpose to saving re-initial time waste (AU6352JKL)
                
                '****************************************************
                
                 'CardResult = DO_WritePort(card, Channel_P1A, &H5E)
                
                 'If CardResult <> 0 Then
                 '   MsgBox "Set MSPro Card Detect On Fail"
                 '   End
                 'End If
                 
                
                 'Call MsecDelay(0.1)
              
                '****************************************************
                
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                
                  Call MsecDelay(0.1)
                 
      
                   OpenPipe
                    rv5 = ReInitial(0)
                    ClosePipe
                
                ClosePipe
                rv5 = CBWTest_New(0, rv0, ChipString)
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               
           
           
                    ClosePipe
                      HubPort = 0
                      ReaderExist = 0
                      
                      
                      rv7 = CBWTest_New_AU6350CF(0, rv5, "6335")
                      ClosePipe
                      Tester.Print rv7, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                       Call LabelMenu(4, rv7, rv5)
           
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
           
                If LightOff <> &HFC Then
                  rv7 = 2
                  Tester.Print rv7, "GPO fail"
             
                End If
           
           
           
                
               '   CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
                           
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Or rv7 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Or rv7 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv0 * rv7 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6350KLF23TestSub()
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                
             
              
              
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
        
       
                '=========================================
                '    POWER on  test anotehr card
                '=========================================
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
            '     CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                 Call MsecDelay(1.2)   'power on time
                 
            
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                  
                  
                 If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect On Fail"
                    End
                 End If
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================

                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     
                       Call MsecDelay(0.1)
                           
                           
                      ChipString = "6366"
                      ClosePipe
                      HubPort = 0
                      ReaderExist = 0
                      rv0 = CBWTest_New_AU6350CF(0, 1, ChipString)
                      
                      
                      If rv0 = 1 Then
                         rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                        If rv0 <> 1 Then
                           rv0 = 2
                           Tester.Print "SD bus width fail"
                        End If
                      End If
                      
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                                   
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        rv6 = 2
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                    End If
                         Call MsecDelay(0.01)
                     
                   
                     
                    
                       
                    
                     
                    
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
                
              
                '2010/5/28 purpose to saving re-initial time waste (AU6352JKL)
                
                '****************************************************
                
                 'CardResult = DO_WritePort(card, Channel_P1A, &H5E)
                
                 'If CardResult <> 0 Then
                 '   MsgBox "Set MSPro Card Detect On Fail"
                 '   End
                 'End If
                 
                
                 'Call MsecDelay(0.1)
              
                '****************************************************
                
    
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                
                  Call MsecDelay(0.1)
                 
      
                   OpenPipe
                    rv5 = ReInitial(0)
                    ClosePipe
                
                ClosePipe
                rv5 = CBWTest_New(0, rv0, ChipString)
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               

'**************************************************************************************************
'2010/7/14 skip AU650JKL Port1(AU6336 reader) test
'
                rv7 = 1
'
'                    ClosePipe
'                      HubPort = 0
'                      ReaderExist = 0
'
'
'                      rv7 = CBWTest_New_AU6350CF(0, rv5, "6335")
'                      ClosePipe
'                      Tester.Print rv7, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
'                       Call LabelMenu(4, rv7, rv5)
'
'                     CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'
'                   If CardResult <> 0 Then
'                    MsgBox "Read light off fail"
'                    End
'                   End If
'
'                If LightOFF <> &HFC Then
'                  rv7 = 2
'                  Tester.Print rv7, "GPO fail"
'
'                End If
'
'**************************************************************************************************
           
                
               '   CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
                           
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Or rv7 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Or rv7 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv0 * rv7 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6350Test()


                      If ChipName = "AU6350ALF20" Then
                          AU6350DLTestSub
                        End If
          
                        If ChipName = "AU6350BFF20" Then
                          AU6350BFTestSub
                        End If
        
                        If ChipName = "AU6350GLF20" Then
                          AU6350GLTestSub
                        End If
                         
                        If ChipName = "AU6350ALF21" Then
                          AU6350ALF21TestSub
                        End If
                        
                        If ChipName = "AU6350ALF22" Then
                          AU6350ALF22TestSub
                        End If
                        
                        If ChipName = "AU6350KLF21" Then
                          AU6350KLF21TestSub
                        End If
                        
                        If ChipName = "AU6350KLF22" Then
                          AU6350KLF22TestSub
                        End If
                        
                        If ChipName = "AU6350KLF23" Then
                          AU6350KLF23TestSub
                        End If
                        
                        If ChipName = "AU6350BFF21" Then
                           AU6350BFF21TestSub
                        End If
        
                        If ChipName = "AU6350CFF21" Then
                           AU6350CFF21TestSub
                        End If
        
                        If ChipName = "AU6350GLF21" Then
                           AU6350GLF21TestSub
                        End If
                         
                         If ChipName = "AU6350OLF21" Then
                           AU6350OLF21TestSub
                        End If
                         
                          If ChipName = "AU6350BLF21" Then
                           AU6350BLF21TestSub
                        End If
          
                         If ChipName = "AU6350CFF22" Then
                           AU6350CFF22TestSub
                        End If
        
End Sub
Public Sub AU6350BLF21TestSub()

' add SD_Speed_Test at AU6371DLTest29
' add Ram_unstable Bin at Bin5 , the Bin are shrink to Bin2,Bin3,Bin4

     ' AU6371


     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
                If Left(ChipName, 10) = "AU6371ELF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371ELF2"
                End If
                 
                 If Left(ChipName, 10) = "AU6371HLF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371HLF2"
                End If
                
                If ChipName = "AU6371S3" Then
                  ChipString = "vid_8751"
                Else
                  ChipString = "vid_058f"
                End If
               
                If ChipName = "AU6371EL" Or Left(ChipName, 10) = "AU6371GLF2" Then
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    If ChipName = "AU6371EL" Then
                        AU6371EL_BootTime = 1.2
                        If OldChipName = "AU6371ELF2" Then
                         AU6371EL_BootTime = 0.5
                        End If
                        
                    Else
                        AU6371EL_BootTime = 0.02
                    End If
                 Else
                    AU6371EL_SD = 0
                    AU6371EL_CF = 0
                    AU6371EL_XD = 0
                    AU6371EL_MS = 0
                    AU6371EL_MSP = 0
                    AU6371EL_BootTime = 0
                End If
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.05)
                 
                 
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(4.5)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371DLResult
                  
                  End If
                  
                      
                      
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
               '  Call MsecDelay(2.2 + AU6371EL_BootTime)  'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                
                 
                  
                  
                 If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect On Fail"
                    End
                 End If
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                   
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                          If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                          End If
                      End If
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                         For i = 1 To 20
                        
                          If rv0 = 1 Then
                             
                              ClosePipe
                               rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                                 
                               ClosePipe
                           End If
                   
                     
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                   
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                                
                              ClosePipe
                          End If
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                           
                          If rv0 <> 1 Then
                          rv6 = 2  ' ram unstable
                          GoTo AU6371DLResult
                          End If
                             
                          Next
                          
                      End If
                        
                         ' Call MsecDelay(0.01)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                        
                        
                  '   If Left(ChipName, 10) = "AU6371DLF2" Then
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    ' ElseIf Left(ChipName, 7) = "AU6366C" Then
                          
                   '      If rv0 <> 0 Then
                   '       If LightON <> 175 Or LightOFF <> 255 Then
                                    
                   '       UsbSpeedTestResult = GPO_FAIL
                   '       rv0 = 3
                   '       End If
                    '    End If
                      
                    ' Else
                     
                    '      If rv0 <> 0 Then
                    '      If LightON <> &HBF Or LightOFF <> &HBF Then
                                    
                     '     UsbSpeedTestResult = GPO_FAIL
                     '     rv0 = 3
                     '     End If
                     '   End If
                   '  End If
                     
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
                If ChipName <> "AU6371S3" Then
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                  If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect On Fail"
                    End
                  End If
                  
                  
                    Call MsecDelay(0.01)
                 If rv0 = 1 Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                 End If
                 If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect Down Fail"
                    End
                 End If
                  Call MsecDelay(AU6371EL_BootTime * 2)
                 If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
               
                 ClosePipe
                 rv1 = CBWTest_New(0, rv0, ChipString)
                 ClosePipe
  
                  
                 If rv1 = 1 Then
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                       
                         
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                       
                    
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                    
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                      
                       If rv1 <> 1 Then ' Ram unstable
                      
                       rv6 = 2
                       GoTo AU6371DLResult
                       End If
              End If
                
                
                
                 
               Else
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
               End If
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              '    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
              '   If CardResult <> 0 Then
              '      MsgBox "Set SMC Card Detect On Fail"
              '      End
              '   End If
                  
              '   Call MsecDelay(0.01)
              '  If rv1 = 1 Then
              '      CardResult = DO_WritePort(card, Channel_P1A, &H7B)
              '  End If
               
              '  If CardResult <> 0 Then
              '      MsgBox "Set SMC Card Detect Down Fail"
              '      End
              '   End If
                 
              '  ClosePipe
              '  rv2 = CBWTest_New(0, rv1, "vid_058f")
              '  Call LabelMenu(21, rv2, rv1)
                
              '       Tester.print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               If rv1 = 1 Then
                   rv2 = 1   ' to complete the SMC asbolish
               Else
                   rv2 = 0
               End If
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.01)
                If rv2 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                End If
                Call MsecDelay(AU6371EL_BootTime)
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect Down Fail"
                    End
                 End If
                 
                   If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                     rv4 = 1
               
               
                     Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv4 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                End If
                  Call MsecDelay(AU6371EL_BootTime * 2)
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect Down Fail"
                    End
                 End If
                 If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               
                
               
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
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
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "MS Fail"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6350BFTestSub()
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
              
             
                ChipString = "vid_058f"
                 
                   
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  
                 
                  Call MsecDelay(1.5)
                    
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &HFE)
                  
                 Call MsecDelay(2.2)  'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                ClosePipe
                Call LabelMenu(0, rv0, 1)
            
            
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &HF6)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.6)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv1 = CBWTest_New(0, rv0, ChipString)
                      ClosePipe
                      
                      
                      
              
                      
                           
                         
                    
                        If rv1 = 1 Then
                          If LightOn <> 252 Or LightOff <> 254 Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv1 = 3
                          End If
                        End If
                    
                     
                    
                     Call LabelMenu(1, rv1, rv0)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
               
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
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
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv0 * rv1 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6350CFF21TestSub()
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
              
             
                ChipString = "vid_058f"
                 
                   
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  
                 
                  Call MsecDelay(0.2)
                    
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &HFE)
                  
                 Call MsecDelay(2.2)  'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                ClosePipe
                Call LabelMenu(0, rv0, 1)
            
            
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &HF6)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                    
                     
                             
                           
                      ClosePipe
                      
                      
                      rv1 = CBWTest_New(0, rv0, ChipString)
                      
                       If rv0 = 1 Then
                         rv0 = Read_SD_Speed_AU6371(0, 0, 18, "4Bits")
                        If rv0 <> 1 Then
                           rv0 = 2
                           Tester.Print "SD bus width fail"
                        End If
                      End If
                      
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                                   
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        rv6 = 2
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                    End If
                      
                      
                      
                      
                     '   Call MsecDelay(0.8)
              
                      
                             CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                         
                    
                        If rv1 = 1 Then
                          If LightOn <> 254 Or LightOff <> 254 Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv1 = 3
                          End If
                        End If
                    
                     
                    
                     Call LabelMenu(1, rv1, rv0)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
               
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                         ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
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
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv0 * rv1 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6350CFF22TestSub()
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
              
             
                ChipString = "6366"
                 
                   
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  
                 
                  Call MsecDelay(0.4)
                    
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &HFC)
                  
                 Call MsecDelay(2.2)  'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                 ClosePipe
                   HubPort = 1
                      ReaderExist = 0
                  rv0 = CBWTest_New_no_card_AU6350CF(0, 1, "vid_058f")
                 ClosePipe
                 Call LabelMenu(0, rv0, 1)
            
            
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &HF4)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                    
                     
                             
                           
                      ClosePipe
                      HubPort = 1
                  '    ReaderExist = 0
                      
                      
                      rv1 = CBWTest_New_AU6350CF(0, rv0, ChipString)
                      
                       If rv0 = 1 Then
                         rv0 = Read_SD_Speed_AU6371(0, 0, 18, "4Bits")
                        If rv0 <> 1 Then
                           rv0 = 2
                           Tester.Print "SD bus width fail"
                        End If
                      End If
                      
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                                   
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        rv6 = 2
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                    End If
                      
                      
                      
                      
                     '   Call MsecDelay(0.8)
              
                      
                             CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                         
                    
                        If rv1 = 1 Then
                          If LightOn <> 254 Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv1 = 3
                          End If
                        End If
                    
                     
                    
                     Call LabelMenu(1, rv1, rv0)   ' no card test fail
                     
                     
                       ClosePipe
                      HubPort = 0
                      ReaderExist = 0
                      
                      
                      rv2 = CBWTest_New_AU6350CF(0, rv1, "6335")
                      ClosePipe
                     
                       Call LabelMenu(2, rv2, rv1)
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv2, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
               
                
                
                
                
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                         ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
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
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv0 * rv1 * rv2 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6350BFF21TestSub()
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
              
             
                ChipString = "vid_058f"
                 
                   
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  
                 
                  Call MsecDelay(0.2)
                    
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &HFE)
                  
                 Call MsecDelay(2.2)  'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                ClosePipe
                Call LabelMenu(0, rv0, 1)
            
            
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &HF6)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                    
                     
                             
                           
                      ClosePipe
                      
                      
                      rv1 = CBWTest_New(0, rv0, ChipString)
                      
                       If rv0 = 1 Then
                         rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                        If rv0 <> 1 Then
                           rv0 = 2
                           Tester.Print "SD bus width fail"
                        End If
                      End If
                      
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                                   
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        rv6 = 2
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                    End If
                      
                      
                      
                      
                     '   Call MsecDelay(0.8)
              
                      
                             CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                         
                    
                        If rv1 = 1 Then
                          If LightOn <> 252 Or LightOff <> 254 Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv1 = 3
                          End If
                        End If
                    
                     
                    
                     Call LabelMenu(1, rv1, rv0)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
               
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                         ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
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
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv0 * rv1 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub



Public Sub AU6350GLTestSubOld()
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
              
             
                ChipString = "vid_058f"
                 
                   
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  
                 
                  Call MsecDelay(1.5)
                    
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &HFE)
                  
                 Call MsecDelay(2.2)  'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                ClosePipe
                Call LabelMenu(0, rv0, 1)
            
            
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &HF6)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.6)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv1 = CBWTest_New(0, rv0, ChipString)
                      ClosePipe
                      
                      
                      
              
                      
                           
                         
                    
                        If rv1 = 1 Then
                          If LightOn <> 252 Or LightOff <> 254 Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv1 = 3
                          End If
                        End If
                    
                     
                    
                     Call LabelMenu(1, rv1, rv0)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
               
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
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
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv0 * rv1 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6350ALF21TestSub()
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                
               
                OldChipName = ""
                If Left(ChipName, 10) = "AU6371ELF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371ELF2"
                End If
                 
                 If Left(ChipName, 10) = "AU6371HLF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371HLF2"
                End If
                
                If ChipName = "AU6371S3" Then
                  ChipString = "vid_8751"
                Else
                  ChipString = "vid_058f"
                End If
               ChipString = "vid_1984"
                If ChipName = "AU6371EL" Or Left(ChipName, 10) = "AU6371GLF2" Then
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    If ChipName = "AU6371EL" Then
                        AU6371EL_BootTime = 1.2
                        If OldChipName = "AU6371ELF2" Then
                         AU6371EL_BootTime = 0.5
                        End If
                        
                    Else
                        AU6371EL_BootTime = 0.02
                    End If
                 Else
                    AU6371EL_SD = 0
                    AU6371EL_CF = 0
                    AU6371EL_XD = 0
                    AU6371EL_MS = 0
                    AU6371EL_MSP = 0
                    AU6371EL_BootTime = 0
                End If
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  
                 
                  Call MsecDelay(0.01)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
              '   Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
                  Call MsecDelay(1.2)
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                  
                  
                 If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect On Fail"
                    End
                 End If
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                    
                     
                       Call MsecDelay(0.8)
                           
                      ClosePipe
                  
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      
                      If rv0 = 1 Then
                         rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                        If rv0 <> 1 Then
                           rv0 = 2
                           Tester.Print "SD bus width fail"
                        End If
                      End If
                      
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                                   
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        rv6 = 2
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                    End If
                         Call MsecDelay(0.01)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                    
                        If rv0 <> 0 Then
                          If LightOn <> 63 Or LightOff <> 127 Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
                    
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
                If ChipName <> "AU6371S3" Then
                  CardResult = DO_WritePort(card, Channel_P1A, &H7C)
                  
                  If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect On Fail"
                    End
                  End If
                  
                  
                    Call MsecDelay(0.01)
                
                     CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                 
                   Call MsecDelay(0.01)
                   OpenPipe
                   rv1 = ReInitial(0)
                   ClosePipe
                   
                   
               
                 ClosePipe
                 rv1 = CBWTest_New(0, rv0, ChipString)
                 ClosePipe
  
                  
                 If rv1 = 1 Then
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                
                 If rv1 <> 1 Then
                   rv6 = 2
                   GoTo AU6371DLResult
                 End If
               
                 End If
                 
               Else
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
               End If
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
                '  SMC Card test  : stop these test for card not enough
                '================================================
                   CardResult = DO_WritePort(card, Channel_P1A, &H75) '1001 CF + SMC
               
                    Call MsecDelay(0.01)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H7B)
                    Call MsecDelay(0.01)
               
                  If CardResult <> 0 Then
                     MsgBox "Set SMC Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv2 = ReInitial(0)
                    ClosePipe
                 
                 ClosePipe
                 rv2 = CBWTest_New(0, rv1, "vid_058f")
                 Call LabelMenu(21, rv2, rv1)
                
                     Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H73) '0011 XD + SMC
               
                     Call MsecDelay(0.01)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                     Call MsecDelay(0.01)
               
                  If CardResult <> 0 Then
                     MsgBox "Set XD Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv3 = ReInitial(0)
                    ClosePipe
               
                  
                 
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
               
                '===============================================
                '  MS Pro Card test
                '================================================
                   rv4 = rv3  ' for MS
                
                
                   CardResult = DO_WritePort(card, Channel_P1A, &H57) ' XD + MSPro
               
                    Call MsecDelay(0.01)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                    Call MsecDelay(0.01)
               
                  If CardResult <> 0 Then
                     MsgBox "Set MSpro Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv5 = ReInitial(0)
                    ClosePipe
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
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
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "MS Fail"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
                
            
End Sub

Public Sub AU6350ALF22TestSub()
     
Dim ChipString As String
Dim i As Integer
Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single
                
               
    OldChipName = ""
                
    ChipString = "vid_1984"
                
    AU6371EL_SD = 0
    AU6371EL_CF = 0
    AU6371EL_XD = 0
    AU6371EL_MS = 0
    AU6371EL_MSP = 0
    AU6371EL_BootTime = 0
            
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
                
    LBA = LBA + 1
                         
    rv0 = 0
    rv1 = 0
    rv2 = 0
    rv3 = 0
    rv4 = 0
    rv5 = 0
    rv6 = 0
    rv7 = 0
             
    Tester.Label3.BackColor = RGB(255, 255, 255)
    Tester.Label4.BackColor = RGB(255, 255, 255)
    Tester.Label5.BackColor = RGB(255, 255, 255)
    Tester.Label6.BackColor = RGB(255, 255, 255)
    Tester.Label7.BackColor = RGB(255, 255, 255)
    Tester.Label8.BackColor = RGB(255, 255, 255)
                
    '=========================================
    '    POWER on
    '=========================================
                  
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(1.2)
        
        'If CardResult <> 0 Then
        '    MsgBox "Power off fail"
        '    End
        'End If
                
        'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                  
        'If CardResult <> 0 Then
        '    MsgBox "Read light off fail"
        '    End
        'End If
                 
                         
        CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
        '   Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
        Call MsecDelay(1.8)
                
    '===============================================
    '  SD Card test
    '================================================
        CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
        Call MsecDelay(0.01)
                  
        If CardResult <> 0 Then
            MsgBox "Read light off fail"
            End
        End If
                 
                 
    '===========================================
    '   NO card test
    '============================================
  
        ' set SD card detect down
        CardResult = DO_WritePort(card, Channel_P1A, &H7E)
        Call MsecDelay(0.2)
                      
        If CardResult <> 0 Then
            MsgBox "Set SD Card Detect Down Fail"
            End
        End If
                     
        'ClosePipe
                  
        rv0 = CBWTest_New(0, 1, ChipString)
                      
        If rv0 = 1 Then
            rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
            ClosePipe
            If rv0 <> 1 Then
                rv0 = 2
                Tester.Print "SD bus width fail"
            End If
        End If
        
        'ClosePipe
                      
        If rv0 = 1 Then
            For i = 1 To 20
                      
                If rv0 = 1 Then
                    'ClosePipe
                    rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                    ClosePipe
                End If
                
                If rv0 = 1 Then
                    'ClosePipe
                    rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                    ClosePipe
                End If
                 
              
                If rv0 = 1 Then
                    'ClosePipe
                    rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                    ClosePipe
                End If
              
                If rv0 = 1 Then
                    'ClosePipe
                    rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                    ClosePipe
                End If
              
                If rv0 <> 1 Then
                    rv6 = 2
                    GoTo AU6371DLResult
                End If
                           
            Next
        End If
                         
                     
        CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
        Call MsecDelay(0.01)
        
        If CardResult <> 0 Then
            MsgBox "Read light On fail"
            End
        End If
                     
                    
        If rv0 <> 0 Then
            If LightOn <> 63 Or LightOff <> 127 Then
                UsbSpeedTestResult = GPO_FAIL
                rv0 = 3
            End If
        End If
                    
        Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
        Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    
    '===============================================
    '  CF Card test
    '================================================
        
            
            CardResult = DO_WritePort(card, Channel_P1A, &H7C)
            Call MsecDelay(0.01)
                  
            If CardResult <> 0 Then
                MsgBox "Set CF Card Detect On Fail"
                End
            End If
                  
                  
                
            CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                 
            Call MsecDelay(0.01)
            
            If rv0 = 1 Then
                OpenPipe
                rv1 = ReInitial(0)
                ClosePipe
            End If
               
            'ClosePipe
            rv1 = CBWTest_New(0, rv0, ChipString)
            ClosePipe
  
                  
                  
            If rv1 = 1 Then
                'ClosePipe
                rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
            End If
                   
            If rv1 = 1 Then
                'ClosePipe
                rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
            End If
              
            If rv1 = 1 Then
                'ClosePipe
                rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
            End If
              
            If rv1 = 1 Then
                'ClosePipe
                rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
            End If
              
            If rv1 <> 1 Then
                rv6 = 2
                GoTo AU6371DLResult
            End If
               
                 
            Call LabelMenu(1, rv1, rv0)
            
            Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '================================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
                   
        CardResult = DO_WritePort(card, Channel_P1A, &H75) '1001 CF + SMC
        Call MsecDelay(0.01)
                
        CardResult = DO_WritePort(card, Channel_P1A, &H7B)
        Call MsecDelay(0.01)
               
        If CardResult <> 0 Then
            MsgBox "Set SMC Card Detect Down Fail"
            End
        End If
                 
        If rv1 = 1 Then
            OpenPipe
            rv2 = ReInitial(0)
            ClosePipe
        End If
                 
        'ClosePipe
        rv2 = CBWTest_New(0, rv1, "vid_058f")
        ClosePipe
        Call LabelMenu(21, rv2, rv1)
                
        Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  XD Card test
    '================================================
        
        CardResult = DO_WritePort(card, Channel_P1A, &H73) '0011 XD + SMC
        Call MsecDelay(0.01)
                
        CardResult = DO_WritePort(card, Channel_P1A, &H77)
        Call MsecDelay(0.01)
               
        If CardResult <> 0 Then
            MsgBox "Set XD Card Detect Down Fail"
            End
        End If
        
        If rv2 = 1 Then
            OpenPipe
            rv3 = ReInitial(0)
            ClosePipe
        End If
                 
        rv3 = CBWTest_New(0, rv2, ChipString)
        ClosePipe
        Call LabelMenu(2, rv3, rv2)
                 
        Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
               
    '===============================================
    '  MS Pro Card test
    '================================================
        
        rv4 = rv3  ' for MS
                
        CardResult = DO_WritePort(card, Channel_P1A, &H57) ' XD + MSPro
        Call MsecDelay(0.01)
                
        CardResult = DO_WritePort(card, Channel_P1A, &H5F)
        Call MsecDelay(0.01)
               
        If CardResult <> 0 Then
            MsgBox "Set MSpro Card Detect Down Fail"
            End
        End If
                 
        If rv4 = 1 Then
            OpenPipe
            rv5 = ReInitial(0)
            ClosePipe
        End If
        
        'ClosePipe
        rv5 = CBWTest_New(0, rv4, ChipString)
        Call LabelMenu(31, rv5, rv4)
        Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
        ClosePipe
                
        
                
AU6371DLResult:
        fnScsi2usb2K_KillEXE
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
        Call MsecDelay(0.01)
        
        
        If rv0 = UNKNOW Then
            UnknowDeviceFail = UnknowDeviceFail + 1
            TestResult = "UNKNOW"
        ElseIf rv6 = 2 Then
            MSWriteFail = MSWriteFail + 1
            TestResult = "MS_WF"
            Tester.Label9.Caption = "ram unsatble Fail"
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
            XDWriteFail = XDWriteFail + 1
            TestResult = "XD_WF"
            Tester.Label9.Caption = "XD Fail"
        ElseIf rv4 = READ_FAIL Then
            XDReadFail = XDReadFail + 1
            TestResult = "XD_RF"
            Tester.Label9.Caption = "XD Fail"
       ElseIf rv5 = WRITE_FAIL Then
            XDWriteFail = XDWriteFail + 1
            TestResult = "XD_WF"
            Tester.Label9.Caption = "MS Fail"
        ElseIf rv5 = READ_FAIL Then
            XDReadFail = XDReadFail + 1
            TestResult = "XD_RF"
            Tester.Label9.Caption = "MS Fail"
        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
            TestResult = "PASS"
        Else
            TestResult = "Bin2"
        End If
                
                
 End Sub
Public Sub AU6350DLTestSub()
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                
               
                OldChipName = ""
                If Left(ChipName, 10) = "AU6371ELF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371ELF2"
                End If
                 
                 If Left(ChipName, 10) = "AU6371HLF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371HLF2"
                End If
                
                If ChipName = "AU6371S3" Then
                  ChipString = "vid_8751"
                Else
                  ChipString = "vid_058f"
                End If
               ChipString = "vid_1984"
                If ChipName = "AU6371EL" Or Left(ChipName, 10) = "AU6371GLF2" Then
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    If ChipName = "AU6371EL" Then
                        AU6371EL_BootTime = 1.2
                        If OldChipName = "AU6371ELF2" Then
                         AU6371EL_BootTime = 0.5
                        End If
                        
                    Else
                        AU6371EL_BootTime = 0.02
                    End If
                 Else
                    AU6371EL_SD = 0
                    AU6371EL_CF = 0
                    AU6371EL_XD = 0
                    AU6371EL_MS = 0
                    AU6371EL_MSP = 0
                    AU6371EL_BootTime = 0
                End If
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  
                 
                 ' Call MsecDelay(2)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                 Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                  
                  
                 If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect On Fail"
                    End
                 End If
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.01)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                                   
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                    
                        If rv0 <> 0 Then
                          If LightOn <> 63 Or LightOff <> 127 Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
                    
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
                If ChipName <> "AU6371S3" Then
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                  If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect On Fail"
                    End
                  End If
                  
                  
                    Call MsecDelay(0.01)
                 If rv0 = 1 Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                 End If
                 If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect Down Fail"
                    End
                 End If
                  Call MsecDelay(AU6371EL_BootTime * 2)
                 If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
               
                 ClosePipe
                 rv1 = CBWTest_New(0, rv0, ChipString)
                 ClosePipe
  
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                
                
                
                 
               Else
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
               End If
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              '    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
              '   If CardResult <> 0 Then
              '      MsgBox "Set SMC Card Detect On Fail"
              '      End
              '   End If
                  
              '   Call MsecDelay(0.01)
              '  If rv1 = 1 Then
              '      CardResult = DO_WritePort(card, Channel_P1A, &H7B)
              '  End If
               
              '  If CardResult <> 0 Then
              '      MsgBox "Set SMC Card Detect Down Fail"
              '      End
              '   End If
                 
              '  ClosePipe
              '  rv2 = CBWTest_New(0, rv1, "vid_058f")
              '  Call LabelMenu(21, rv2, rv1)
                
              '       Tester.print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               If rv1 = 1 Then
                   rv2 = 1   ' to complete the SMC asbolish
               Else
                   rv2 = 0
               End If
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.01)
                If rv2 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                End If
                Call MsecDelay(AU6371EL_BootTime)
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect Down Fail"
                    End
                 End If
                 
                   If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                     rv4 = 1
               
               
                     Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv4 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                End If
                  Call MsecDelay(AU6371EL_BootTime * 2)
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect Down Fail"
                    End
                 End If
                 If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               
                
               
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
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
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub


Public Sub AU6350OLF21TestSub()
'20131028 revise for usb host hang-out issue(like AU6350GLF21)
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim EnumRes As Byte
             
              
              
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
                EnumRes = 0
                
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
            '     CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                  EnumRes = WaitHUBOn("pid_6254")
                 If EnumRes <> 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &HA1)   ' Close power
                    WaitHUBOFF ("pid_6254")
                    Call MsecDelay(0.3)
                    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                    EnumRes = WaitHUBOn("pid_6254")
                 End If
                  
                 'Call MsecDelay(1.2)   'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                    
                
                  
                  
                 If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect On Fail"
                    End
                 End If
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     
                       Call MsecDelay(0.1)
                           
                      ClosePipe
                  
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      
                      If rv0 = 1 Then
                         rv0 = Read_SD_Speed_AU6371(0, 0, 18, "4Bits")
                        If rv0 <> 1 Then
                           rv0 = 2
                           Tester.Print "SD bus width fail"
                        End If
                      End If
                      
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                                   
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        rv6 = 2
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                    End If
                         Call MsecDelay(0.01)
                     
                   
                     
                    
                       
                    
                     
                    
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
                
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H5E)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                 
                
                 Call MsecDelay(0.1)
              
                 
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                
                  Call MsecDelay(0.1)
                 
      
                   OpenPipe
                    rv5 = ReInitial(0)
                    ClosePipe
                
                ClosePipe
                rv5 = CBWTest_New(0, rv0, ChipString)
                   ClosePipe
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               
                 CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
           
                If LightOff <> &HFC Then
                  rv5 = 2
                  Tester.Print rv5, "GPO fail"
             
                End If
                
               '   CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:

                ClosePipe
                CardResult = DO_WritePort(card, Channel_P1A, &HA1)   ' Close power
                If (WaitHUBOFF("pid_6254") <> 1) Then
                    ResetHubReturn = Shell(App.Path & "\devcon rescan @USB\VID_058F&PID_6254", vbNormalFocus)
                    WaitProcQuit (ResetHubReturn)
                End If
                
                Call MsecDelay(0.3)
                
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv6 = 2 Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                    Tester.Label9.Caption = "ram unsatble Fail"
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
                 ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv5 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Public Sub AU6350GLF21TestSub()
'2013096 revise for usb host hang-out issue
     
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim EnumRes As Byte
             
              
              
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
                EnumRes = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
            '     CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                 EnumRes = WaitHUBOn("pid_6254")
                 If EnumRes <> 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &HA1)   ' Close power
                    WaitHUBOFF ("pid_6254")
                    Call MsecDelay(0.3)
                    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                    EnumRes = WaitHUBOn("pid_6254")
                 End If
                 
                 'Call MsecDelay(1.2)   'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                  
                  
                 If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect On Fail"
                    End
                 End If
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     
                    If EnumRes = 1 Then
                        EnumRes = WaitDevOn(ChipString)
                    End If
                    
                    Call MsecDelay(0.2)
                           
                      ClosePipe
                                      
                      If EnumRes = 1 Then
                        rv0 = CBWTest_New(0, 1, ChipString)
                      Else
                        rv0 = 0
                        Tester.Print "VID= unknow"; " ; ";
                        Tester.Print "PID= unknow"
                      End If
                      
                      If rv0 = 1 Then
                         rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                        If rv0 <> 1 Then
                           rv0 = 2
                           Tester.Print "SD bus width fail"
                        End If
                      End If
                      
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                                   
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        rv6 = 2
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                    End If
                         Call MsecDelay(0.01)
                     
                   
                     
                    
                       
                    
                     
                    
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
                
'                 CardResult = DO_WritePort(card, Channel_P1A, &H5E)
'
'                 If CardResult <> 0 Then
'                    MsgBox "Set MSPro Card Detect On Fail"
'                    End
'                 End If
'
'
'                 Call MsecDelay(0.1)
'
'
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                
                    Call MsecDelay(0.1)
                 
                   OpenPipe
                    rv5 = ReInitial(0)
                    ClosePipe
                    MsecDelay (0.02)
                End If
      
                ClosePipe
                rv5 = CBWTest_New(0, rv0, ChipString)
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               
           
                
               '   CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
                ClosePipe
                CardResult = DO_WritePort(card, Channel_P1A, &HA1)   ' Close power
                If (WaitHUBOFF("pid_6254") <> 1) Then
                    ResetHubReturn = Shell(App.Path & "\devcon rescan @USB\VID_058F&PID_6254", vbNormalFocus)
                    WaitProcQuit (ResetHubReturn)
                End If
                
                Call MsecDelay(0.3)
                
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
                           
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
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6350GLTestSub()
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                
             
              
              
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
            '     CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                 Call MsecDelay(2.2)   'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                  
                  
                 If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect On Fail"
                    End
                 End If
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     
                          Call MsecDelay(1)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                                   
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                    
                       
                    
                     
                    
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
                
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                 If rv0 = 1 Then
                
                 Call MsecDelay(0.03)
              
                 
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                
                  Call MsecDelay(0.2)
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect Down Fail"
                    End
                 End If
                 End If
                
                ClosePipe
                rv5 = CBWTest_New(0, rv0, ChipString)
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               
           
                
               '   CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
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
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

