Attribute VB_Name = "AU6990Sorting"
Public Sub AU6990HW_SortTest()
  
    Tester.Print "AU6990 HW Sorting"

    Dim ChipString As String
    Dim VersionCode As Byte
    Dim AU6371EL_SD As Byte
    Dim AU6371EL_CF As Byte
    Dim AU6371EL_XD As Byte
    Dim AU6371EL_MS As Byte
    Dim AU6371EL_MSP  As Byte
    Dim AU6371EL_BootTime As Single
                
    OldChipName = ""
    Tester.Label9.Caption = ""
    
    ' initial condition
                
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
                    
    AU6371EL_BootTime = 0.6
                     
                    
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
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                End
            End If
                
            Call MsecDelay(0.5)
                 
            ChipString = "vid"
                
            If GetDeviceName(ChipString) <> "" Then
                rv0 = 0
                GoTo AU6990HW_Result
            End If
                 
        '================================================
                
            CardResult = DO_WritePort(card, Channel_P1A, &HFB)
            
            Call MsecDelay(6#)
                
            rv0 = AU6990HW_Version
            If rv0 = 1 Then
                rv1 = Read_6990HW_Code(1, 0, 512)
            End If
            
            ClosePipe
                      
            Tester.Print "rv0= "; rv0
               
            Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
            
                     
            If rv1 = 1 Then     'AU6990A5X or AU6990R5X
                rv0 = 1     'bin4
                rv1 = 1
                rv2 = 2
                rv3 = 2
                rv4 = 1
                rv5 = 1
                Tester.Print "Chip-Type is AU6990-A5X or AU6990-R5X"
                Tester.Label9.Caption = "AU6990 -A5X or AU6990 -R5X"
            ElseIf rv1 = 2 Then 'AU6990B5X
                rv0 = 1     'bin1
                rv1 = 1
                rv2 = 1
                rv3 = 1
                rv4 = 1
                rv5 = 1
                Tester.Print "Chip-Type is AU6990-B5X"
                Tester.Label9.Caption = "AU6990 -B5X"
            ElseIf rv1 = 3 Then 'AU6990S5X
                rv0 = 1     'bin5
                rv1 = 1
                rv2 = 1
                rv3 = 1
                rv4 = 2
                rv5 = 2
                Tester.Print "Chip-Type is AU6990-S5X"
                Tester.Label9.Caption = "AU6990 -S5X"
            End If
                
AU6990HW_Result:
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
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


Public Sub AU6922HW_SortTest()
  
    Tester.Print "AU6922 HW Sorting"

    Dim ChipString As String
    Dim VersionCode As Byte
    Dim AU6371EL_SD As Byte
    Dim AU6371EL_CF As Byte
    Dim AU6371EL_XD As Byte
    Dim AU6371EL_MS As Byte
    Dim AU6371EL_MSP  As Byte
    Dim AU6371EL_BootTime As Single
                
    OldChipName = ""
    Tester.Label9.Caption = ""
    
    ' initial condition
                
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
                    
    AU6371EL_BootTime = 0.6
                     
                    
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
            Call PowerSet2(1, "5.0", "0.15", 1, "3.3", "0.15", 1) ' close power to disable chip
            CardResult = DO_WritePort(card, Channel_P1A, &HFB)
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                End
            End If
                
            'Call MsecDelay(0.5)
                 
            ChipString = "vid"
                
'            If GetDeviceName(ChipString) <> "" Then
'                rv0 = 0
'                GoTo AU6922HW_Result
'            End If

            If (WaitDevOn("058f") <> 1) Then
                GoTo AU6922HW_Result
            End If
            
            'Call MsecDelay(0.5)
                 
        '================================================
            
            'Call MsecDelay(6#)
                
            rv0 = AU6922HW_Version
            If rv0 = 1 Then
                rv0 = Read_6922HW_Code(1, 0, 512)
            End If
            
            ClosePipe
                      
            Tester.Print "rv0= "; rv0
               
            Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
            
                     
'            If rv1 = 1 Then     'AU6990A5X or AU6990R5X
'                rv0 = 1     'bin4
'                rv1 = 1
'                rv2 = 2
'                rv3 = 2
'                rv4 = 1
'                rv5 = 1
'                Tester.Print "Chip-Type is AU6990-A5X or AU6990-R5X"
'                Tester.Label9.Caption = "AU6990 -A5X or AU6990 -R5X"
'            ElseIf rv1 = 2 Then 'AU6990B5X
'                rv0 = 1     'bin1
'                rv1 = 1
'                rv2 = 1
'                rv3 = 1
'                rv4 = 1
'                rv5 = 1
'                Tester.Print "Chip-Type is AU6990-B5X"
'                Tester.Label9.Caption = "AU6990 -B5X"
'            ElseIf rv1 = 3 Then 'AU6990S5X
'                rv0 = 1     'bin5
'                rv1 = 1
'                rv2 = 1
'                rv3 = 1
'                rv4 = 2
'                rv5 = 2
'                Tester.Print "Chip-Type is AU6990-S5X"
'                Tester.Label9.Caption = "AU6990 -S5X"
'            End If
                
AU6922HW_Result:
                
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
        If rv0 = 0 Then
            TestResult = "Bin2"
        ElseIf rv0 = 1 Then
            TestResult = "PASS"
        ElseIf rv0 = 2 Then
            TestResult = "Bin3"
        Else
            TestResult = "Bin2"
        End If
End Sub



