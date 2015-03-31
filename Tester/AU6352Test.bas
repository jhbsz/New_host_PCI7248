Attribute VB_Name = "AU6352Test"
Public Sub AU6352TestSub()
    
    If ChipName = "AU6352LLF20" Then
        Call AU6352LLF20TestSub
    ElseIf ChipName = "AU6352LLF00" Then
        Call AU6352LLF00TestSub
    ElseIf ChipName = "AU6352DFF20" Then
        Call AU6352DFF20TestSub
    End If
    
End Sub


Public Sub AU6352LLF20TestSub()
     
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
        'CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
        Call MsecDelay(1.2)   'power on time
                 
            
              
    '===============================================
    '  SD Card test
    '================================================
        
        CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                  
        If CardResult <> 0 Then
            MsgBox "Read light off fail"
            End
        End If
                  
                  
        If CardResult <> 0 Then
            MsgBox "Set SD Card Detect On Fail"
            End
        End If
                 
        Call MsecDelay(0.01)
                 
    '===========================================
    '   NO card test
    '============================================

        'set SD card detect down
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
                
        
    '===============================================
    '  CF Card test
    '================================================
                
        CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                
        Call MsecDelay(0.1)
        
        OpenPipe
        rv5 = ReInitial(0)
        ClosePipe
        
        Call MsecDelay(0.4)
        
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
           
                
               
                
AU6371DLResult:

           CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
            
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

Public Sub AU6352LLF00TestSub()

'2012/5/31 add for AU6352A41-JLL FT3

Dim ChipString As String
Dim i As Integer
Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single
Dim LV_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String
          
If PCI7248InitFinish = 0 Then
    PCI7248Exist
End If

LV_Flag = False
HV_Result = ""
LV_Result = ""
Tester.Cls
    

                     
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

Routine_Label:

'CardResult = DO_WritePort(card, Channel_P1A, &H7F)
'Call MsecDelay(0.2)   'power on time



CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
          
If CardResult <> 0 Then
    MsgBox "Read light off fail"
    End
End If

          
Call MsecDelay(0.01)
       
LBA = LBA + 1
       
'=========================================
'    POWER on  test anotehr card
'=========================================
            
CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                  
If Not LV_Flag Then
    Call PowerSet2(0, "3.6", "0.5", 1, "3.6", "0.5", 1)
    Tester.Print "AU6352JL : HV Begin Test ..."
Else
    Call PowerSet2(0, "3.0", "0.5", 1, "3.0", "0.5", 1)
    Tester.Print "AU6352JL : LV Begin Test ..."
    'Call MsecDelay(1#)
End If

Call MsecDelay(1.5)   'power on time
rv0 = WaitDevOn("058f")
Call MsecDelay(0.1)

'===========================================
'   SD card test
'============================================

'set SD card detect down
'CardResult = DO_WritePort(card, Channel_P1A, &H7E)
'
'If CardResult <> 0 Then
'    MsgBox "Set SD Card Detect Down Fail"
'    End
'End If
'
'Call MsecDelay(0.1)
                   
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
            'rv6 = 2
            GoTo AU9560End_Label
        End If
    Next
Else
    Call LabelMenu(0, rv0, 1)   ' no card test fail
    GoTo AU9560End_Label
End If
    
Call MsecDelay(0.01)
Call LabelMenu(0, rv0, 1)   ' no card test fail
Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
'===============================================
'  MS Card test
'================================================
                
CardResult = DO_WritePort(card, Channel_P1A, &H5F)
        
Call MsecDelay(0.1)

OpenPipe
rv5 = ReInitial(0)
ClosePipe

Call MsecDelay(0.2)

ClosePipe
rv5 = CBWTest_New(0, rv0, ChipString)
Call LabelMenu(31, rv5, rv0)

Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
ClosePipe
               

'**************************************************************************************************
'2010/7/14 skip AU650JKL Port1(AU6336 reader) test
'
'                rv7 = 1
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

AU9560End_Label:

CardResult = DO_WritePort(card, Channel_P1A, &HFF)
Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
WaitDevOFF ("vid_058f")
Call MsecDelay(0.2)

If Not LV_Flag Then
    If rv0 = 0 Then
        LV_Result = "Bin2"
    ElseIf rv0 * rv5 <> 1 Then
        LV_Result = "Fail"
    ElseIf rv0 * rv5 = 1 Then
        LV_Result = "PASS"
    Else
        LV_Result = "Fail"
    End If
    
    rv0 = 0
    rv5 = 0
    LV_Flag = True
    GoTo Routine_Label
    
Else
    If rv0 = 0 Then
        HV_Result = "Bin2"
    ElseIf rv0 * rv5 <> 1 Then
        HV_Result = "Fail"
    ElseIf rv0 * rv5 = 1 Then
        HV_Result = "PASS"
    Else
        HV_Result = "Fail"
    End If
End If



If (HV_Result = "Bin2") And (LV_Result = "Bin2") Then
    TestResult = "Bin2"
    AU6254CMediaFailCounter = AU6254CMediaFailCounter + 1
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
    AU6254CMediaFailCounter = AU6254CMediaFailCounter + 1
End If

If AU6254CMediaFailCounter >= 5 Then
    Shell "cmd /c shutdown -r  -t 0", vbHide
End If

            
End Sub

Public Sub AU6352DFF20TestSub()

'2011/4/24 This code copy from AU6350CFF22

Dim ChipString As String
Dim HubReader As String
Dim Port1Reader As String
             
'ChipString = "6366"
Port1Reader = "6335"
HubReader = "6366"
               
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

Call MsecDelay(0.2)

CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)

If CardResult <> 0 Then
    MsgBox "Read light off fail"
    End
End If
                 
CardResult = DO_WritePort(card, Channel_P1A, &HFC)  'Ena & Port1 power on
                  
Call MsecDelay(0.2)  'power on time
rv0 = WaitDevOn(Port1Reader)
Call MsecDelay(0.1)
              
'===============================================
'  SD Card test
'================================================
                
ClosePipe
HubPort = 0
ReaderExist = 0

rv0 = CBWTest_New_no_card_AU6352DF(0, rv0, Port1Reader)
ClosePipe
'Call MsecDelay(0.02)
        
ReaderExist = 0
rv0 = CBWTest_New_AU6350CF(0, rv0, Port1Reader)
                 
If rv0 = 1 Then
    rv0 = Read_SD_Speed_AU6371(0, 0, 18, "4Bits")
    
    If rv0 <> 1 Then
        rv0 = 2
        Tester.Print "Hub Reader SD bus width fail"
    End If
End If
ClosePipe

Call LabelMenu(1, rv0, 1)   ' no card test fail
Tester.Print rv0, " \\Port1SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===========================================
'HubReader card test
'============================================
  
' set SD card detect down

CardResult = DO_WritePort(card, Channel_P1A, &HF4)
  
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If
                           
ClosePipe
HubPort = 1
ReaderExist = 0
                      
rv1 = CBWTest_New_AU6350CF(0, rv0, HubReader)          'Hub reader
                      
If rv1 = 1 Then
    rv1 = Read_SD_Speed_AU6371(0, 0, 18, "4Bits")
    
    If rv1 <> 1 Then
        rv1 = 2
        Tester.Print "Hub Reader SD bus width fail"
    End If
End If
                                            
ClosePipe
                      
                      
If rv1 = 1 Then
                      
    For i = 1 To 20
                      
        If rv1 = 1 Then
            ClosePipe
            rv1 = CBWTest_New_AU6375IncPattern(0, 1, HubReader)
            ClosePipe
        End If
                   
        If rv1 = 1 Then
            ClosePipe
            rv1 = CBWTest_New_AU6375IncPattern(0, 1, HubReader)
            ClosePipe
        End If
                 
        If rv1 = 1 Then
            ClosePipe
            rv1 = CBWTest_New_AU6375IncPattern2(0, 1, HubReader)
            ClosePipe
        End If
              
        If rv1 = 1 Then
            ClosePipe
            rv1 = CBWTest_New_AU6375IncPattern2(0, 1, HubReader)
            ClosePipe
        End If
              
        If rv1 <> 1 Then
            rv1 = 2
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
                 
Tester.Print rv1, " \\HubSD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
'Tester.Print rv2, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                

CardResult = DO_WritePort(card, Channel_P1A, &HFC)  'Ena & Port1 power on
Call MsecDelay(0.2)
HubPort = 0
ReaderExist = 0

If GetDeviceNameMulti(HubReader) = "" Then
    rv2 = 1
Else
    rv2 = 0
End If
                
Call LabelMenu(1, rv2, rv1)   ' no card test fail
Tester.Print rv2, " \\NBMode :0 NBMode Fail, 1 pass "
                
                
AU6371DLResult:
                
            CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
            Call MsecDelay(0.2)
            
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
