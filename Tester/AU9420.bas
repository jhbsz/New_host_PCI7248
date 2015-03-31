Attribute VB_Name = "AU9420"
Public Sub AU9420TestSub()

If (ChipName = "AU9420BLF30") Or (ChipName = "AU9420DLF30") Then
    Call AU9420FT6TestSub
ElseIf (ChipName = "AU9420DLF00") Then
    Call AU9420FT3TestSub
End If

End Sub


Public Sub AU9420FT6TestSub()
      
'Using AU9420-JBL-NS-FT6-V1 Socket Board

Dim ChipString As String
OldChipName = ""
                
If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If
               
LBA = LBA + 1
                         
rv0 = 0 'Enum
rv1 = 0 'GPIO (O/S)
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
                

'==========================================================
'    Start Open Shot test
'==========================================================

OS_Result = 0
rv0 = 0
rv1 = 1

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.2)

OpenShortTest_Result

'CardResult = DO_WritePort(card, Channel_P1A, &H1E)

If OS_Result <> 1 Then
    rv1 = 0                 'OS Fail
    GoTo AU9420BLResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
Call MsecDelay(0.2)

Tester.Print "Check KeyBoard on Bus ..."


'=========================================
'   POWER on
'   SD Card test
'=========================================
ChipString = "pid_9420"

CardResult = DO_WritePort(card, Channel_P1A, &H7F)
Call MsecDelay(0.2)

If CardResult <> 0 Then
    MsgBox "Set Ena Detect Down Fail"
    End
End If

'Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)

CardResult = DO_WritePort(card, Channel_P1A, &HFF)
Call MsecDelay(0.2)
WaitDevOFF (ChipString)
                
AU9420BLResult:
    
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                      
    If rv1 = 0 Then
        'UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv0 = 0 Then
        Tester.Print "Eunm Fail"
        TestResult = "Bin3"
    ElseIf rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin5"
    End If
    
End Sub

Public Sub AU9420FT3TestSub()
      
'Using AU9420-JBL-NS-FT6-V1 Socket Board
'Only FT3 Enum

Dim ChipString As String
Dim HV_Done_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String


OldChipName = ""
                
If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If


Routine_Label:


If Not HV_Done_Flag Then
    Call PowerSet2(0, "3.6", "0.5", 1, "3.6", "0.5", 1)
    Call MsecDelay(0.3)
    Tester.Print "AU9420 : HV Begin Test ..."
    SetSiteStatus (RunHV)
Else
    Call PowerSet2(0, "3.0", "0.5", 1, "3.0", "0.5", 1)
    Call MsecDelay(0.3)
    Tester.Print vbCrLf & "AU9420 : LV Begin Test ..."
    SetSiteStatus (RunLV)
End If

               
LBA = LBA + 1
                         
rv0 = 0 'Enum
rv1 = 0 'GPIO (O/S)
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
                

'CardResult = DO_WritePort(card, Channel_P1C, &HFF)
Call MsecDelay(0.2)

Tester.Print "Check KeyBoard on Bus ..."


'=========================================
'   POWER on
'   SD Card test
'=========================================
ChipString = "pid_9420"

CardResult = DO_WritePort(card, Channel_P1A, &H7F)
Call MsecDelay(0.2)

If CardResult <> 0 Then
    MsgBox "Set Ena Detect Down Fail"
    End
End If

'Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)

CardResult = DO_WritePort(card, Channel_P1A, &HFF)
Call MsecDelay(0.2)
'WaitDevOFF (ChipString)
                
AU9420BLResult:

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
        ElseIf rv0 <> 1 Then
            HV_Result = "Fail"
            Tester.Print "HV Fail"
        ElseIf rv0 = 1 Then
            HV_Result = "PASS"
            Tester.Print "HV PASS"
        End If
        
        HV_Done_Flag = True
        GoTo Routine_Label
    Else
        If rv0 <> 1 Then
            LV_Result = "Bin2"
            Tester.Print "LV Unknow"
        ElseIf rv0 <> 1 Then
            LV_Result = "Fail"
            Tester.Print "LV Fail"
        ElseIf rv0 = 1 Then
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
    
    
End Sub
