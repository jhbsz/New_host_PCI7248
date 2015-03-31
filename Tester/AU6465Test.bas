Attribute VB_Name = "AU6465Test"
Public Sub AU6465TestSub()

    If (ChipName = "AU6465CFF20") Then
        Call AU6465CFF20TestSub
    End If

End Sub

Public Sub AU6465CFF20TestSub()

'2012/8/15 Use S/B:AU6465-GCF_GBF 28QFN SOCKET

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If

DetectCount = 0

Tester.Print Left(ChipName, 8) & " Begin Test ..."

Dim ChipString As String
OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'SD (Lun0)
rv2 = 0     'MS (Lun0)
rv3 = 0     'NBMD


Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)
                

' Ctrl_1

'8:XTAL_SEL (0: 12MHz[V], 1: 48MHz)
'7:
'6:
'5:

'4:
'3:MSINS
'2:SDCDN
'1:ENA


'=========================================
'    POWER on
'=========================================
CardResult = DO_WritePort(card, Channel_P1A, &HFC)      'SDCDN¡BENA On
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

ChipString = "058f"
             
'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)
Call MsecDelay(0.02)

LBA = LBA + 1
Tester.Print "LBA="; LBA

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
'  MSPro MS1 Card test
'===============================================
If rv1 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &HFA)  'MS¡BENA On
    rv2 = ReInitial(0)
    Call MsecDelay(0.1)
End If

If rv2 <> 1 Then
    Call NewLabelMenu(2, "Init", rv2, rv1)
Else
    rv2 = CBWTest_New_PipeReady(0, rv2, ChipString)
    
    If rv2 <> 1 Then
        Call NewLabelMenu(2, "MS", rv2, rv1)
    Else
        rv2 = CBWTest_New_128_Sector_PipeReady(0, rv2)   ' write
        Call NewLabelMenu(2, "MS_64K", rv2, rv1)
    
        If rv2 = 1 Then
            rv2 = Read_Speed2ReadData(LBA, 0, 64)
            If rv2 = 1 Then
                If (ReadData(22) = &H2) And (ReadData(23) = &H70) And (ReadData(24) = &H1) Then
                    Tester.Print "MS is 4 Bit, Speed 40 MHz"
                Else
                    Tester.Print "MS BusWidth/Speed Fail"
                    rv2 = 3
                End If
            End If
            Call NewLabelMenu(2, "MS Bus Width/Speed", rv2, rv1)
        End If
    End If

End If

Tester.Print rv2, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  NBMD test
'===============================================

If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)   ' Remove All Card
    Call MsecDelay(0.2)
    
    rv3 = WaitDevOFF(ChipString)
    
    If rv3 <> 1 Then
        Tester.Print "NB Mode Test Fail ..."
    End If
    Call MsecDelay(0.2)
    Call NewLabelMenu(3, "NBMD Test", rv3, rv2)
End If


AU6485CFFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)   ' Close power

                
    If rv0 <> 1 Then            'Enum
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf rv1 <> 1 Then        'SD
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then        'MS
        TestResult = "Bin5"
    ElseIf rv3 <> 1 Then        'NBMD
        TestResult = "Bin4"
    ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub
