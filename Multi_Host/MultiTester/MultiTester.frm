VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form MultiTester 
   Caption         =   "Tester"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   1215
      Left            =   1680
      TabIndex        =   2
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   855
      Left            =   5160
      TabIndex        =   1
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   855
      Left            =   5640
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   0
      Left            =   2040
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   1
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   2
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   3
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   4
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   5
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   6
      Left            =   240
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   7
      Left            =   1200
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "MultiTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TestName(0 To 300) As String
Dim ChipName As String
Dim AllenStop As Byte
Dim AllStop As Byte
Dim result


Sub SetTimer_1ms()
Dim err As Integer
err = CTR_Setup(card, 1, RATE_GENERATOR, 200, BINTimer)
err = CTR_Setup(card, 2, RATE_GENERATOR, 10, BINTimer)

End Sub
Private Sub Timer_1ms(ms As Integer)
Dim result
Dim old_value1
Dim old_value2
 
Dim i As Integer


 
result = CTR_Read(0, 2, old_value1)

   For i = 1 To ms
            Do
            DoEvents
            result = CTR_Read(0, 2, old_value2)
            Loop Until old_value1 <> old_value2
    
            Do
            DoEvents
            result = CTR_Read(0, 2, old_value2)
            Loop Until old_value1 = old_value2
    
    
    Next
     

 
End Sub
Private Sub Card_Initial()
  
  Dim result As Integer
  
  
    result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
    result = DO_WritePort(card, Channel_P1A, &HFF)
    
    result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
    result = DO_WritePort(card, Channel_P1B, &HFF)
    
     result = DIO_PortConfig(card, Channel_P1C, OUTPUT_PORT)
     result = DO_WritePort(card, Channel_P1C, &HFF)
     
     result = DIO_PortConfig(card, Channel_P2A, OUTPUT_PORT)
     result = DO_WritePort(card, Channel_P2A, &HFF)
     
     result = DIO_PortConfig(card, Channel_P2B, OUTPUT_PORT)
     result = DO_WritePort(card, Channel_P2B, &HFF)
     
     result = DIO_PortConfig(card, Channel_P2C, OUTPUT_PORT)
    result = DO_WritePort(card, Channel_P2C, &HFF)
    
     result = DIO_PortConfig(card, Channel_P3A, OUTPUT_PORT)
     result = DO_WritePort(card, Channel_P3A, &HFF)
     
     result = DIO_PortConfig(card, Channel_P3B, OUTPUT_PORT)
     result = DO_WritePort(card, Channel_P3B, &HFF)
    
     result = DIO_PortConfig(card, Channel_P3C, INPUT_PORT)
End Sub
    
Sub TestNameSub()
TestName(0) = "AU6371DF"
TestName(1) = "AU6332BS"
TestName(2) = "AU6331CS"
TestName(3) = "AU6254AL"
TestName(4) = "AU6254BL"
TestName(5) = "AU6376FL"
TestName(6) = "AU6371GL"
TestName(7) = "AU6376EL"
TestName(8) = "AU6376IL"
TestName(9) = "AU6337BS"
TestName(10) = "AU3130BL"
TestName(11) = "AU6391BL"
TestName(12) = "AU3130CL"
TestName(13) = "AU6337BL"
TestName(14) = "AU6254AF"
TestName(15) = "AU6376BL"
TestName(16) = "AU6337CS"
TestName(17) = "AU6375HL"
TestName(18) = "AU6981HL"
TestName(19) = "AU6254XL"
TestName(19) = "AU6371EL"
TestName(20) = "AU6377AL"
TestName(21) = "AU6371DL"
TestName(22) = "AU6982HL"
TestName(23) = "AU6254DL"
TestName(24) = "AU6370DL"

TestName(25) = "AU6334CL"
TestName(26) = "AU6371HL"
TestName(27) = "AU6376JL"
TestName(28) = "AU6336AF"
TestName(29) = "AU3150JL"
TestName(30) = "AU6366CL"
TestName(31) = "AU6337CF"
TestName(32) = "AU6254XL"

TestName(33) = "AU6370GL"
TestName(34) = "AU6336DF"
TestName(35) = "AU3150IL"
TestName(36) = "AU6337GL"
TestName(37) = "AU6332GF"
TestName(38) = "AU6332FF"
TestName(39) = "AU6371PL"
TestName(40) = "AU3150CL"
TestName(41) = "AU6371NL"
TestName(42) = "AU6986HL"
TestName(43) = "AU6986AL"
TestName(44) = "AU6378AL"
TestName(45) = "AU6371SL"
TestName(46) = "AU6371EL"
TestName(47) = "AU6430QL"
TestName(48) = "AU3150LL"
TestName(49) = "AU6395BL"
TestName(50) = "AU6395CL"
TestName(51) = "AU6420AL"
TestName(52) = "AU6371TL"
TestName(53) = "AU6420BL"
TestName(54) = "AU3152AL"
TestName(55) = "AU6336AS"
 
TestName(56) = "AU6336EF"
TestName(57) = "AU6376KL"
TestName(58) = "AU3150AL"
TestName(59) = "AU3150KL"
TestName(60) = "AU9520AL"
TestName(61) = "AU6710AS"
TestName(62) = "AU6336IF"
TestName(63) = "AU6337IL"
TestName(64) = "AU6430DL"
TestName(65) = "AU6430EL"
TestName(66) = "AU6430BL"
TestName(67) = "AU6256BL"
TestName(68) = "AU6336LF"
TestName(69) = "AU6471FL"
TestName(70) = "AU6471GL"
TestName(71) = "AU6420CL"
TestName(72) = "AU6350AL"
TestName(73) = "AU6378HL"
TestName(74) = "AU3150ML"
TestName(75) = "AU9368AL"
TestName(76) = "AU6336DL"
TestName(77) = "AU6980HL"
TestName(78) = "AU6476BL"
TestName(79) = "AU6433EF"
TestName(80) = "AU6336AA"
TestName(81) = "AU6378FL"
TestName(82) = "AU6254AS"
TestName(83) = "AU6256CF"
TestName(84) = "AU6433HF"
TestName(85) = "AU6350BF"
TestName(86) = "AU3152CL"
TestName(87) = "AU6433DF"
TestName(88) = "AU6433BS"
TestName(89) = "AU6476CL"
TestName(90) = "AU6432BS"
TestName(91) = "AU6476FL"
TestName(92) = "AU6476DL"
TestName(93) = "AU6476EL"
TestName(94) = "AU6376AL"
TestName(95) = "AU6433KF"
TestName(96) = "AU698XHL"
TestName(97) = "AU3150NL"
TestName(98) = "AU698XIL"
TestName(99) = "AU6476IL"
TestName(100) = "AU3150PL"
TestName(101) = "AU3150QL"
TestName(102) = "AU6433GS"
TestName(103) = "AU6336CA"
TestName(104) = "AU9525AL"
TestName(105) = "AU698XEL"
TestName(106) = "AU6366AL"
TestName(107) = "AU6433ES"
TestName(108) = "AU6433FS"
TestName(109) = "AU3152HL"
TestName(110) = "AU6336ZF"
TestName(111) = "AU6433HS"
TestName(112) = "AU6433IF"
TestName(113) = "AU6980OC"
TestName(114) = "AU6350GL"
TestName(115) = "AU6378RL"
TestName(116) = "AU6433LF"
TestName(117) = "AU6350BL"
TestName(118) = "AU6350CF"
TestName(119) = "AU6433BL"
TestName(120) = "AU9520FL"
TestName(121) = "AU9520GL"
TestName(122) = "AU6476JL"
TestName(123) = "AU6350OL"
TestName(124) = "AU6336HF"
TestName(125) = "AU6350KL"
TestName(126) = "AU6476LL"
TestName(127) = "AU1111AA"
TestName(128) = "AU6476ML"
TestName(129) = "AU6476QL"
End Sub
Function NameLen() As Integer
Dim i As Integer
Dim ChipInside As Integer
NameLen = 0
ChipInside = 0
Do

 If InStr(ChipName, TestName(i)) <> 0 Then
 ChipInside = 1
     If Len(ChipName) < 11 Then
            NameLen = 0
    Else
     
            NameLen = 1
            Exit Function
          
    End If

     If InStr(ChipName, "AU6375HL") <> 0 And Len(ChipName) = 8 Then
            NameLen = 1
             Exit Function
    
    Else
            NameLen = 0
           
    End If
    
 End If
i = i + 1
Loop While Len(TestName(i)) <> 0


   If InStr(ChipName, "AU6332BSF0") = 0 Then
            NameLen = 0
    Else
            NameLen = 1
            Exit Function
    End If
 
If ChipInside = 0 Then
 If Len(ChipName) < 6 Then
            NameLen = 0
    Else
            NameLen = 1
  End If
End If
 
End Function
Private Sub Command1_Click()

Dim i As Byte
Dim ChipName(0 To 7) As String
Dim buf As String

Dim k As Byte
Dim t1
Dim t2
Dim t3
Dim t4

        ' For i = 0 To 7
        ' MSComm1(i).InBufferCount = 0
        ' MSComm1(i).OutBufferCount = 0
        ' Next
  Do
         Call Timer_1ms(700)
        result = DO_WritePort(card, Channel_P1A, 0)
                  
                  Call Timer_1ms(7)
                result = DO_WritePort(card, Channel_P1A, &HFF)
                 Cls
                 
                  Call Timer_1ms(700)
                  
                'result = DO_WritePort(card, Channel_P1A, 0)
                  
                 ' Call Timer_1ms(7)
                'result = DO_WritePort(card, Channel_P1A, &HFF)
                 Cls
                 
                  Call Timer_1ms(700)
                  
                  t1 = Timer
                  Print "send Start"
  
         Do
               For k = 0 To 1
               For i = 0 To 7
               'If i <> 4 Then
                MSComm1(i).Output = "Ready"
                Call MsecDelay(0.1)
                
            '  End If
                DoEvents
               Next
               Next
               t2 = Timer
               Print Timer - t1
               For i = 0 To 7
               
                    buf = MSComm1(i).Input
                    ChipName(i) = ChipName(i) & buf
                     If InStr(1, ChipName(i), "AU6433BLF26") = 1 Then
                        Print "begin Test"; i, ChipName(i)
                        MSComm1(i).InBufferCount = 0
                        MSComm1(i).InputLen = 0
                          ChipName(i) = ""
                      End If
                        'ChipName(i) = ""
                    ' End If
                       ' MSComm1(i).InBufferCount = 0
                       ' MSComm1(i).OutBufferCount = 0
                       ' Call MsecDelay(1)
                          
                        '  MSComm1(i).Output = "PASS"
                        '  Call MsecDelay(1)
                          
                         '   MSComm1(i).InBufferCount = 0
                       ' MSComm1(i).OutBufferCount = 0
                        
                          
                    'End If
                   
               
                
               Next
          '   Call MsecDelay(1)
             AllenStop = 1
              t3 = Timer
               Print Timer - t2
             '     ChipName = "AU6395BLF20"
         Loop Until AllenStop = 1
         
     '   MsgBox "stop"
        Call MsecDelay(2)
        
         t4 = Timer
               'Print Timer - t3
         For i = 0 To 7
                  MSComm1(i).OutBufferCount = 0
                  Call MsecDelay(0.01)
                MSComm1(i).Output = "PASS"
                Call MsecDelay(0.01)
                
                DoEvents
               Next
         
      
   Loop Until AllStop = 1
   MsgBox "ALLTOP"
End Sub

Private Sub Command2_Click()
AllenStop = 1
End Sub

Private Sub Command3_Click()
AllStop = 1
MsgBox "STOP"
End Sub

Private Sub Form_Load()

Call TestNameSub
Dim i As Byte

For i = 0 To 7
    MSComm1(i).CommPort = i + 2
    
    MSComm1(i).Settings = "9600,N,8,1"
    MSComm1(i).PortOpen = True
    
    MSComm1(i).InBufferCount = 0
    MSComm1(i).InputLen = 0
Next i

  card = Register_Card(PCI_7296, 0) 'FOR PCI_7248
        Call SetTimer_1ms
        'SettingForm.Show 1 'FOR PCI_7248
        If card < 0 Then 'FOR PCI_7248
           MsgBox "Register Card Failed" 'FOR PCI_7248
        '   End 'FOR PCI_7248
        End If 'FOR PCI_7248
        Card_Initial 'FOR PCI_7248
End Sub
Public Sub MsecDelay(Msec As Single)
Dim start As Single
Dim pause As Single
start = Timer
    Do
         
        DoEvents
        pause = Timer
    Loop Until pause - start >= Msec
End Sub
