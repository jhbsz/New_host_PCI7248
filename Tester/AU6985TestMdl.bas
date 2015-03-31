Attribute VB_Name = "AU6985TestMdl"
Public Function Read_DataAU6985HLF20(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Long
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

For i = 0 To 30
   
        CBW(i) = 0
    
Next i

For i = 0 To CBWDataTransferLength
ReadData(i) = 0
Next

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

CBWDataTransferLen(0) = (CBWDataTransferLength Mod 256)
tmpV(0) = Int(CBWDataTransferLength / 256)
CBWDataTransferLen(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
CBWDataTransferLen(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
CBWDataTransferLen(3) = (tmpV(2) Mod 256)

CBW(8) = CBWDataTransferLen(0)  '00
CBW(9) = CBWDataTransferLen(1)  '08
CBW(10) = CBWDataTransferLen(2) '00
CBW(11) = CBWDataTransferLen(3) '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = &H28
CBW(16) = Lun * 32
LBAByte(0) = (LBA Mod 256)
tmpV(0) = Int(LBA / 256)
LBAByte(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
LBAByte(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
LBAByte(3) = (tmpV(2) Mod 256)

CBW(17) = LBAByte(3)         '00
CBW(18) = LBAByte(2)         '00
CBW(19) = LBAByte(1)         '00
CBW(20) = LBAByte(0)         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

TransferLen = Int(CBWDataTransferLength / 512)

TransferLenLSB = (TransferLen Mod 256)
tmpV(0) = Int(TransferLen / 256)
TransferLenMSB = (tmpV(0) / 256)

CBW(22) = TransferLenMSB      '00
CBW(23) = TransferLenLSB      '04

For i = 24 To 30
    CBW(i) = 0
Next

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command

 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
 Read_DataAU6985HLF20 = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 
'If result = 0 Then
' Read_DataAU6985HLF20 = 0
' Exit Function
'End If

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
 Read_DataAU6985HLF20 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
    Read_DataAU6985HLF20 = 0
Else
     Read_DataAU6985HLF20 = 1
   
End If

 
End Function

Public Function Write_DataAU6985HLF20(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim CBWDataTransferLen(0 To 3) As Byte
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

opcode = &H2A
'Buffer(0) = &H33 'CByte(Text2.Text)
'Buffer(1) = &H44


    For i = 0 To 30
    
        CBW(i) = 0
    
    Next i
    
Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

CBWDataTransferLen(0) = (CBWDataTransferLength Mod 256)
tmpV(0) = Int(CBWDataTransferLength / 256)
CBWDataTransferLen(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
CBWDataTransferLen(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
CBWDataTransferLen(3) = (tmpV(2) Mod 256)

CBW(8) = CBWDataTransferLen(0)  '00
CBW(9) = CBWDataTransferLen(1)  '08
CBW(10) = CBWDataTransferLen(2) '00
CBW(11) = CBWDataTransferLen(3) '00

'///////////////  CBW Flag
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = opcode
CBW(16) = Lun * 32
LBAByte(0) = (LBA Mod 256)
tmpV(0) = Int(LBA / 256)
LBAByte(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
LBAByte(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
LBAByte(3) = (tmpV(2) Mod 256)

CBW(17) = LBAByte(3)         '00
CBW(18) = LBAByte(2)         '00
CBW(19) = LBAByte(1)         '00
CBW(20) = LBAByte(0)         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

TransferLen = Int(CBWDataTransferLength / 512)

TransferLenLSB = (TransferLen Mod 256)
tmpV(0) = Int(TransferLen / 256)
TransferLenMSB = (tmpV(0) / 256)

CBW(22) = TransferLenMSB      '00
CBW(23) = TransferLenLSB      '04

For i = 24 To 30
    CBW(i) = 0
Next

 
'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
    Write_DataAU6985HLF20 = 0
    Exit Function
End If
 
 
 
'2, Output data
result = WriteFile _
       (WriteHandle, _
       Pattern_64k(0), _
       CBWDataTransferLength, _
       NumberOfBytesWritten, _
       0)    'out

 
If result = 0 Then
    Write_DataAU6985HLF20 = 0
    Exit Function
End If

'3 . CSW
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
        
If result = 0 Then
    Write_DataAU6985HLF20 = 0
    Exit Function
End If
 
 
 
If CSW(12) = 1 Then
Write_DataAU6985HLF20 = 0

Else
Write_DataAU6985HLF20 = 1
End If
End Function

Public Function CBWTest_New_AU6985HLF21(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String, Flash As Byte) As Byte
'different read capacity routine: it will check capacity for MDG flash
Dim i As Long
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long
Dim HalfFlashCapacity As Long
Dim OldLBa As Long

   CBWDataTransferLength = 65536  ' 8 sector, MB flash 2k/page, and it is 4 flash
                                    '  8 sector for 2 page for 2 flash
                                    ' the another 8 sector for another 2 flash set
   
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6985HLF21 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6985HLF21 = 0
    If Flash = 1 Or Flash = 0 Then
     If LBA > 30000000 Then
         LBA = 0
     End If
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
     CBWTest_New_AU6985HLF21 = 0    ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6985HLF21 = 2    ' Write fail
      Exit Function
    End If
 
    '======================================
    
     ' for unitSpeed
    
    TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       CBWTest_New_AU6985HLF21 = 2    ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Exit Function
    End If
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU6985HLF21 = 2   'Write fail
           Exit Function
        End If
        
    End If
    '======================================
   CBWDataTransferLength = 512
    TmpInteger = Read_Data2(LBA, Lun, CBWDataTransferLength)
    '  TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
   ' If TmpInteger = 0 Then
   '      Read_CapacityAU6985HLF21 = 2   'write fail
   '       Exit Function
   '  End If
    
      
   ' for AU6390MB to read capacity
   
   
     TmpInteger = Read_CapacityAU6985HLF21(LBA, Lun, 8)
      
    If TmpInteger = 0 Then
         CBWTest_New_AU6985HLF21 = 3   'Read fail
          Exit Function
     ElseIf TmpInteger = 2 Then
           FlashCapacityError = 2
           CBWTest_New_AU6985HLF21 = 3   'card format has problem
          Exit Function
          
     End If
      
      
     If Flash = 0 Then
      CBWTest_New_AU6985HLF21 = 1
        Exit Function
     End If
      
 
   CBWDataTransferLength = 65536
    TmpInteger = Write_DataAU6985HLF20(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6985HLF21 = 2   'write fail
        Exit Function
    End If
    
    TmpInteger = Read_DataAU6985HLF20(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6985HLF21 = 3     'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
         
        
        If ReadData(i) <> Pattern_64k(i) Then
       
          CBWTest_New_AU6985HLF21 = 3     'Read fail
          Exit Function
        End If
    
    Next
  
    ' another 2 flash R/W
  
   
    
    
    CBWTest_New_AU6985HLF21 = 1
        
    
    End Function
    
Public Function CBWTest_New_AU6985ELF21(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String, Flash As Byte) As Byte
'different read capacity routine: it will check capacity for MDG flash
Dim i As Long
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long
Dim HalfFlashCapacity As Long
Dim OldLBa As Long

   CBWDataTransferLength = 65536  ' 8 sector, MB flash 2k/page, and it is 4 flash
                                    '  8 sector for 2 page for 2 flash
                                    ' the another 8 sector for another 2 flash set
   
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6985ELF21 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6985ELF21 = 0
    If Flash = 1 Or Flash = 0 Then
     If LBA > 30000000 Then
         LBA = 0
     End If
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
     CBWTest_New_AU6985ELF21 = 0    ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6985ELF21 = 2    ' Write fail
      Exit Function
    End If
 
    '======================================
    
     ' for unitSpeed
    
    TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       CBWTest_New_AU6985ELF21 = 2    ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Exit Function
    End If
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU6985ELF21 = 2   'Write fail
           Exit Function
        End If
        
    End If
    '======================================
   CBWDataTransferLength = 512
    TmpInteger = Read_Data2(LBA, Lun, CBWDataTransferLength)
    '  TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
   ' If TmpInteger = 0 Then
   '      Read_CapacityAU6985HLF21 = 2   'write fail
   '       Exit Function
   '  End If
    
      
   ' for AU6390MB to read capacity
   
   
     TmpInteger = Read_CapacityAU6985ELF21(LBA, Lun, 8)
      
    If TmpInteger = 0 Then
         CBWTest_New_AU6985ELF21 = 3   'Read fail
          Exit Function
     ElseIf TmpInteger = 2 Then
           FlashCapacityError = 2
           CBWTest_New_AU6985ELF21 = 3   'card format has problem
          Exit Function
          
     End If
      
      
     If Flash = 0 Then
      CBWTest_New_AU6985ELF21 = 1
        Exit Function
     End If
      
 
   CBWDataTransferLength = 65536
    TmpInteger = Write_DataAU6985HLF20(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6985ELF21 = 2   'write fail
        Exit Function
    End If
    
    TmpInteger = Read_DataAU6985HLF20(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6985ELF21 = 3     'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
         
        
        If ReadData(i) <> Pattern_64k(i) Then
       
          CBWTest_New_AU6985ELF21 = 3     'Read fail
          Exit Function
        End If
    
    Next
  
    ' another 2 flash R/W
  
   
    
    
    CBWTest_New_AU6985ELF21 = 1
        
    
    End Function
    
Public Function CBWTest_New_AU6985HLF20(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String, Flash As Byte) As Byte
Dim i As Long
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long
Dim HalfFlashCapacity As Long
Dim OldLBa As Long

   CBWDataTransferLength = 65536  ' 8 sector, MB flash 2k/page, and it is 4 flash
                                    '  8 sector for 2 page for 2 flash
                                    ' the another 8 sector for another 2 flash set
   
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6985HLF20 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6985HLF20 = 0
    If Flash = 1 Or Flash = 0 Then
     If LBA > 14000000 Then
         LBA = 0
     End If
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
      CBWTest_New_AU6985HLF20 = 0    ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6985HLF20 = 2    ' Write fail
      Exit Function
    End If
 
    '======================================
    
     ' for unitSpeed
    
    TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       CBWTest_New_AU6985HLF20 = 2    ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Exit Function
    End If
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU6985HLF20 = 2   'Write fail
           Exit Function
        End If
        
    End If
    '======================================
   CBWDataTransferLength = 512
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
      TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
   ' If TmpInteger = 0 Then
   '      CBWTest_New_AU6985HLF20 = 2   'write fail
   '       Exit Function
   '  End If
    
      
   ' for AU6390MB to read capacity
   
   
     TmpInteger = Read_CapacityAU6985HLF20(LBA, Lun, 8)
      
    If TmpInteger = 0 Then
         CBWTest_New_AU6985HLF20 = 3   'Read fail
          Exit Function
     ElseIf TmpInteger = 2 Then
           FlashCapacityError = 2
           CBWTest_New_AU6985HLF20 = 3   'card format has problem
          Exit Function
          
     End If
      
      
     If Flash = 0 Then
      CBWTest_New_AU6985HLF20 = 1
        Exit Function
     End If
      
 
   CBWDataTransferLength = 65536
    TmpInteger = Write_DataAU6985HLF20(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6985HLF20 = 2   'write fail
        Exit Function
    End If
    
    TmpInteger = Read_DataAU6985HLF20(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6985HLF20 = 3     'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
         
        
        If ReadData(i) <> Pattern_64k(i) Then
       
          CBWTest_New_AU6985HLF20 = 3     'Read fail
          Exit Function
        End If
    
    Next
  
    ' another 2 flash R/W
  
   
    
    
    CBWTest_New_AU6985HLF20 = 1
        
    
    End Function

Public Function Read_CapacityAU6985HLF20(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

Dim Capacity(0 To 7) As Byte

'Capacity(0) = &H0
'Capacity(1) = &H78
'Capacity(2) = &HFF
'Capacity(3) = &HFF
'Capacity(4) = &H0
'Capacity(5) = &H0
'Capacity(6) = &H2
'Capacity(7) = &H0

 Capacity(0) = &H1   ' MCG flash
  Capacity(1) = &HE6
 Capacity(2) = &H6F
 Capacity(3) = &HFF
 Capacity(4) = &H0
 Capacity(5) = &H0
 Capacity(6) = &H2
 Capacity(7) = &H0


 



For i = 0 To 30
   
        CBW(i) = 0
    
Next i

For i = 0 To CBWDataTransferLength
ReadData(i) = 0
Next

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

 
CBW(8) = &H8  '00
CBW(9) = &H0  '08
CBW(10) = &H0 '00
CBW(11) = &H0 '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = &H25
CBW(16) = Lun * 32
 
CBW(17) = &H0         '00
CBW(18) = &H0        '00
CBW(19) = &H0        '00
CBW(20) = &H0         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

 
CBW(22) = &H0     '00
CBW(23) = &H0     '04

For i = 24 To 30
    CBW(i) = 0
Next

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command

 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
 Read_CapacityAU6985HLF20 = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in




 
If result = 0 Then
  Read_CapacityAU6985HLF20 = 0
 Exit Function
End If


For i = 0 To CBWDataTransferLength - 1
Debug.Print "k", i, Hex(ReadData(i)), Capacity(i)
 'If ReadData(i) <> Capacity(i) Then
  
 ' Read_CapacityAU6985HLF20 = 2  ' card format capacity has problem
 '  Exit Function
 'End If


Next i


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
 Read_CapacityAU6985HLF20 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_CapacityAU6985HLF20 = 0
Else
      Read_CapacityAU6985HLF20 = 1
   
End If

 
End Function
Public Function Read_CapacityAU6985HLF21(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

Dim Capacity(0 To 7) As Byte

'Capacity(0) = &H0
'Capacity(1) = &H78
'Capacity(2) = &HFF
'Capacity(3) = &HFF
'Capacity(4) = &H0
'Capacity(5) = &H0
'Capacity(6) = &H2
'Capacity(7) = &H0

 'Capacity(0) = &H1   ' MCG flash
 'Capacity(1) = &HE6
'Capacity(2) = &H6F
'Capacity(3) = &HFF
'Capacity(4) = &H0
'Capacity(5) = &H0
'Capacity(6) = &H2
'Capacity(7) = &H0


Capacity(0) = &H3    'MDG flash
Capacity(1) = &HCC
Capacity(2) = &HDF
Capacity(3) = &HFF
Capacity(4) = &H0
Capacity(5) = &H0
Capacity(6) = &H2
Capacity(7) = &H0



For i = 0 To 30
   
        CBW(i) = 0
    
Next i

For i = 0 To CBWDataTransferLength
ReadData(i) = 0
Next

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

 
CBW(8) = &H8  '00
CBW(9) = &H0  '08
CBW(10) = &H0 '00
CBW(11) = &H0 '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = &H25
CBW(16) = Lun * 32
 
CBW(17) = &H0         '00
CBW(18) = &H0        '00
CBW(19) = &H0        '00
CBW(20) = &H0         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

 
CBW(22) = &H0     '00
CBW(23) = &H0     '04

For i = 24 To 30
    CBW(i) = 0
Next

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command

 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
 Read_CapacityAU6985HLF21 = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in




 
If result = 0 Then
  Read_CapacityAU6985HLF21 = 0
 Exit Function
End If


For i = 0 To CBWDataTransferLength - 1
Debug.Print "k", i, Hex(ReadData(i)), Capacity(i)
  If ReadData(i) <> Capacity(i) Then
  
  Read_CapacityAU6985HLF21 = 2  ' card format capacity has problem
   Exit Function
  End If


Next i


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
 Read_CapacityAU6985HLF21 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_CapacityAU6985HLF21 = 0
Else
      Read_CapacityAU6985HLF21 = 1
   
End If

 
End Function

Public Function Read_CapacityAU6985ELF21(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

Dim Capacity(0 To 7) As Byte

'Capacity(0) = &H0
'Capacity(1) = &H78
'Capacity(2) = &HFF
'Capacity(3) = &HFF
'Capacity(4) = &H0
'Capacity(5) = &H0
'Capacity(6) = &H2
'Capacity(7) = &H0

 Capacity(0) = &H1   ' MCG flash
 Capacity(1) = &HE6
 Capacity(2) = &H6F
 Capacity(3) = &HFF
 Capacity(4) = &H0
 Capacity(5) = &H0
 Capacity(6) = &H2
 Capacity(7) = &H0


 

For i = 0 To 30
   
        CBW(i) = 0
    
Next i

For i = 0 To CBWDataTransferLength
ReadData(i) = 0
Next

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

 
CBW(8) = &H8  '00
CBW(9) = &H0  '08
CBW(10) = &H0 '00
CBW(11) = &H0 '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = &H25
CBW(16) = Lun * 32
 
CBW(17) = &H0         '00
CBW(18) = &H0        '00
CBW(19) = &H0        '00
CBW(20) = &H0         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

 
CBW(22) = &H0     '00
CBW(23) = &H0     '04

For i = 24 To 30
    CBW(i) = 0
Next

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command

 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
 Read_CapacityAU6985ELF21 = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in




 
If result = 0 Then
  Read_CapacityAU6985ELF21 = 0
 Exit Function
End If


For i = 0 To CBWDataTransferLength - 1
Debug.Print "k", i, Hex(ReadData(i)), Capacity(i)
   If ReadData(i) <> Capacity(i) Then
  
  Read_CapacityAU6985ELF21 = 2  ' card format capacity has problem
   Exit Function
  End If


Next i


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
 Read_CapacityAU6985ELF21 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_CapacityAU6985ELF21 = 0
Else
      Read_CapacityAU6985ELF21 = 1
   
End If

 
End Function



Public Sub AU698XILF21TestSub()

  If PCI7248InitFinish = 0 Then
                  PCI7248Exist
  End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H0)  'Power Enable
                 Call MsecDelay(0.8)     'power on time
                 
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 Call MsecDelay(0.05)
                  LightOn = 0
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
              
                  If LightOn <> 254 Then
           
                  Call MsecDelay(0.2)
                  LightOn = 0
                 CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                 End If
             '    Debug.Print LightON
                   Tester.Print "get light on value"; LightOn
                  
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light ON fail"
                    End
                   End If
                 
               LBA = LBA + 1   ' for dual channel
               
              
               
                ClosePipe
                
                rv0 = CBWTest_New_AU6985HLF21(0, 1, "vid_058f", 0)
                 Call LabelMenu(0, rv0, 1)
                ClosePipe
                Tester.Print "1 st Flash R/W"
                rv1 = CBWTest_New_AU6985HLF21(0, rv0, "vid_058f", 1)
                Call LabelMenu(1, rv1, rv0)
                 ClosePipe
                  OldLBa = LBA
                  Tester.Print rv1; " 1 :pass,the other Fail "
                  LBA = LBA + 31879167 'org= 31879167 *2=3CCDFFFE *0.5=
                  LBA = LBA + 65535
                   Tester.Print "2 nd Flash R/W"
                 ClosePipe
                 rv2 = CBWTest_New_AU6985HLF21(0, rv1, "vid_058f", 2)
                 ClosePipe
                Call LabelMenu(2, rv2, rv1)
                   Tester.Print rv2; " 1 :pass,the other Fail "
                 
                
                 LBA = OldLBa
                
                ClosePipe
                
                
                   
            
                  
                       If rv2 = 1 Then
                          If LightOn = 254 Then
                             rv3 = 1
                           Else
                             GPOFail = 2
                             rv3 = 2
                          End If
                       Else
                            rv3 = 4
        
                       End If
               
                 
                 
                  Call LabelMenu(32, rv3, rv2)
                 
                     
                   CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  
                  
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
                ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
               

End Sub

Public Sub AU698XILF20TestSub()

  If PCI7248InitFinish = 0 Then
                  PCI7248Exist
  End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H0)  'Power Enable
                 Call MsecDelay(1.8)    'power on time
                 
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
               LightOn = 0
                 CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light ON fail"
                    End
                   End If
                 
               LBA = LBA + 1   ' for dual channel
               
              
               
                ClosePipe
                
                rv0 = CBWTest_New_AU6985HLF20(0, 1, "vid_058f", 0)
                 Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_AU6985HLF20(0, rv0, "vid_058f", 1)
                Call LabelMenu(1, rv1, rv0)
                 ClosePipe
                  OldLBa = LBA
                  
                  LBA = LBA + 15939584 'org= 31879167=1E66FFF*0.5
                  LBA = LBA + 65535
                 ClosePipe
                 rv2 = CBWTest_New_AU6985HLF20(0, rv1, "vid_058f", 2)
                 ClosePipe
                Call LabelMenu(2, rv2, rv1)
                
                 
                
                 LBA = OldLBa
                
                ClosePipe
                
                
                   
            
                  
                       If rv2 = 1 Then
                          If LightOn = 254 Then
                             rv3 = 1
                           Else
                             GPOFail = 2
                             rv3 = 2
                          End If
                       Else
                            rv3 = 4
        
                       End If
               
                 
                 
                  Call LabelMenu(32, rv3, rv2)
                 
                     
                   CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  
                  
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
                ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
               

End Sub

Public Sub AU698XHLF21TestSub()

  If PCI7248InitFinish = 0 Then
                  PCI7248Exist
  End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H0)  'Power Enable
                 Call MsecDelay(0.8)     'power on time
                 
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
                  Call MsecDelay(0.05)
                  LightOn = 0
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
              
                  If LightOn <> 223 Then
           
                  Call MsecDelay(0.2)
                  LightOn = 0
                 CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                 End If
             '    Debug.Print LightON
                   Tester.Print "get light on value"; LightOn
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light ON fail"
                    End
                   End If
                 
               LBA = LBA + 1   ' for dual channel
               
              
               
                ClosePipe
                
                rv0 = CBWTest_New_AU6985HLF21(0, 1, "vid_058f", 0)
                 Call LabelMenu(0, rv0, 1)
                ClosePipe
                Tester.Print "1 st Flash R/W"
                rv1 = CBWTest_New_AU6985HLF21(0, rv0, "vid_058f", 1)
                Call LabelMenu(1, rv1, rv0)
                 ClosePipe
                  OldLBa = LBA
                  Tester.Print rv1; " 1 :pass,the other Fail "
                  LBA = LBA + 31879167 'org= 31879167 *2=3CCDFFFE *0.5=
                  LBA = LBA + 65535
                   Tester.Print "2 nd Flash R/W"
                 ClosePipe
                 rv2 = CBWTest_New_AU6985HLF21(0, rv1, "vid_058f", 2)
                 ClosePipe
                Call LabelMenu(2, rv2, rv1)
                   Tester.Print rv2; " 1 :pass,the other Fail "
                 
                
                 LBA = OldLBa
                
                ClosePipe
                
                
                   
            
                  
                       If rv2 = 1 Then
                          If LightOn = 223 Then
                             rv3 = 1
                           Else
                             GPOFail = 2
                             rv3 = 2
                          End If
                       Else
                            rv3 = 4
        
                       End If
               
                 
                 
                  Call LabelMenu(32, rv3, rv2)
                 
                     
                   CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  
                  
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
                ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
               

End Sub
Public Sub AU698XELF21TestSub()

  If PCI7248InitFinish = 0 Then
                  PCI7248Exist
  End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H0)  'Power Enable
                 Call MsecDelay(0.8)     'power on time
                 
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
                  Call MsecDelay(0.05)
                  LightOn = 0
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
              
                  If LightOn <> 223 Then
           
                  Call MsecDelay(0.2)
                  LightOn = 0
                 CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                 End If
             '    Debug.Print LightON
                   Tester.Print "get light on value"; LightOn
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light ON fail"
                    End
                   End If
                 
               LBA = LBA + 1   ' for dual channel
               
              
               
                ClosePipe
                
                rv0 = CBWTest_New_AU6985ELF21(0, 1, "vid_058f", 0)
                 Call LabelMenu(0, rv0, 1)
                ClosePipe
                Tester.Print "1 st Flash R/W"
                rv1 = CBWTest_New_AU6985ELF21(0, rv0, "vid_058f", 1)
                Call LabelMenu(1, rv1, rv0)
                 ClosePipe
                  OldLBa = LBA
                  Tester.Print rv1; " 1 :pass,the other Fail "
               '   LBA = LBA + 31879167 'org= 31879167 *2=3CCDFFFE *0.5=
               '   LBA = LBA + 65535
              '     Tester.Print "2 nd Flash R/W"
              '   ClosePipe
              '   rv2 = CBWTest_New_AU6985ELF21(0, rv1, "vid_058f", 2)
              '   ClosePipe
              '  Call LabelMenu(2, rv2, rv1)
               '    Tester.Print rv2; " 1 :pass,the other Fail "
                 
                 rv2 = 1  ' no 2nd flash
                 LBA = OldLBa
                
                ClosePipe
                
                
                   
            
                  
                       If rv2 = 1 Then
                          If LightOn = 223 Then
                             rv3 = 1
                           Else
                             GPOFail = 2
                             rv3 = 2
                          End If
                       Else
                            rv3 = 4
        
                       End If
               
                 
                 
                  Call LabelMenu(32, rv3, rv2)
                 
                     
                   CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  
                  
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
                ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
               

End Sub
Public Sub AU698XHLF20TestSub()

  If PCI7248InitFinish = 0 Then
                  PCI7248Exist
  End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H0)  'Power Enable
                 Call MsecDelay(1.8)    'power on time
                 
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
               LightOn = 0
                 CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light ON fail"
                    End
                   End If
                 
               LBA = LBA + 1   ' for dual channel
               
              
               
                ClosePipe
                
                rv0 = CBWTest_New_AU6985HLF20(0, 1, "vid_058f", 0)
                 Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_AU6985HLF20(0, rv0, "vid_058f", 1)
                Call LabelMenu(1, rv1, rv0)
                 ClosePipe
                  OldLBa = LBA
                  
                  LBA = LBA + 15939584 'org= 31879167=1E66FFF*0.5=
                  LBA = LBA + 65535
                 ClosePipe
                 rv2 = CBWTest_New_AU6985HLF20(0, rv1, "vid_058f", 2)
                 ClosePipe
                Call LabelMenu(2, rv2, rv1)
                
                 
                
                 LBA = OldLBa
                
                ClosePipe
                
                
                   
            
                  
                       If rv2 = 1 Then
                          If LightOn = 223 Then
                             rv3 = 1
                           Else
                             GPOFail = 2
                             rv3 = 2
                          End If
                       Else
                            rv3 = 4
        
                       End If
               
                 
                 
                  Call LabelMenu(32, rv3, rv2)
                 
                     
                   CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                  
                  
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
                ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
               

End Sub


Public Sub AU698XTestSub()

If ChipName = "AU6985HLS10" Then
  
    Call AU6985HLS10TestSub
End If


If ChipName = "AU698XHLF20" Then
    Call AU698XHLF20TestSub
End If

If ChipName = "AU698XHLF21" Then
    Call AU698XHLF21TestSub
End If

If ChipName = "AU698XILF20" Then
    Call AU698XILF20TestSub
End If


If ChipName = "AU698XILF21" Then
    Call AU698XILF21TestSub
End If

If ChipName = "AU698XELF21" Then
    Call AU698XELF21TestSub
End If



End Sub



