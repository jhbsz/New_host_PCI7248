Attribute VB_Name = "GeneralModule"
Public Declare Function fnScsi2usb_GetDiskHandle _
       Lib "usb2scsi.dll" (ByVal HDNumber As Long) _
As Long

Public Declare Function fnScsi2usb2K_KillEXE _
       Lib "usbreset.dll" _
       () _
As Integer

Public Declare Function WinExec _
       Lib "kernel32" _
       (ByVal lpCmdLine As String, _
       ByVal nCmdShow As Long) _
As Long

Public Type Dev
       DeviceCode  As Byte
       ZoneNumber As Byte
       BlockNumber As Integer
       SectorNumber As Integer
End Type
Type FlashMemory
      Handle As Long
      InstanceStr As String
      ChipNumber As Integer
      TotalZoneNumber As Integer
      PhysicalDiskCapacity As Double
      PhysicalDiskSectors As Double
      BadBlockNumber As Integer
      ErrStatus As Integer
End Type
Public hRemovealDisk As Long
Public DeviceChange As Integer
Public DeviceDetail As String
Public DCode(1 To 512) As Dev
Public Disk(1 To 16) As FlashMemory
Public TmpDisk(1 To 16) As FlashMemory
Public SendDataBuffer(1 To 512) As Byte
Public ReceiveDataBuffer(1 To 65536) As Byte
Public DataBufferArray(1 To 512) As Byte
Public AbortFlag As Boolean
Public StartCmdFlag As Boolean
Public ClearCmdFlag As Boolean
Public DiskCount As Integer
'Configuration String
Public ScsiManufacturerTxt As String
Public ScsiProductTxt As String
Public VidTxt As String
Public PidTxt As String
Public UsbManufacturerTxt As String
Public UsbProductTxt As String
Public PrimaryRation As Double
Public OldPrimaryRation As Double
Public SecondRation As Double
Public Re1m As Boolean
Public AppMode As Integer
Public tmpResult As String
Public Function AppPath()
    On Error Resume Next
    
    If Right(App.Path, 1) <> "\" Then
       AppPath = App.Path & "\"
    Else
       AppPath = App.Path
    End If
End Function
Public Sub PrintOut(result As Integer)
    Dim f As Integer
    Dim i As Long
    Dim FileName As String
    f = FreeFile
    On Error Resume Next
 
    'temporary file
    FileName = "C:\usb2scsi.txt"
    linestr = ""
     Open FileName For Input As #f
    If (result <> 1000) Then
        Line Input #1, tmpResult
    Else
        Line Input #1, DeviceDetail
    End If
    Close #f
   
    Kill FileName
End Sub
Public Sub Receive_DataBuffer()
    Dim FileName As String
    Dim f As Integer
    Dim i As Long
    On Error Resume Next
    
    i = 1
    f = FreeFile
    FileName = "c:\usb2scsi.bin"
    Open FileName For Binary As #f
    For i = 1 To 512
        ReceiveDataBuffer(i) = 0        'initialize ReceiveDataBuffer to 0
        Get #f, , ReceiveDataBuffer(i)
    Next i
    'Loop Until EOF(f)
    Close #f
    Kill FileName
End Sub

Public Function OpenScriptFile()
    Dim Fid As String, linestr As String
    Dim f As Integer, pos As Integer
    Dim i As Long, lineid As Integer
    Dim FileName As String
    f = FreeFile
    On Error GoTo ERR
 
    'temporary file
    FileName = AppPath & "Product.ini"
    linestr = ""
    
    VidTxt = "0ED1"
    PidTxt = "6680"
    
      Open FileName For Input As #f
            Do
                 Line Input #1, linestr
                 Fid = Mid(linestr, 1, 1)
                 If (Fid <> ";") Then
                    If Fid = "[" Then
                       Select Case linestr
                       Case "[VID/PID]"
                             lineid = 1
                       Case "[USB MANUFACTURE STRING]"
                             lineid = 2
                       Case "[USB PRODUCT STRING]"
                             lineid = 3
                       Case "[SCSI MANUFACTURE STRING]"
                             lineid = 4
                       Case "[SCSI PRODUCT STRING]"
                             lineid = 5
                       Case "[PARTITION STATUS]"
                             lineid = 6
                       Case "[1M Reserved]"
                             lineid = 7
                       Case "[APPLICATION MODE]"
                             lineid = 8
                       End Select
                    Else
                       If lineid = 1 Then
                          'If Mid(linestr, 1, 3) = "VID" Then
                          '   VidTxt = Mid(linestr, 5, 4)
                          '   Debug.Print VidTxt
                          'Else
                          '   PidTxt = Mid(linestr, 5, 4)
                          '   Debug.Print PidTxt
                          'End If
                       ElseIf lineid = 2 Then
                          UsbManufacturerTxt = linestr
                          Debug.Print UsbManufacturerTxt
                       ElseIf lineid = 3 Then
                          UsbProductTxt = linestr
                          Debug.Print UsbProductTxt
                       ElseIf lineid = 4 Then
                          ScsiManufacturerTxt = linestr
                          Debug.Print ScsiManufacturerTxt
                       ElseIf lineid = 5 Then
                          ScsiProductTxt = linestr
                          Debug.Print ScsiProductTxt
                       ElseIf lineid = 6 Then
                          If Mid(linestr, 1, 6) = "ParNum" Then
                                Partitionnum = CInt(Mid(linestr, 8, 1))
                                If Partitionnum > 2 Or Partitionnum < 1 Then
                                   MsgBox "Error Partition Number", vbCritical
                                   GoTo ERR
                                End If
                          ElseIf Mid(linestr, 1, 3) = "1st" Then
                                 pos = InStr(1, linestr, "=", vbTextCompare)
                                 PrimaryRation = CDbl(Right(linestr, Len(linestr) - pos))
                                 OldPrimaryRation = PrimaryRation
                             
                          ElseIf Mid(linestr, 1, 3) = "2nd" Then
                                 pos = InStr(1, linestr, "=", vbTextCompare)
                                 SecondRation = CDbl(Right(linestr, Len(linestr) - pos))
                             
                          End If
                       ElseIf lineid = 7 Then
                          If Mid(linestr, 1, 4) = "TRUE" Then
                             Re1m = True
                          Else
                             Re1m = False
                          End If
                       ElseIf lineid = 8 Then
                          If Mid(linestr, 1, 6) = "ApMode" Then
                             pos = InStr(1, linestr, "=", vbTextCompare)
                             AppMode = Val(Right(linestr, Len(linestr) - pos))
                          End If
                       End If
                  End If
                 End If
            Loop Until EOF(f)
      Close #f
      OpenScriptFile = 2
      If PrimaryRation + SecondRation <> 1 Then
         OpenScriptFile = 0
      End If
      Exit Function
ERR:
   Close #f
   OpenScriptFile = 1
   MsgBox "Open product.ini failed", vbCritical
 End Function
Public Sub delay(ByVal n As Single)
    On Error GoTo ERR_delay
    Dim tm1 As Single, tm2 As Single
    tm1 = Timer
    Do
      tm2 = Timer
      If tm2 < tm1 Then tm2 = tm2 + 86400
      If tm2 - tm1 > n Then Exit Do
      DoEvents
    Loop
    Exit Sub
ERR_delay:
      Resume Next
End Sub

