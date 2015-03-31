VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form MPTester 
   AutoRedraw      =   -1  'True
   Caption         =   "ALCOR TESTER"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   1815
   ClientWidth     =   3165
   Icon            =   "AlcorTester.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   3165
   Begin VB.CheckBox LLF 
      Caption         =   "87100 LLF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   25
      Top             =   3360
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   1560
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5040
      TabIndex        =   24
      Top             =   3120
      Width           =   735
   End
   Begin VB.CheckBox skip_other_usb 
      Caption         =   "skip_other_usb"
      Height          =   495
      Left            =   3240
      TabIndex        =   23
      Top             =   3000
      Width           =   1695
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1560
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton UpdateFW_Btn 
      Caption         =   "Update FW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton HLV_Test 
      Caption         =   "HLV Test"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   20
      Top             =   360
      Width           =   735
   End
   Begin VB.CheckBox FailCloseFT 
      Caption         =   "Fail Close FT Tool"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   19
      Top             =   4560
      Value           =   1  '核取
      Width           =   1935
   End
   Begin VB.ComboBox SkipCtrlCount 
      Height          =   315
      ItemData        =   "AlcorTester.frx":0CCA
      Left            =   2400
      List            =   "AlcorTester.frx":0CEC
      Style           =   2  '單純下拉式
      TabIndex        =   18
      Top             =   4200
      Width           =   630
   End
   Begin VB.CheckBox AutoMP_Option 
      Caption         =   "AutoMP"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   5760
      Value           =   1  '核取
      Width           =   975
   End
   Begin VB.CheckBox ResetMPFailCounter 
      Caption         =   "Reset  MP Fail Counter"
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CheckBox NoMP 
      Caption         =   "NO MP"
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CheckBox Bin2 
      Caption         =   "Bin2 >5, shutdown"
      Height          =   555
      Left            =   1560
      TabIndex        =   10
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "STOP"
      Height          =   495
      Left            =   -120
      TabIndex        =   7
      Top             =   8640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton StartRWTest 
      Caption         =   "Start RW Test"
      Height          =   495
      Left            =   -120
      TabIndex        =   5
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton LoadRWTest 
      Caption         =   "Load RW Test/Satrt MP"
      Height          =   495
      Left            =   -120
      TabIndex        =   4
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox MPText 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   8880
      Width           =   975
   End
   Begin VB.CommandButton StartMP 
      Caption         =   "Start MP /Scan"
      Height          =   495
      Left            =   -120
      TabIndex        =   1
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton LoadMP 
      Caption         =   "Auto MP"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Height          =   330
      Left            =   1560
      TabIndex        =   9
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label FWFail_Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0FFFF&
      Caption         =   "Receive FW Fail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   2040
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Skip GPIB Counter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label GPIBCARD_Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H0080FFFF&
      Caption         =   "GPIB Exist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Line Line3 
      X1              =   1320
      X2              =   3240
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   1320
      Y1              =   5640
      Y2              =   9240
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3120
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.Label TestResultLab 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   2895
   End
End
Attribute VB_Name = "MPTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ContiUnknowFailCounter As Integer
Public ResetHubString As String
Public ResetReturn As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, _
  ByVal lpWindowName As String) As Long
  
Private Declare Function ShowWindow Lib "user32" _
  (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
  
Private Declare Function SetForegroundWindow Lib "user32" _
  (ByVal hwnd As Long) As Long
  
Public AU7670 As Boolean

Private Function CheckMe(fm As Form) As Boolean
Dim hwnd As Long
Dim PrevAppCaption As String
    PrevAppCaption = fm.Caption
    If App.PrevInstance Then
        fm.Caption = ""
        hwnd = FindWindow(vbNullString, PrevAppCaption)
        ShowWindow hwnd, 9
        SetForegroundWindow hwnd
        CheckMe = True
    Else
        CheckMe = False
    End If
End Function

Public Sub ClearCurDeviceParameter()

    With CurDevicePar
        .Dual_Flag = False
        .EQC_Flag = False
        .FullName = ""
        .ShortName = ""
        .SetStdV1 = ""
        .SetStdV2 = ""
        .SetHLVStd1 = ""
        .SetHLVStd2 = ""
        .SetHV1 = ""
        .SetHV2 = ""
        .SetLV1 = ""
        .SetLV2 = ""
        .SetStdI1 = ""
        .SetStdI2 = ""
        .SetHI1 = ""
        .SetHI2 = ""
        .SetLI1 = ""
        .SetLI2 = ""
        .MP_ToolFileName = ""
        .MP_LoadTitle = ""
        .MP_WorkTitle = ""
        .DeviceFolder = ""
'        .UpdateModuleName = ""
        .FT_ToolFileName = ""
        .FT_ToolTitle = ""
        .Exec_Par1 = ""
        .Exec_Par2 = ""
    End With
    
End Sub

Public Function WaitProcQuit(pid As Long)

On Error Resume Next
Dim objProcess
Dim Pid_Exist As Boolean

    Do
        Pid_Exist = True
        For Each objProcess In GetObject("winmgmts:\\.\root\cimv2:win32_process").instances_
            ''debug.print objProcess.Handle; objProcess.Name
             If objProcess.Handle = pid Then
                Pid_Exist = False
                ''debug.print "Ongo"
                Exit For
             End If
        Next
        
    Loop Until (Pid_Exist)

    Set objProcess = Nothing
    
End Function

Public Sub ReadCurDeviceInfo()

Dim sDBPAth As String
Dim sConStr As String
Dim oCn As New ADODB.Connection
Dim RS_Files As New ADODB.Recordset
Dim RS_Par As New ADODB.Recordset

    '-----------------------------
    ' set Path and connection string
    '---------------------------
    sDBPAth = App.Path & "\AlcorMPTesterDB\FT_Program.mdb"

    If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
        MsgBox "MDB no EXIST"
        Exit Sub
    End If

    Call ClearCurDeviceParameter

    sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDBPAth
    oCn.Open sConStr
    RS_Par.CursorLocation = adUseClient
    RS_Par.Open "AU69XX_Parameter", oCn, adOpenKeyset, adLockReadOnly
    Set RS_Par = oCn.Execute("Select * From [AU69XX_Parameter] Where [InProcess] = 1 Order By [FullName]")
    RS_Par.MoveFirst
    
    Do Until RS_Par.EOF
        With CurDevicePar
            If (Trim(RS_Par.Fields("FullName")) = ChipName) Then
                .FullName = Trim(RS_Par.Fields("Fullname"))
                .EQC_Flag = RS_Par.Fields("EQC_Flag")
                
                'Dim tmp As String
                'tmp = .FullName
                '.FullName = Replace(tmp, vbCrLf, "")
                .Dual_Flag = RS_Par.Fields("Dual_Flag")
                OldVer_Flag = RS_Par.Fields("OldAP_Flag")
                
                If IsNumeric(RS_Par.Fields("SetStdV1")) Then
                    .SetStdV1 = Trim(RS_Par.Fields("SetStdV1"))
                End If
                
                If IsNumeric(RS_Par.Fields("SetStdV2")) Then
                    .SetStdV2 = Trim(RS_Par.Fields("SetStdV2"))
                End If
                
                If IsNumeric(RS_Par.Fields("SetHLVStd1")) Then
                    .SetHLVStd1 = Trim(RS_Par.Fields("SetHLVStd1"))
                End If
                
                If IsNumeric(RS_Par.Fields("SetHLVStd2")) Then
                    .SetHLVStd2 = Trim(RS_Par.Fields("SetHLVStd2"))
                End If
                
                If IsNumeric(RS_Par.Fields("SetHV1")) Then
                    .SetHV1 = Trim(RS_Par.Fields("SetHV1"))
                End If
                
                If IsNumeric(RS_Par.Fields("SetHV2")) Then
                    .SetHV2 = Trim(RS_Par.Fields("SetHV2"))
                End If
                
                If IsNumeric(RS_Par.Fields("SetLV1")) Then
                    .SetLV1 = Trim(RS_Par.Fields("SetLV1"))
                End If
                
                If IsNumeric(RS_Par.Fields("SetLV2")) Then
                    .SetLV2 = Trim(RS_Par.Fields("SetLV2"))
                End If
                
                If IsNumeric(RS_Par.Fields("SetStdI1")) Then
                    .SetStdI1 = Trim(RS_Par.Fields("SetStdI1"))
                End If
                
                If IsNumeric(RS_Par.Fields("SetStdI2")) Then
                    .SetStdI2 = Trim(RS_Par.Fields("SetStdI2"))
                End If
                
                If IsNumeric(RS_Par.Fields("SetHI1")) Then
                    .SetHI1 = Trim(RS_Par.Fields("SetHI1"))
                End If
                
                If IsNumeric(RS_Par.Fields("SetHI2")) Then
                    .SetHI2 = Trim(RS_Par.Fields("SetHI2"))
                End If
                
                If IsNumeric(RS_Par.Fields("SetLI1")) Then
                    .SetLI1 = Trim(RS_Par.Fields("SetLI1"))
                End If
                
                If IsNumeric(RS_Par.Fields("SetLI2")) Then
                    .SetLI2 = Trim(RS_Par.Fields("SetLI2"))
                End If
                
                Exit Do
            End If
            RS_Par.MoveNext
        End With
    Loop

    Set RS_Files = oCn.Execute("Select * From [AU69XX_Files] Order By [ShortName]")
    RS_Files.MoveFirst
    Do Until RS_Files.EOF
        With CurDevicePar
            If (Trim(RS_Files.Fields("ShortName")) = Left(ChipName, 6)) Then
                .ShortName = Trim(RS_Files.Fields("ShortName"))
                .DeviceFolder = Trim(RS_Files.Fields("DeviceFolder"))
                .MP_LoadTitle = Trim(RS_Files.Fields("MP_LoadTitle"))
                .MP_WorkTitle = Trim(RS_Files.Fields("MP_WorkTitle"))
                .MP_ToolFileName = Trim(RS_Files.Fields("MP_ToolFileName"))
'                .UpdateModuleName = Trim(RS_Files.Fields("UpdateModuleName"))
                .FT_ToolFileName = Trim(RS_Files.Fields("FT_ToolFileName"))
                If (OldVer_Flag) Then
                    .FT_ToolFileName = "Old_" & .FT_ToolFileName
                    '.MP_ToolFileName = "Old_" & .MP_ToolFileName
                    If InStr(ChipName, "6991") Then
                        .MP_WorkTitle = "698x UFD MP, Cycle Time : 33 ns"
                    End If
                    Print .FT_ToolFileName
                End If
                
'                If Left(ChipName, 6) = "AU6919" Or Left(ChipName, 6) = "AU6921" Or Left(ChipName, 6) = "AU6922" Or Left(ChipName, 6) = "AU6925" Or Left(ChipName, 6) = "AU6998" Then
'
'                    If (OldVer_Flag = False) And Left(ChipName, 6) = "AU6921" Then
'                        FileCopy App.Path & CurDevicePar.DeviceFolder & "\98SNnew\10_D1_K9F.BIN", App.Path & CurDevicePar.DeviceFolder & "\98SN\10_D1_K9F.BIN"
'                        FileCopy App.Path & CurDevicePar.DeviceFolder & "\98SNnew\FlashList.csv", App.Path & CurDevicePar.DeviceFolder & "\98SN\FlashList.csv"
'
'                    ElseIf (OldVer_Flag = False) And Left(ChipName, 6) = "AU6925" Then
'                        FileCopy App.Path & CurDevicePar.DeviceFolder & "\98TNnew\10_D1_K9F.BIN", App.Path & CurDevicePar.DeviceFolder & "\98TN\10_D1_K9F.BIN"
'                        FileCopy App.Path & CurDevicePar.DeviceFolder & "\98TNnew\FlashList.csv", App.Path & CurDevicePar.DeviceFolder & "\98TN\FlashList.csv"
'
'                    ElseIf (OldVer_Flag = False) And Left(ChipName, 6) = "AU6998" Then
'                        FileCopy App.Path & CurDevicePar.DeviceFolder & "\98new\98_D1_K9F.BIN", App.Path & CurDevicePar.DeviceFolder & "\98\98_D1_K9F.BIN"
'                        FileCopy App.Path & CurDevicePar.DeviceFolder & "\98new\FlashList.csv", App.Path & CurDevicePar.DeviceFolder & "\98\FlashList.csv"
'
'                    ElseIf (OldVer_Flag = False) Then
'                        FileCopy App.Path & CurDevicePar.DeviceFolder & "\98ANnew\10_D1_K9F.BIN", App.Path & CurDevicePar.DeviceFolder & "\98AN\10_D1_K9F.BIN"
'                        FileCopy App.Path & CurDevicePar.DeviceFolder & "\98ANnew\FlashList.csv", App.Path & CurDevicePar.DeviceFolder & "\98AN\FlashList.csv"
'
'                    End If
'
'                    If Left(ChipName, 6) = "AU6921" Then
'                        ' add for copy MSL file
'                        If InStr(ChipName, "SLF") Then
'                            FileCopy App.Path & CurDevicePar.DeviceFolder & "\MSL\10_D1_K9F.BIN", App.Path & CurDevicePar.DeviceFolder & "\98SN\10_D1_K9F.BIN"
'                        Else
'                            FileCopy App.Path & CurDevicePar.DeviceFolder & "\normal\10_D1_K9F.BIN", App.Path & CurDevicePar.DeviceFolder & "\98SN\10_D1_K9F.BIN"
'                        End If
'                    End If
                    
'                    If Left(ChipName, 6) = "AU692A" Then
'                        ' add for copy MSL file
'                        If InStr(ChipName, "SLF") Then
'                            FileCopy App.Path & CurDevicePar.DeviceFolder & "\MSL\10_D1_K9F.BIN", App.Path & CurDevicePar.DeviceFolder & "\98SN\10_D1_K9F.BIN"
'                        Else
'                            FileCopy App.Path & CurDevicePar.DeviceFolder & "\normal\10_D1_K9F.BIN", App.Path & CurDevicePar.DeviceFolder & "\98SN\10_D1_K9F.BIN"
'                        End If
'                    End If
                    
'                ElseIf (Left(ChipName, 6) = "AU6991") And (OldVer_Flag = True) Then
'                        FileCopy App.Path & CurDevicePar.DeviceFolder & "\90old\90_D1_K9F.BIN", App.Path & CurDevicePar.DeviceFolder & "\90\90_D1_K9F.BIN"
'
'                ElseIf (Left(ChipName, 6) = "AU6991") And (OldVer_Flag = False) Then
'                        FileCopy App.Path & CurDevicePar.DeviceFolder & "\90new\90_D1_K9F.BIN", App.Path & CurDevicePar.DeviceFolder & "\90\90_D1_K9F.BIN"

                ' 20121227 是否需要old跟new dll尚待確認
                If (OldVer_Flag = False) And Left(ChipName, 6) = "AU6927" Then
                    FileCopy App.Path & CurDevicePar.DeviceFolder & "\newDLL\SetupDll.dll", App.Path & CurDevicePar.DeviceFolder & "\SetupDll.dll"

                ElseIf (OldVer_Flag = True) And Left(ChipName, 6) = "AU6927" Then
                    FileCopy App.Path & CurDevicePar.DeviceFolder & "\oldDLL\SetupDll.dll", App.Path & CurDevicePar.DeviceFolder & "\SetupDll.dll"

                End If
                
                .FT_ToolTitle = Trim(RS_Files.Fields("FT_ToolTitle"))
                
                If Len(Trim(RS_Files.Fields("Exec_Par1"))) Then
                    .Exec_Par1 = Trim(RS_Files.Fields("Exec_Par1"))
                End If
                
                If Len(Trim(RS_Files.Fields("Exec_Par2"))) Then
                    .Exec_Par2 = Trim(RS_Files.Fields("Exec_Par2"))
                End If
                
                Exit Do
            End If
            RS_Files.MoveNext
        End With
    Loop


CloseRS:
    
    RS_Par.Close
    If Not RS_Par Is Nothing Then Set RS_Par = Nothing
    RS_Files.Close
    If Not RS_Files Is Nothing Then Set RS_Files = Nothing
    If Not oCn Is Nothing Then Set oCn = Nothing
    
'    If CurDevicePar.ShortName = "AU6910" Or _
'       CurDevicePar.ShortName = "AU6915" Or _
'       CurDevicePar.ShortName = "AU6917" Or _
'       CurDevicePar.ShortName = "AU6988" Or _
'       CurDevicePar.ShortName = "AU6990" Then
'       CheckLED_Flag = True
'    Else
'        CheckLED_Flag = False
'    End If
    
    CheckLED_Flag = True
    
End Sub
 
Public Sub Begin_Sub()

Dim OldAlcorMPMessage As Long
Dim OldTime
Dim Bin2Counter As Integer
Dim buf
Dim mMsg As MSG
Dim AlcorMPMessage As Long
Dim rt2 As Long
Dim i As Long
Dim GPIBStatus As String

    Print "Begin"
    AllenTest = 0
    ChDir App.Path
    Do
        '============================================
        '             RS232 interface
        '============================================
GPIB_Flag:
        
        DoEvents
        MSComm1.InBufferCount = 0
        MSComm1.OutBufferCount = 0
        ChipName = ""
        TestResult = ""
        Print "Wait host command"
               
        Do
            MSComm1.Output = "Ready"
            Call MsecDelay(0.1)
            DoEvents
            buf = MSComm1.Input
            ChipName = ChipName & buf
            fnScsi2usb2K_KillEXE     'Clear removable device message box
        Loop Until (InStr(1, ChipName, "AU") <> 0 And Len(ChipName) >= 14) Or AllenStop = 1
            
        Cls
           
        If InStr(1, ChipName, "AUGPIB") <> 0 Then
            If SkipCtrlCount.Enabled = True Then
                GPIBStatus = "GPIBReadyGPIBReady"
            ElseIf SkipCtrlCount.Enabled = False Then
                GPIBStatus = "GPIBUNReadyGPIBUNReady"
            End If
              
            Do
                MSComm1.Output = GPIBStatus
                Call MsecDelay(0.02)
                DoEvents
                buf = MSComm1.Input
                ChipName = ChipName & buf
            Loop Until (InStr(1, ChipName, "AUGPIBACK") <> 0)
            GoTo GPIB_Flag
        End If
           
        If AllenStop = 1 Then
            End
        End If
          
        '==============================================
        '               begin Testing
        '==============================================
    
        OldTime = Timer   ' get timer
        ReMP_Flag = 0
        ReMP_Counter = 0
            
        If SkipCtrlCount.Enabled = True Then
            If SkipCtrlCount = "0" Then
                If Dir("D:\NoGPIB.PC") = "NoGPIB.PC" Then
                    Kill ("D:\NoGPIB.PC")
                End If
            Else
                SkipCtrlCount = CStr(CInt(SkipCtrlCount) - 1)
                If Dir("D:\NoGPIB.PC") = "" Then
                    Open "D:\NoGPIB.PC" For Output As #55
                    Call MsecDelay(0.02)
                    Close #55
                End If
            End If
        End If
        
        Dim temp As String
        Debug.Print ChipName
            
        
        If Left(ChipName, 5) <> "AU210" Then
            If Mid(ChipName, 11, 1) = "M" Then            'Force MP Function
                ForceMP_Flag = True
                
                temp = Right(ChipName, 4)
                temp = Replace(temp, "M", "L")
                
                ChipName = Mid(ChipName, 1, 10)
                ChipName = ChipName & temp
                
                Print "Force MP ..."
            ElseIf Mid(ChipName, 12, 1) = "M" Then
                ForceMP_Flag = True
                
                temp = Right(ChipName, 4)
                temp = Replace(temp, "M", "F")
                
                ChipName = Mid(ChipName, 1, 11)
                ChipName = ChipName & temp
                
                Print "Force MP ..."
            Else
                ForceMP_Flag = False
            End If
        End If
        
        Print "ChipName="; ChipName
        
        Debug.Print ChipName
        
        Select Case ChipName
            
            Case "AU6930XXXHLS52"
            
                CurDevicePar.ShortName = "AU6930"
                CurDevicePar.DeviceFolder = "\AlcorMP_6930\newMP"
                CurDevicePar.MP_LoadTitle = "698x UFD MP"
                CurDevicePar.MP_WorkTitle = "FT Mode , NormalMode , UFD MP Flash , Mode : Auto"
                CurDevicePar.MP_ToolFileName = "AlcorMP.exe"
                CurDevicePar.FT_ToolTitle = "UFD Test"

                Call AU6928XXXHLS50TestSub
            
            Case "AU6928XXXHLS52"
            
                CurDevicePar.ShortName = "AU6928"
                CurDevicePar.DeviceFolder = "\AlcorMP_6928\newMP"
                CurDevicePar.MP_LoadTitle = "698x UFD MP"
                CurDevicePar.MP_WorkTitle = "FT Mode , NormalMode , UFD MP Flash , Mode : Auto"
                CurDevicePar.MP_ToolFileName = "AlcorMP.exe"
                CurDevicePar.FT_ToolTitle = "UFD Test"
                'CurDevicePar.FT_ToolFileName = "UFDTest.exe"

                Call AU6928XXXHLS50TestSub
        
            Case "AU7310XTEELF20"
                Call AU7670XXXELF20TestSub
            
            ' FT2
            Case "AU3826A81AFF21", "AU3826A81BFF21", "AU3826A82AFF21", "AU3826A82BFF21", "AU3826B82AFF21", "AU3826B82BFF21", "AU3826C82BFF21", "AU3826C82AFF22", "AU3826C82BFF22", "AU3826C82CFF22", "AU3826C82DFF22"
                Call AU3826A81AFF20TestSub
                
            ' 殘電下地
            Case "AU3826C82AFF23", "AU3826C82BFF23", "AU3826C82CFF23", "AU3826C82DFF23", "AU3826D82BFF23", "AU3826E82BFF23", "AU3826F82AFF23", "AU3826F82BFF23"
                Call AU3826A81AFF23TestSub
                
            ' FT2+ST2
            Case "AU3826A81AFS11", "AU3826A81BFS11", "AU3826A82AFS11", "AU3826A82BFS11", "AU3826B82AFS11", "AU3826B82BFS11", "AU3826C82BFS11"
                Call AU3826A81AFS10TestSub
                
            ' ST2 (only test on/off for LDO <XTAL>)
            Case "AU3826A81AFS21", "AU3826A81BFS21", "AU3826A82AFS21", "AU3826A82BFS21", "AU3826B82AFS21", "AU3826B82BFS21", "AU3826C82BFS21", "AU3826C82AFS22", "AU3826C82BFS22", "AU3826C82CFS22", "AU3826C82DFS22", "AU3826D82BFS23", "AU3826E82BFS23", "AU3826F82BFS23"
                Call AU3826A81AFS20TestSub
                
            ' ST2 (only test on/off for LDO <LC>)
            Case "AU3826A81AFS24", "AU3826A81BFS24", "AU3826A82AFS24", "AU3826A82BFS24", "AU3826B82AFS24", "AU3826B82BFS24", "AU3826C82AFS24", "AU3826C82BFS24", "AU3826C82CFS24", "AU3826C82DFS24", "AU3826D82BFS24", "AU3826D82DFS24", "AU3826E82BFS24", "AU3826F82BFS24"
                Call AU3826A81AFS24TestSub
                
            ' FT3
            Case "AU3826A81AFF31", "AU3826A81BFF31", "AU3826A82AFF31", "AU3826A82BFF31", "AU3826B82AFF31", "AU3826B82BFF31", "AU3826D82BFF31"
                Call AU3826A81AFF30TestSub
                
            Case "AU6988D52HLF20", "AU6988D51HLF20", "AU6988D54HLF20"
                Call AU6988D52HLF20TestSub
                
            
            Case "AU6988D52ILF20", "AU6988D51ILF20", "AU6988D53ILF20", "AU6988D54ILF20"
                Call AU6988D52ILF20TestSub
            
            Case "AU6988D52HLF21", "AU6988D51HLF21" ', "AU6988D53HLF21"
                Call AU6988D52HLF21TestSub
            
            Case "AU6988D52HLF22", "AU6988D51HLF22", "AU6988D53HLF22", "AU6988D54HLF22"
                Call AU6988D52HLF22TestSub
            
            Case "AU6988D52HLF23"   'add unload driver
                Call AU6988D52HLF23TestSub
            
            Case "AU6988D52HLF24"   'add unload driver
                Call AU6988D52HLF24TestSub
            
            Case "AU6988D52HLF25"   'add unload driver
                Call AU6988D52HLF25TestSub
            
            Case "AU6988D52HLF26"   'add unload driver
                Call AU6988D52HLF26TestSub
            
            Case "AU6988D52HLF27"   'add unload driver
                Call AU6988D52HLF27TestSub
                
            '====================================================
            Case "AU6988H55ILF26", "AU6988G56ILF26"  'add unload driver
                Call AU6988H55ILF26TestSub
            
            Case "AU6988H55HLF26", "AU6988G56HLF26"  'add unload driver
                Call AU6988D52HLF26TestSub
            '========================================================
            Case "AU6988H55ILF28", "AU6988G56ILF28"  'add unload driver
                Call AU6988H55ILF28TestSub
            
            Case "AU6988H55HLF28", "AU6988G56HLF28"  'add unload driver
                Call AU6988H55HLF28TestSub
            '======================================================== for 20090904 version
            Case "AU6988D51HLF2A", "AU6988D52HLF2A", "AU6988D53HLF2A", "AU6988D54HLF2A", "AU6988G55HLF2A", "AU6988H56HLF2A"
                Call AU6988H56HLF2ATestSub
            
            Case "AU6988D52HLF29"
                Call AU6988D52HLF29TestSub
                
            Case "AU6988D52HLF2I"           'K9F1G
                Call AU6988D52HLF2ITestSub
                
            Case "AU6988D53HLF20"           'K9F1G
                Call AU6988D53HLF20TestSub
            
            Case "AU6988D51ILF2A", "AU6988D52ILF2A", "AU6988D53ILF2A", "AU6988D54ILF2A", "AU6988G55ILF2A", "AU6988H56ILF2A"
                Call AU6988H56ILF2ATestSub
            '========================================================
            Case "AU6992A52HLF20", "AU6992A51HLF20", "AU6992A53DLF20", "AU6992A54DLF20" 'from AU6992A52HLF20
                Call AU6992A52HLF20TestSub
                
            'New MP-Tool test
            '********************************
            Case "AU6992A52HLF2A", "AU6992A51HLF2A", "AU6992A53DLF2A", "AU6992A54DLF2A" 'from AU6992A51HLF2A
                Dual_Flag = True
                Call AU6992A51HLF2ATestSub
            
            Case "AU6996A51ILF2A"   'for NewMP tool AU6996A51ILF2A
                Call AU6996A51ILF2ATestSub
            '*********************************
            
            'New MP-Tool & MP by golden-sample
            '---------------------------------
            Case "AU6992A53DLF2B"
                Dual_Flag = True
                Call AU6992_MP_Golden
            
            Case "AU6996A51ILF2B"   'Just Modify AU6996Flash
                Dual_Flag = True
                Call AU6996_MP_Golden
            
            Case "AU6992A52HLS10"
                Call AU6992A52HLS10SortingSub    'shmoo from 3.6 to 3.0 step 0.05
            
            Case "AU6992A53DLF21"   'from AU6992A53HLF20 for "AU6992 SOCKET V1.1 FOR NS-6000" use
                Dual_Flag = True
                Call AU6992A53DLF21TestSub
            '=========================================================
            Case "AU7510A41ALF20", "AU7511A41BGF20", "AU7510A43ALF20", "AU7511A43BGF20"      'add unload driver
                Call AU7510A41ALF20TestSub
            
            Case "AU7511A45AGF20"
                Call AU7511A45AGF20TestSub
            
            Case "AU7510A45BGF20", "AU7510A45ALF20"
                ChipName = "AU7510A45BGF20"
                Call AU7510A45BGF20TestSub
            '==========================================================
            Case "AU6997A51BLF20" ' copy from AU6992 , only change MP tool
                Call AU6997A51BLF20TestSub
            
            Case "AU6996A51BLF20", "AU6996A51ILF20" ' copy from AU6992 , only change MP tool
                Call AU6996A51BLF20TestSub
            
            Case "AU6996B51ILF21", "AU6996C51ILF21"
                Call AU6996A51BLF21TestSub
            
            Case "AU6996B51ILF31"
                Call AU6996A51BLF31TestSub
            
            Case "AU3830A53ACF20"
                Call AU3830A53ACF20TestSub
            
            Case "AU3821B54CFF20"
                Call AU3821B54CFF20TestSub
            
            Case "AU3821A66FNF20", "AU3821A66INF20", "AU3821A66JNF20"
                Call AU3821A66XNF20TestSub
            
            Case "AU3821A66FNF21"
                Call AU3821A66FNF21TestSub
            
            'Case "AU3825A61BFF29"
            '    Call AU3825A61BFF29TestSub
            '    Bin2Counter = 0
            
             Case "AU3825A61BFS6E"
                
                UpdateFW_Btn.Visible = True
                Call AU3825A61BFS6ETestSub
                Bin2Counter = 0
                ContiUnknowFailCounter = 0
            
            Case "AU3825A61AFS6E"
                
                UpdateFW_Btn.Visible = True
                Call AU3825A61AFS6ETestSub
                Bin2Counter = 0
                ContiUnknowFailCounter = 0
                
            Case "AU3825A61BFF2F"
            
                UpdateFW_Btn.Visible = True
                Call AU3825A61BFF2FTestSub
                Bin2Counter = 0
                
             Case "AU3825D61BFF2G"
            
                UpdateFW_Btn.Visible = True
                Call AU3825D61BFF2GTestSub
                Bin2Counter = 0
                
            Case "AU3825D61BFE10"
            
                UpdateFW_Btn.Visible = True
                Call AU3825D61BFE10TestSub
                Bin2Counter = 0
                
            Case "AU3825A61BFImg"
            
                UpdateFW_Btn.Visible = True
                Call AU3825A61BFImgTestSub
                Bin2Counter = 0
            
            Case "AU3825A61BFS7E"
            
                UpdateFW_Btn.Visible = True
                Call AU3825A61BFS7ETestSub
                Bin2Counter = 0
                
            Case "AU3825A61AFF2F", "AU3825D61AFF2G"
                
                UpdateFW_Btn.Visible = True
                Call AU3825A61AFF2FTestSub
                Bin2Counter = 0
            
            Case "AU3825A61AFS7E"
            
                UpdateFW_Btn.Visible = True
                Call AU3825A61AFS7ETestSub
                Bin2Counter = 0
                
            Case "AU3825A61BFQ2E"
            
                UpdateFW_Btn.Visible = True
                Call AU3825A61BFQ2ETestSub
                Bin2Counter = 0
            
            Case "AU3825A61AFQ2E"

                UpdateFW_Btn.Visible = True
                Call AU3825A61AFQ2ETestSub
                Bin2Counter = 0
            
            Case "AU3825A61BFS31", "AU3825A61AFS31"
            
                Call AU3825A61SortingTestSub
                Bin2Counter = 0
            
            Case "AU3825A61XFS41"
                Call AU3825A61ST4TestSub    'phy-board test
                Bin2Counter = 0
                
            Case "AU2100A41DFM20", "AU2100A41DFF20", "AU2100A41BFM20", "AU2100A41BFF20", "AU2100A41CFM20", "AU2100A41CFF20"
                Call AU2100_NameSub
                Bin2Counter = 0
            
            Case "AU2101A41AFM20", "AU2101A41AFF20", "AU2101A41BFM20", "AU2101A41BFF20", "AU2101A41CFM20", "AU2101A41CFF20", "AU2101B41DFM20", "AU2101B41DFF20"
                Call AU2101_NameSub
                Bin2Counter = 0
                
            Case "AU2101DFFP0002", "AU2101HFFP1403", "AU2101HFFP1404", "AU2101HFFP1501"
                Call AU2101_FPXXXX_NameSub
                Bin2Counter = 0
                
            Case "AU2100CFFP0101"
                Call AU2100_ProgNameSub
                Bin2Counter = 0
            
            Case "AU6992A53HLF03", "AU6992B53HLF03", "AU6992R53HLF03", "AU6992S53HLF03"
                EQC_Flag = True
                Call AU6992EQCReNameSub
            
            Case "AU6992A53HLF0C", "AU6992B53HLF0C", "AU6992R53HLF0C", "AU6992S53HLF0C"
                EQC_Flag = True
                Dual_Flag = True
                Call AU6992EQCReNameDualSub
                
            Case "AU6992A51DLF23", "AU6992A52DLF23", "AU6992A53HLF23", "AU6992A54HLF23", "AU6992B53HLF23", "AU6992B54HLF23", "AU6992R53HLF23", "AU6992S53HLF23"
                Call AU6992HWSingleTestSub
          
            Case "AU6992A51DLF2C", "AU6992A52DLF2C", "AU6992A53HLF2C", "AU6992A54HLF2C", "AU6992B54HLF2C", "AU6992B53HLF2C", "AU6992R53HLF2C", "AU6992S53HLF2C"
                Dual_Flag = True
                Call AU6992HWGOLDTestSub
                            
            Case "AU6919C62HLS39", "AU6919C62HLS3A", "AU6919D62HLS39", "AU6919D62HLS3A"
                
                If (CurDevicePar.FullName = "") Or (CurDevicePar.FullName <> ChipName) Then
                    Call ReadCurDeviceInfo
                End If
                Dual_Flag = CurDevicePar.Dual_Flag
                EQC_Flag = CurDevicePar.EQC_Flag
                Call AU6919ST3TestSub
                ChDir App.Path
            
'            Case "AU6919C62HLS49", "AU6919C62HLS4A", "AU6919D62HLS49", "AU6919D62HLS4A", "AU6922B61HLS40", "AU6922B61HLS41"
'
'                If (CurDevicePar.FullName = "") Or (CurDevicePar.FullName <> ChipName) Then
'                    Call ReadCurDeviceInfo
'                End If
'                Dual_Flag = CurDevicePar.Dual_Flag
'                EQC_Flag = CurDevicePar.EQC_Flag
'                Call AU6919ST4TestSub
'                ChDir App.Path
'
'            Case "AU6919A61HLS49", "AU6919A61HLS4A", "AU6919B61HLS49", "AU6919B61HLS4A"
'
'                If (CurDevicePar.FullName = "") Or (CurDevicePar.FullName <> ChipName) Then
'                    Call ReadCurDeviceInfo
'                End If
'                Dual_Flag = CurDevicePar.Dual_Flag
'                EQC_Flag = CurDevicePar.EQC_Flag
'                Call AU6919ST4TestSub
'                ChDir App.Path
'
'            Case "AU6916B61HLS48", "AU6916B61HLS49", "AU6919J62HLS49", "AU6919N62HLS49", "AU6922I62HLS40"
'
'                If (CurDevicePar.FullName = "") Or (CurDevicePar.FullName <> ChipName) Then
'                    Call ReadCurDeviceInfo
'                End If
'                Dual_Flag = CurDevicePar.Dual_Flag
'                EQC_Flag = CurDevicePar.EQC_Flag
'                Call AU6919ST4TestSub               ' 共用st4流程
'                ChDir App.Path
'
'            Case "AU6922I61HLS40", "AU6922I62HLS40"
'
'                If (CurDevicePar.FullName = "") Or (CurDevicePar.FullName <> ChipName) Then
'                    Call ReadCurDeviceInfo
'                End If
'                Dual_Flag = CurDevicePar.Dual_Flag
'                EQC_Flag = CurDevicePar.EQC_Flag
'                Call AU6919ST4TestSub               ' 共用st4流程
'                ChDir App.Path
            
            Case "AU6913H62HLS11", "AU6913H63HLS11"
                
                If (CurDevicePar.FullName = "") Or (CurDevicePar.FullName <> ChipName) Then
                    Call ReadCurDeviceInfo
                    If FailCloseAP Then
                        FailCloseAP = False
                        FailCloseFT.Value = 0
                    End If
                    
                End If
                Dual_Flag = CurDevicePar.Dual_Flag
                EQC_Flag = CurDevicePar.EQC_Flag
                Call AU6913ST1TestSub
                ChDir App.Path
            
            Case "AU6991C61HLS11", "AU6922B61HLS11", "AU6922G61HLS11", "AU6922F61HLS11", "AU6991C61HLS13"
                
                If (CurDevicePar.FullName = "") Or (CurDevicePar.FullName <> ChipName) Then
                    Call ReadCurDeviceInfo
                    If FailCloseAP Then
                        FailCloseAP = False
                        FailCloseFT.Value = 0
                    End If
                    
                End If
                Dual_Flag = CurDevicePar.Dual_Flag
                EQC_Flag = CurDevicePar.EQC_Flag
                Call AU69XXLC_ST1TestSub
                ChDir App.Path
            
            Case "AU9420BYDOTP10"
            
                Call AU9420OTPSub
            
            Case "AU9420A41ALF20", "AU9420A42BLF20"
                Call AU9420TestSub
                
            Case "AU9420A42BLF00"
                Call AU9420EQCTestSub
                
            Case "AU9420A42ASF20"
            
                Call AU9420ReadEEPROMTestSub
                
            Case "AU6621C83CFF21"
            
                Call AU6621C83CFF21Testsub
                ContiUnknowFailCounter = 0  'skip use DevCon.exe rescan PCI bus
                
            Case "AU6621E84CFF21"
            
                Call AU6621E84CFF21Testsub
                ContiUnknowFailCounter = 0  'skip use DevCon.exe rescan PCI bus
                
            Case "AU6601A71CLF21"
            
                Call AU6601A71CLF21Testsub
                ContiUnknowFailCounter = 0  'skip use DevCon.exe rescan PCI bus
                
            Case "AU6601A71CFF30"
                Call AU6601A71CFF30TestSub
                ContiUnknowFailCounter = 0  'skip use DevCon.exe rescan PCI bus
            
            Case "AU6621C83CFF30", "AU6621C83GFF30"
                Call AU6621C83CFF30Testsub
                ContiUnknowFailCounter = 0  'skip use DevCon.exe rescan PCI bus
            
            Case "AU6621E84CFF30", "AU6621E84GFF30"
                Call AU6621E84CFF30Testsub
                ContiUnknowFailCounter = 0  'skip use DevCon.exe rescan PCI bus
            
            Case "AU87100XXXXFF20", "AU87100XXXXFF21"

                If Dir(App.Path & "\AlcorMP_8710\Hub_Info_Normal.ini") <> "" Then
                    Kill App.Path & "\AlcorMP_8710\Hub_Info_Normal.ini"
                End If

                U3_Test = True
                If (CurDevicePar.FullName = "") Or (CurDevicePar.FullName <> ChipName) Then
                    Call ReadCurDeviceInfo
                End If
                
                SetSiteStatus (RunU2)
                WaitAnotherSiteDone (RunU2)
                
                ' test U2
                Call AU69XX_StdTestSub
                
                SetSiteStatus (U2Done)
                WaitAnotherSiteDone (U2Done)
                
                ForceMP_Flag = False
                
                ' test U3
                If TestResult = "PASS" Then
                    Call AU69XX_StdTestSub
                    U2_Pass = False
                End If
                
                ForceMP_Flag = False
                
                ChDir App.Path
                
'                SetSiteStatus (U3Done)
'                WaitAnotherSiteDone (U3Done)
'                MsecDelay (0.1)
            
            Case "AU87100XXXXFF31"

                If Dir(App.Path & "\AlcorMP_8710\Hub_Info_Normal.ini") <> "" Then
                    Kill App.Path & "\AlcorMP_8710\Hub_Info_Normal.ini"
                End If

                U3_Test = True
                If (CurDevicePar.FullName = "") Or (CurDevicePar.FullName <> ChipName) Then
                    Call ReadCurDeviceInfo
                End If
                
                ' test U2
                Call AU87100HV_LVTestSub
                
                ForceMP_Flag = False
                
                ' test U3
                If TestResult = "PASS" Then
                    Call AU87100HV_LVTestSub
                    U2_Pass = False
                End If
                
                ForceMP_Flag = False
                
                ChDir App.Path
            
            Case Else
            
                If (CurDevicePar.FullName = "") Or (CurDevicePar.FullName <> ChipName) Then
                    Call ReadCurDeviceInfo
                End If
            
            
                If CurDevicePar.FullName <> "" Then
                    Dual_Flag = CurDevicePar.Dual_Flag
                    EQC_Flag = CurDevicePar.EQC_Flag
  
                    If EQC_Flag Then
                        Call AU69XXHV_LVTestSub
                    Else
                        If InStr(CurDevicePar.FullName, "S4") <> 0 Then ' for 共用ST4流程
                            AU6919ST4TestSub
                        Else
                            Call AU69XX_StdTestSub
                        End If
                    End If
                End If
                ChDir App.Path
        End Select
          
        If Dual_Flag <> True Then
            Dual_Flag = False
        End If
          
        Print "TestTime:"; Timer - OldTime
        
        '2012/7/17: Reset root-hub
        If TestResult = "Bin2" Then
            ContiUnknowFailCounter = ContiUnknowFailCounter + 1
                
            If ContiUnknowFailCounter >= 4 Then
                If InStr(1, ChipName, "7670") = 0 Then                  '7670 don't reset hub
                    ResetReturn = Shell(ResetHubString, vbNormalFocus)
                    WaitProcQuit (ResetReturn)
                    ContiUnknowFailCounter = 0
                    MsecDelay (2#)
                End If
            End If
        Else
            ContiUnknowFailCounter = 0
        End If
        
        '===============================================
        '                end Testing
        '==============================================
        Label2.Caption = "Bin2 " & CStr(Bin2Counter)
        MSComm1.Output = TestResult
        Call MsecDelay(0.1) 'arch add
        Print "TestResult :"; TestResult
        MSComm1.InBufferCount = 0
        MSComm1.InputLen = 0
        '===============================================
        '                Reset condition
        '==============================================
        
        
        ' for AU6988 : shut down condition
        If Bin2.Value = 1 Then
             If TestResult = "Bin2" Then
                Bin2Counter = Bin2Counter + 1
             Else
                Bin2Counter = 0
             End If
            
            If Bin2Counter >= 5 Then
               Call MsecDelay(2)
               Shell "cmd /c shutdown -r  -t 0", vbHide
            End If
        
        End If
        
        ' for AU7510  : shut down condition
        If ResetFlag = 1 Then ' for AU7510
            Call MsecDelay(2)
            Shell "cmd /c shutdown -r  -t 0", vbHide
        End If
        fnScsi2usb2K_KillEXE ' clear removable device message box
    Loop While AllenTest = 0
    
End Sub

Private Sub AutoMP_Option_Click()

    If AutoMP_Option.Value = 1 Then
        LoadMP.Caption = "Auto MP"
    Else
        LoadMP.Caption = "Load MP"
    End If
    
End Sub

Private Sub Command1_Click()
    
    Cls
    Call Begin_Sub
    
End Sub

Private Sub Command2_Click()

    AllenTest = 1

End Sub


Private Sub FailCloseFT_Click()

    If FailCloseFT.Value = 1 Then
        FailCloseAP = True
    Else
        FailCloseAP = False
    End If
    
End Sub

Private Sub Form_Activate()
    
    Call Command1_Click
    
End Sub

Private Sub Form_Load()
Dim ProgramName() As String
Dim U As Integer, i As Integer, j As Integer
Dim MyExeDate() As String
Dim ShowMyDate() As String
Dim intC As Integer
Dim strName As String

    U2_Pass = False
    U3_Pass = False
    U3_Test = False
    
    AU7670 = False
    
    If App.EXEName <> "MPTester" Or CheckMe(Me) Then
        End
    End If

    If Dir(App.Path & "\AlcorMP_6928\RC_Value-Pass.txt") = "RC_Value-Pass.txt" Then
        Kill (App.Path & "\AlcorMP_6928\RC_Value-Pass.txt")
    End If
    
    If Dir(App.Path & "\AlcorMP_6928\RC_Value-Fail.txt") = "RC_Value-Fail.txt" Then
        Kill (App.Path & "\AlcorMP_6928\RC_Value-Fail.txt")
    End If

    ' Get folder name for program name
    ProgramName = Split(App.Path, "\")
    U = UBound(ProgramName)
    For i = 0 To U
        If Left(ProgramName(U - 1), 3) = "New" Then
            'Label1.Caption = "2012" & Right(ProgramName(U - 1), 4)
            Label1.Caption = (Val(Mid(ProgramName(U - 1), 20, 3)) + 1911) & Right(ProgramName(U - 1), 4)
            Exit For
        Else
            MyExeDate = Split(FileDateTime(App.Path), " ")
            ShowMyDate = Split(MyExeDate(0), "/")
            For j = LBound(ShowMyDate) To UBound(ShowMyDate)
                If Len(ShowMyDate(j)) < 2 Then
                    ShowMyDate(j) = "0" & ShowMyDate(j)
                End If
                Label1.Caption = Label1.Caption & ShowMyDate(j)
            Next
            Exit For
        End If
    Next

    
    MSComm1.CommPort = 1
    'MSComm1.PortOpen = False
    MSComm1.Settings = "9600,N,8,1"
    MSComm1.PortOpen = True
    
    MSComm1.InBufferCount = 0
    MSComm1.InputLen = 0
    Bin2.Value = 0
    ForceMP_Flag = False
    FailCloseAP = True
    'NonEcc_Flag = False
 
    GPIBCard_Exist = False
    GPIBCard_Exist = CheckGPIB()
    ResetHubString = App.Path & "\devcon restart @USB\ROOT_HUB*"

    SkipCtrlCount = "0"

    If GPIBCard_Exist = False Then
        GPIBCARD_Label.Caption = "No GPIB"
        GPIBCARD_Label.ForeColor = &HFF&
        SkipCtrlCount.Enabled = False
        
        If Dir("D:\NoGPIB.PC") = "" Then
            Open "D:\NoGPIB.PC" For Output As #55
            Call MsecDelay(0.02)
            Close #55
        End If
    Else
        GPIBCARD_Label.Caption = "GPIB Exist"
        GPIBCARD_Label.ForeColor = &HFF0000
        If Dir("D:\NoGPIB.PC") = "NoGPIB.PC" Then
            Kill ("D:\NoGPIB.PC")
        End If
    End If
    
    GetSystemInfo CPUInfo 'Get CPU Number for AU3825A61 sorting test save log file

End Sub

Private Sub Form_Unload(Cancel As Integer)

    AllenStop = 1

End Sub

Private Sub HLV_Test_Click()

    Call ShellExecute(MPTester.hwnd, "open", App.Path & "\Hi_Lo_V_Test.exe", "", "", SW_SHOW)
    Unload Me
    
End Sub

Private Sub LoadMP_Click()
    
    Cls
    Call AutoMP_sub
    
End Sub

Private Sub LoadRWTest_Click()

    Call LoadRWTest_Click_AU6990

End Sub

Private Sub StartMP_Click()

    Call StartScan_Click_AU7510

End Sub

Private Sub StartRWTest_Click()

    Call StartRWTest_Click_AU6988

End Sub

Private Sub UpdateFW_Btn_Click()
        
    UpdateFW_Btn.Enabled = False
    
    Call CloseVedioCap
    cardresult = DO_WritePort(card, Channel_P1A, &HFA) 'Open ENA Power 1111_1010 (Bit3 using External clock)

    Call MsecDelay(0.2)
    If Not WaitDevOn("vid_058f") Then
        MPTester.TestResultLab = "UnKnow Device"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        Exit Sub
    End If
    Call MsecDelay(0.8)
    
    If OldChipName = "AU3825A61AFF2D" Then
        Call Load_AU3825_28QFN_FW_Update
    ElseIf OldChipName = "AU3825A61BFF2D" Then
        Call Load_AU3825_40QFN_FW_Update
    End If

    KillProcess ("MPTool_lite_v3.12.620.exe")
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    
    UpdateFW_Btn.Enabled = True
End Sub
