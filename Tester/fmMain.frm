VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fmMain 
   Caption         =   "Alcor Smartcard Manufacture Application"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   Icon            =   "fmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMsg 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   8295
   End
   Begin VB.CommandButton cmdStartTest 
      Caption         =   "Run(&R)"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdStopTest 
      Caption         =   "Stop(&S)"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3120
      Top             =   7800
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   3225
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11853
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close(&X)"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   6855
   End
End
Attribute VB_Name = "fmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public bStop As Boolean
Public strAppPath As String
Public lngLineNum As Long
Public bReadyToClose As Boolean





Private Sub cmdAT45D041_Click()
fmAT45D041_1.Show vbModal
End Sub


Private Sub cmdClose_Click()
bReadyToClose = True
Unload fmMain
End Sub











Private Sub cmdStartTest_Click()
    cmdStartTest.Enabled = False
    cmdStopTest.Enabled = True
    cmdClose.Enabled = False

    Call StartTest
End Sub

Private Sub cmdStopTest_Click()
bStop = True
cmdStartTest.Enabled = True
cmdStopTest.Enabled = False
cmdClose.Enabled = True

End Sub


Private Sub Form_Load()
     
    Dim lngResult As Long
    
    bReadyToClose = False
    bCurSlotNum = 0
    fmMain.Caption = fmMain.Caption & "V" & App.Major & "." & App.Minor & "." & App.Revision & " (Win2000)"
    strAppPath = App.Path
    If (Right(strAppPath, 1) <> "\") Then
        strAppPath = strAppPath & "\"
    End If
 
    ' Connect to the smartcard resource manager
    lngResult = SCardEstablishContext(SCARD_SCOPE_SYSTEM, lngNull, lngNull, lngContext)
    If (lngResult <> 0) Then
        MsgBox ("SCardEstablishContext Failed")
        GoTo Error_Exit
    End If
    
    strReaderName = VenderString & " " & ProductString & " 0"
    udtReaderStates(0).szReader = strReaderName & vbNullChar

    IsReaderLost = True
    lngResult = ConnectAlcorReader
    'Timer1.Enabled = True 'Start the timer to polling the state of the reader and the card
    cmdStartTest.Enabled = True
    cmdStopTest.Enabled = False
    

    Exit Sub
    
Error_Exit:
    End
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Dim lngResult As Long
   
    If (bReadyToClose = False) Then
        Cancel = True
        Exit Sub
    End If
    
    'Scarddisconnect will cold reset the card and then power down the card.
    'if the EEPROM card presents, the function will take several seconds to cold reset it 3 times(because of time out).
    'for save time to close the program, this function can be deleted from the code, and the pc/sc subsystem will
    ' automatically do the function after the program is terminated.
    lngResult = SCardDisconnect(lngCard, SCARD_LEAVE_CARD) 'disconnect the connection with the reader
    lngResult = SCardReleaseContext(lngContext) ' release the connection with the resource manager
End Sub

Public Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case Chr(KeyCode)
        Case "R", "r"

            If cmdStartTest.Enabled = True Then Call cmdStartTest_Click

        Case "S", "s"
            If cmdStopTest.Enabled = True Then Call cmdStopTest_Click

        Case "X", "x"
            If cmdClose.Enabled = True Then Call cmdClose_Click

    End Select
End Sub
Public Sub Form_Keydown(KeyCode As Integer, Shift As Integer)
    Select Case Chr(KeyCode)
        Case "R", "r"
            
            If cmdStartTest.Enabled = True Then cmdStartTest.SetFocus
        
        Case "S", "s"
            If cmdStopTest.Enabled = True Then cmdStopTest.SetFocus
            
        Case "X", "x"
            If cmdClose.Enabled = True Then cmdClose.SetFocus
            
    End Select

End Sub


Public Sub StartTest()
    ' length of list of readers
    Dim lngReadersLen As Long
    Dim lngReaderstatelen As Long
    ' array of readerstate with only one element
    Dim lngResult As Long
       ' cardinality of readerstate array
    Dim lngReaderNameLen As Long
    Dim lngSCardStatus As Long
    Dim ATR As ARRAY_TYPE
    Dim ATRLen As Long
    Dim i As Long
    
    bStop = False
    
    Do
    
        lngReaderstatelen = 1
        udtReaderStates(0).dwCurrentState = SCARD_STATE_UNAWARE
        ' check if the reader state was changed
        lngResult = SCardGetStatusChangeA(lngContext, 0, udtReaderStates(0), _
                                            lngReaderstatelen)
        DoEvents
        DoEvents
        DoEvents
        If (lngResult = 0) Then ' a reader presents
            'ReportCardStatus udtReaderStates(0).dwEventState
            'lblATRstring.Caption = ""
            sbStatus.Panels(1).Text = udtReaderStates(0).szReader
            Call AlcorCardTest
            
            'This reader was complete the test, wait for reader unplug
            Do
                DoEvents
                DoEvents
                DoEvents
                lngReaderstatelen = 1
                udtReaderStates(0).dwCurrentState = SCARD_STATE_UNAWARE
                ' check if the reader state was changed
                lngResult = SCardGetStatusChangeA(lngContext, 0, udtReaderStates(0), _
                                                lngReaderstatelen)
                If (lngResult <> 0) Then
                    ShowResult ("CLEAR")
                    txtMsg.Text = ""
                    fnScsi2usb2K_KillEXE
                    Exit Do
                End If
                    
            Loop While bStop = False
                
        Else ' no reader found
            
'            Call EvaluateResult("SCardGetStatusChangeA", lngResult)
            sbStatus.Panels(1).Text = "No Reader"
            sbStatus.Panels(2).Text = ""
           ' lblATRstring.Caption = ""
'            End
    
        End If
        DoEvents
        DoEvents
        DoEvents
    Loop While bStop = False
    
End Sub

Public Function VerifyAlcorReader() As Boolean
    
    Dim lngResult As Long
    Dim OutBuffer(1) As Byte
    Dim InBuffer(1) As Byte
    
    VerifyAlcorReader = False
    lngActiveProtocol = 0
    OutBuffer(0) = ASYNCHRONOUS_CARD_MODE
            
    lngResult = Alcor_SwitchCardMode(lngCard, 0, OutBuffer(0))
    If (lngResult <> 0) Then
        Exit Function
    End If
    
    VerifyAlcorReader = True

End Function


Private Sub EvaluateResult(strFnName As String, lngResult As Long)
    ' output string
    Dim strOut As String
    ' string for text output of error value
    Dim strErrorValue As String
    ' leading zero padding for low hex values
    Dim strLeader As String
    Dim strMsg As String
    Dim lngOverallResult As Long
    Dim lngHexResultLen As Long
    strLeader = "00000000"
    
    ' print function name and pass/fail result
    strOut = strFnName
    
    If SCARD_S_SUCCESS = lngResult Then
        ' concatenate PASS tagline to output
 '       OutputLine strOut + strPass
    Else
        ' mark overall result flag
        lngOverallResult = SCARD_F_UNKNOWN_ERROR
        ' concatenate FAIL tagline to output
        strMsg = strOut + strFail
        ' convert result decimal number to hex string
        strErrorValue = Hex(lngResult)
        ' pad with leading zeros if necessary
        lngHexResultLen = Len(strErrorValue)
        strErrorValue = strTab + "Error: 0x" _
                        + Left(strLeader, 8 - lngHexResultLen) + strErrorValue
        strMsg = strMsg & vbCrLf & strErrorValue
        ' give message associated with error code
        strOut = ApiErrorMessage(lngResult)
        If 0 <> Len(strOut) Then
            strMsg = strMsg & vbCrLf & strTab + strOut
        End If
        MsgBox strMsg
    End If
    
End Sub ' EvaluateResult

Public Sub AlcorCardTest()

    If (AlcorFindTheCard() = False) Then
        ShowResult ("FAIL")
        Exit Sub
    End If
    
    If (ProcessScriptTest() = False) Then
        ShowResult ("FAIL")
    Else
        ShowResult ("PASS")
    End If
    
    
End Sub

Public Function AlcorFindTheCard() As Boolean

    Dim lngReaderstatelen As Long
    Dim lngResult As Long
    Dim i As Long
    Dim strTmp As String
    
    AlcorFindTheCard = False
    
    Do
        lngReaderstatelen = 1
        udtReaderStates(0).dwCurrentState = SCARD_STATE_PRESENT 'SCARD_STATE_UNAWARE
        ' check if the reader state was changed
        lngResult = SCardGetStatusChangeA(lngContext, 0, udtReaderStates(0), _
                                                lngReaderstatelen)
        DoEvents
        DoEvents
        DoEvents
        'ReportCardStatus udtReaderStates(0).dwEventState
        If (lngResult <> 0) Then ' a reader not present,exit the sub
            OutMsg "Error! Can't find reader"
            Exit Do
        End If
        
        'If card is unrecognized, exit the sub
        If (udtReaderStates(0).dwEventState And SCARD_STATE_EMPTY) Then
'            OutMsg "No Card"
'            Exit Do
        ElseIf (udtReaderStates(0).dwEventState And SCARD_STATE_MUTE) Then
           OutMsg "Error! Unrecognized card"
           Exit Do
        
        Else
            
            For i = 0 To udtReaderStates(0).cbAtr - 1 'print the ATR string
                strTmp = strTmp & (Byte2Char(udtReaderStates(0).rgbAtr(i)) & " ")
            Next i
            OutMsg strTmp
            
            AlcorFindTheCard = True
            Exit Do
        End If
'        ReportCardStatus udtReaderStates(0).dwEventState
'        MsgBox udtReaderStates(0).rgbAtr
        
        DoEvents
        DoEvents
        DoEvents
        
        '    lblATRstring.Caption = ""

        
    Loop While bStop = False

End Function
Public Function WaitCardChange() As Boolean

    Dim lngReaderstatelen As Long
    Dim lngResult As Long
    Dim i As Long
    Dim strTmp As String
    
    WaitCardChange = False
    
    Do
        lngReaderstatelen = 1
        udtReaderStates(0).dwCurrentState = SCARD_STATE_PRESENT 'SCARD_STATE_UNAWARE
        ' check if the reader state was changed
        lngResult = SCardGetStatusChangeA(lngContext, 0, udtReaderStates(0), _
                                                lngReaderstatelen)
        DoEvents
        DoEvents
        DoEvents
        'ReportCardStatus udtReaderStates(0).dwEventState
        If (lngResult <> 0) Then ' a reader not present,exit the sub
            OutMsg "Error! Can't find reader"
            Exit Do
        End If
        
        'If card is unrecognized, exit the sub
        If (udtReaderStates(0).dwEventState And SCARD_STATE_EMPTY) Then
            OutMsg "Find Card Out"
            
            WaitCardChange = True
            Exit Do
        End If
'        ReportCardStatus udtReaderStates(0).dwEventState
'        MsgBox udtReaderStates(0).rgbAtr
        
        DoEvents
        DoEvents
        DoEvents
        
        '    lblATRstring.Caption = ""

        
    Loop While bStop = False

End Function


Public Function ReportCardStatus(lngStatus As Long) As Long
    Dim strCardStatus As String
    
    strCardStatus = ""
    If (lngStatus And SCARD_STATE_IGNORE) Then
        strCardStatus = strCardStatus & "SCARD_STATE_IGNORE" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_CHANGED) Then
        strCardStatus = strCardStatus & "SCARD_STATE_CHANGED" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_UNKNOWN) Then
        strCardStatus = strCardStatus & "SCARD_STATE_UNKNOWN" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_UNAVAILABLE) Then
        strCardStatus = strCardStatus & "SCARD_STATE_UNAVAILABLE" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_EMPTY) Then
        strCardStatus = strCardStatus & "SCARD_STATE_EMPTY" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_PRESENT) Then
        strCardStatus = strCardStatus & "SCARD_STATE_PRESENT" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_ATRMATCH) Then
        strCardStatus = strCardStatus & "SCARD_STATE_ATRMATCH" & vbCrLf
    End If
    
    If (lngStatus And SCARD_STATE_EXCLUSIVE) Then
        strCardStatus = strCardStatus & "SCARD_STATE_EXCLUSIVE" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_INUSE) Then
        strCardStatus = strCardStatus & "SCARD_STATE_INUSE" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_MUTE) Then
        strCardStatus = strCardStatus & "SCARD_STATE_MUTE" & vbCrLf
    End If
    
    
    MsgBox strCardStatus
    
    

End Function

Public Sub ShowResult(strResult As String)

    If (strResult = "PASS") Then
        lblResult.Caption = "PASS"
        lblResult.ForeColor = vbGreen
    ElseIf (strResult = "FAIL") Then
        lblResult.Caption = "FAIL"
        lblResult.ForeColor = vbRed
    ElseIf (strResult = "CLEAR") Then
        lblResult.Caption = ""
        lblResult.ForeColor = vbGreen
    End If
        
  
End Sub

Public Function ProcessScriptTest() As Boolean
    Dim strOneLine As String

    ProcessScriptTest = True
    Open strAppPath & "AlcorEMV.ini" For Input As #1
    lngLineNum = 0
    
    
    Do
        If (GetOneValidLine(strOneLine) = False) Then
            
            Exit Do
        End If
        
        Select Case (strOneLine)
            Case "[Card Protocol]"
                If (DoCardProtocol() = False) Then
                    ProcessScriptTest = False
                    Exit Do
                End If
               
            
            Case "[ATR]"
                If (CheckCardATR() = False) Then
                    OutMsg "Fail to check ATR"
                    ProcessScriptTest = False
                    Exit Do
                End If
                OutMsg "Check ATR PASS!"
            
            Case "[XFR]"
                If (DoCardXfrTest() = False) Then
                    OutMsg "Fail to Xfr Test"
                    ProcessScriptTest = False
                    Exit Do
                End If
                OutMsg "Check card command PASS"

            
            Case "[ChangeCard]"
                lblResult.Caption = "Change card"
                lblResult.ForeColor = vbBlue
                If (WaitCardChange() = False) Then
                    OutMsg "Error occurs when waiting card out"
                    ProcessScriptTest = False
                    Exit Do
                End If

                If (AlcorFindTheCard() = False) Then
                    OutMsg "Error occurs when waiting card out"
                    ProcessScriptTest = False
                    Exit Do
                End If
                
            Case Else
                OutMsg "Illigal line in AlcorEMV.ini, Line " & lngLineNum
                ProcessScriptTest = False
                Exit Do
        End Select
        
    Loop While (Not EOF(1))
    Close #1
    
End Function

Public Function CheckCardATR() As Boolean
    Dim i As Long
    Dim strTmp As String
    Dim strOneLine As String
    Dim lngReaderstatelen
    Dim lngResult As Long
    
    '=========================
    'Get ATR from smartcard service
    '=========================

    CheckCardATR = False
    lngReaderstatelen = 1
    udtReaderStates(0).dwCurrentState = SCARD_STATE_UNAWARE
    ' check if the reader state was changed
    lngResult = SCardGetStatusChangeA(lngContext, 0, udtReaderStates(0), _
                                                lngReaderstatelen)
    DoEvents
    DoEvents
    DoEvents
    'ReportCardStatus udtReaderStates(0).dwEventState
    If (lngResult <> 0 Or udtReaderStates(0).cbAtr = 0) Then ' a reader not present,exit the sub
        OutMsg "Fail to call SCardGetStatusChangeA in CheckCardATR()"
        Exit Function
    End If

    For i = 0 To udtReaderStates(0).cbAtr - 1 'print the ATR string
        strTmp = strTmp & (Byte2Char(udtReaderStates(0).rgbAtr(i)) & " ")
    Next i
    strTmp = Trim(strTmp)
    
    '=========================
    'Get ATR from file
    '=========================
    If (GetOneValidLine(strOneLine) = False) Then
        OutMsg "Can't get protocol setting in AlcorEMV.ini"
        Exit Function
    End If
   
    '=========================
    'Compare ATR
    '=========================
    If (strOneLine <> strTmp) Then
        OutMsg "The ATR not match!"
        OutMsg "Expect : " & strOneLine
        OutMsg "Actual : " & strTmp
        Exit Function
    End If

    CheckCardATR = True
        
End Function

Public Function DoCardProtocol() As Boolean
    Dim strOneLine As String
    
    DoCardProtocol = False
    If (GetOneValidLine(strOneLine) = False) Then
        OutMsg "Can't get protocol setting in AlcorEMV.ini"
        Exit Function
    End If
    lngPreProtocol = 0
    Select Case (strOneLine)
        Case "Auto"
            lngPreProtocol = SCARD_PROTOCOL_T0 Or SCARD_PROTOCOL_T1
            DoCardProtocol = True
        Case "T0"
            lngPreProtocol = SCARD_PROTOCOL_T0
            DoCardProtocol = True
        Case "T1"
            lngPreProtocol = SCARD_PROTOCOL_T1
            DoCardProtocol = True
        Case Else
            OutMsg "Unknown protocol setting :" & strOneLine
            Exit Function
    End Select
    Call OutMsg("Card Protocol : " & strOneLine)
End Function

Public Function DoCardXfrTest() As Boolean

    Dim i As Long
    Dim strCommand As String
    Dim lngCommandLen As Long
    Dim aCommand As ARRAY_TYPE
    Dim aResponse As ARRAY_TYPE
    Dim lngReplyLen As Long
    Dim lngReplyBufferLen As Long
    Dim lngResult As Long
    Dim pioSendPci As SCARD_IO_REQUEST
    Dim strReplyData As String
    Dim strOneLine As String
    Dim lngTestNo As Long
    Dim lNumOfTest As Long

    lngTestNo = 0
    DoCardXfrTest = False
    
    'Get the number of test to do
    If (GetOneValidLine(strOneLine) = False) Then
         OutMsg "Can't get protocol setting in AlcorEMV.ini"
    End If
    strOneLine = Trim(strOneLine)
    If (Mid(strOneLine, 1, 2) = "0x" Or Mid(strOneLine, 1, 2) = "0X") Then
        lNumOfTest = Val("&H" & Mid(strOneLine, 3, Len(strOneLine) - 2))
    Else
        lNumOfTest = strOneLine
    End If
    
    
    '=========================
    'Connect to the smartcard
    '=========================
    lngResult = SCardConnectA(lngContext, strReaderName, SCARD_SHARE_SHARED, lngPreProtocol, lngCard, lngActiveProtocol)
    
    If (lngResult <> 0) Then
        OutMsg "Can't connect to the card"
        Exit Function
    End If
    If (VerifyAlcorReader() = False) Then
        MsgBox "Not Alcor's reader"
        Exit Function
    End If
    
    lngResult = SCardBeginTransaction(lngCard)
    
    '=========================
    'Get command from file
    '=========================
    While (lngTestNo <> lNumOfTest)
        If (GetOneValidLine(strOneLine) = False) Then
            OutMsg "Invalid line in AlcorEMV.ini"
            Exit Function
        End If
        lngTestNo = lngTestNo + 1
        OutMsg "Command Test Case : " & lngTestNo
        strCommand = strOneLine
        lngCommandLen = (Len(strCommand) + 1) / 3

        For i = 1 To lngCommandLen
            aCommand.byteData(i - 1) = Val("&H" & Mid(strCommand, (i - 1) * 3 + 1, 2))
        Next i
        lngReplyLen = 258
        pioSendPci.dwProtocol = lngActiveProtocol
        pioSendPci.dbPciLength = Len(pioSendPci)
    
        lngResult = SCardTransmit(lngCard, pioSendPci, aCommand.byteData(0), lngCommandLen, pioSendPci, aResponse.byteData(0), lngReplyLen)
    
        DoEvents
        DoEvents
        DoEvents
        'txtCardResponse.Text = ""
        If (lngResult <> 0) Then
            OutMsg "Error! Fail to send smartcard data !"
            OutMsg "Cmd : " & strOneLine
            lngResult = SCardEndTransaction(lngCard, SCARD_LEAVE_CARD)
            Exit Function
        End If
        strReplyData = ""
        For i = 1 To lngReplyLen
            strReplyData = strReplyData & Byte2Char(aResponse.byteData(i - 1)) & " "
        Next i
        strReplyData = Trim(strReplyData)
        
        '=========================
        'Get the expected data from file
        '=========================
        If (GetOneValidLine(strOneLine) = False) Then
            OutMsg "Can't get protocol setting in AlcorEMV.ini"
            GoTo ExitFun
        End If
        
         '=========================
        'compare the result
        '=========================
        If (strOneLine <> strReplyData) Then
            OutMsg "Error! Card response compare error!"
            OutMsg "Expect : " & strOneLine
            OutMsg "Actual : " & strReplyData
            GoTo ExitFun
        End If
        DoEvents
        DoEvents
        DoEvents
    Wend
    
    DoCardXfrTest = True
    
ExitFun:
    lngResult = SCardEndTransaction(lngCard, SCARD_LEAVE_CARD)
    lngResult = SCardDisconnect(lngCard, SCARD_LEAVE_CARD)

End Function

Public Function GetOneValidLine(ByRef strOneLine As String) As Boolean
Dim str As String
    While (Not EOF(1))
        Line Input #1, strOneLine
        lngLineNum = lngLineNum + 1
        strOneLine = Trim(strOneLine)
'        str = Mid(strOneLine, 1, 1)
        If (Mid(strOneLine, 1, 1) <> ";" And strOneLine <> "") Then
            strOneLine = strOneLine
            GetOneValidLine = True
            Exit Function
        End If
    Wend
    
    GetOneValidLine = False

End Function

Public Sub OutMsg(strMsg As String)
    txtMsg.Text = txtMsg.Text & strMsg & vbCrLf
    
End Sub

