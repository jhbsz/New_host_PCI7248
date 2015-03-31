VERSION 5.00
Begin VB.Form fmAsynCard 
   Caption         =   "Asynchronous Card Operation"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   Icon            =   "fmAsynCard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCardResponse 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   3120
      Width           =   5055
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "fmAsynCard.frx":1CCA
      Top             =   7200
      Width           =   3495
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox txtLength 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Text            =   "110"
      Top             =   6840
      Width           =   735
   End
   Begin VB.CommandButton cmdStressTest 
      Caption         =   "Stress Test"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdSendAPDU 
      Caption         =   "Send APDU"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox txtAPDU 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Text            =   "A0 A4 00 00 02 3F 00"
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label lblHistory 
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   3720
      X2              =   5380
      Y1              =   6740
      Y2              =   6740
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   3705
      X2              =   5365
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Label Label1 
      Caption         =   "Length"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   6960
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   5400
      Y1              =   6135
      Y2              =   6135
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   5400
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "APDU Command"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "Card Response"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   2880
      Width           =   1815
   End
End
Attribute VB_Name = "fmAsynCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StressStop As Boolean


Private Sub cmdClose_Click()
fmMain.Timer1.Enabled = True
StressStop = True
Unload fmAsynCard

End Sub

Private Sub cmdSendAPDU_Click()

    Dim i As Long
    Dim strCommand As String
    Dim lngCommandLen As Long
    Dim aCommand As ARRAY_TYPE
    Dim aResponse As ARRAY_TYPE
    Dim lngReplyLen As Long
    Dim lngReplyBufferLen As Long
    Dim lngresult As Long
    Dim pioSendPci As SCARD_IO_REQUEST

    
    lngresult = SCardBeginTransaction(lngCard)
    
    strCommand = Trim(txtAPDU.Text)
    lngCommandLen = (Len(strCommand) + 1) / 3
    
    For i = 1 To lngCommandLen
    aCommand.byteData(i - 1) = Val("&H" & Mid(strCommand, (i - 1) * 3 + 1, 2))
    Next i
    lngReplyLen = 258
    pioSendPci.dwProtocol = lngActiveProtocol   'SCARD_PROTOCOL_T0
    pioSendPci.dbPciLength = Len(pioSendPci)

    lngresult = SCardTransmit(lngCard, pioSendPci, aCommand.byteData(0), lngCommandLen, pioSendPci, aResponse.byteData(0), lngReplyLen)


    txtCardResponse.Text = ""
    If (lngresult = 0) Then
        For i = 1 To lngReplyLen
            txtCardResponse.Text = txtCardResponse.Text & Byte2Char(aResponse.byteData(i - 1)) & " "
        Next i
        lngresult = MsgBox("Send command successfully !", vbOKOnly)
    Else
        lngresult = MsgBox("Send command error !", vbOKOnly)
    End If


    lngresult = SCardEndTransaction(lngCard, SCARD_LEAVE_CARD)

End Sub

Private Sub cmdStop_Click()
    cmdStressTest.Enabled = True
    fmMain.Timer1.Enabled = True
StressStop = True
End Sub

Private Sub cmdStressTest_Click()
    Dim i As Long
    Dim strCommand As String
    Dim lngCommandLen As Long
    Dim bOutBuffer(300) As Byte
    Dim bReplyBuffer(300) As Byte
    Dim refBuffer(300) As Byte
    Dim lngReplyLen As Long
    Dim lngReplyBufferLen As Long
    Dim lngresult As Long
    Dim pioSendPci As SCARD_IO_REQUEST
    Dim CompareCorrect As Boolean
    Dim TotalNo As Long
    Dim ErrorNo As Long
    Dim strTmp As String
    Dim lngWriteLen As Long
    Dim lngOutBufferLen As Long
        
    fmMain.Timer1.Enabled = False
    cmdStressTest.Enabled = False
    
    txtCardResponse.Text = ""
    StressStop = False
    TotalNo = 0
    ErrorNo = 0
    Open "Error.txt" For Output As #1
    
    lngWriteLen = 110
    lngOutBufferLen = Val(txtLength.Text)
'    lngResult = SCardBeginTransaction(lngCard)
    Do While (StressStop = False)
        TotalNo = TotalNo + 1
        CompareCorrect = True
    '========================
    'Send Write command
        For i = 0 To 299
            bReplyBuffer(i) = 0
            refBuffer(i) = 0
            bOutBuffer(i) = 0
        Next i
        
        For i = 1 To lngWriteLen
            refBuffer(i - 1) = Int(255 * Rnd)
        Next
        bOutBuffer(0) = &H0
        bOutBuffer(1) = &HD6
        bOutBuffer(2) = &H82
        bOutBuffer(3) = &H0
        bOutBuffer(4) = lngWriteLen
        For i = 1 To lngWriteLen
            bOutBuffer(i + 4) = refBuffer(i - 1)
        Next
        
            
        lngReplyLen = 2
        pioSendPci.dwProtocol = lngActiveProtocol   'SCARD_PROTOCOL_T0
        pioSendPci.dbPciLength = Len(pioSendPci)
    
        lngresult = SCardTransmit(lngCard, pioSendPci, bOutBuffer(0), lngWriteLen + 5, pioSendPci, bReplyBuffer(0), lngReplyLen)
        
        
        If (lngresult <> 0) Then
            txtCardResponse.Text = txtCardResponse.Text & "P"
            ErrorNo = ErrorNo + 1
            Print #1, "Loop: " & ErrorNo & "    SCardTransmit(Update binary) Error : 0x" & Hex(lngresult) & vbCrLf
            
        ElseIf (bReplyBuffer(0) <> &H90 Or bReplyBuffer(1) <> 0) Then
            txtCardResponse.Text = txtCardResponse.Text & "S"
            ErrorNo = ErrorNo + 1
            Print #1, "Loop: " & ErrorNo & "    SCardTransmit(Update binary) Error : 0x" & Hex(lngresult) & vbCrLf
            
        Else
        
       '=======================================================
       'Send Read command
        
            bOutBuffer(0) = &H0
            bOutBuffer(1) = &HB0
            bOutBuffer(2) = &H82
            bOutBuffer(3) = &H0
            bOutBuffer(4) = Val(txtLength.Text)
            lngReplyLen = 258
            pioSendPci.dwProtocol = lngActiveProtocol   'SCARD_PROTOCOL_T0
            pioSendPci.dbPciLength = Len(pioSendPci)
            lngOutBufferLen = 5
        
        
            lngresult = SCardTransmit(lngCard, pioSendPci, bOutBuffer(0), lngOutBufferLen, pioSendPci, bReplyBuffer(0), lngReplyLen)
            
            If (lngresult <> 0) Then
                txtCardResponse.Text = txtCardResponse.Text & "P"
                 ErrorNo = ErrorNo + 1
                Print #1, "Loop: " & TotalNo & "    SCardTransmit(Read binary) Error : 0x" & Hex(lngresult) & vbCrLf
            ElseIf (lngReplyLen <> (Val(txtLength.Text) + 2)) Then
                ErrorNo = ErrorNo + 1
                Print #1, "Loop " & TotalNo & "  Status Bytes =" & Byte2Char(bReplyBuffer(lngReplyLen - 2)) & Byte2Char(bReplyBuffer(lngReplyLen - 1))
            Else
            
                For i = 1 To Val(txtLength.Text)
                    If (refBuffer(i - 1) <> bReplyBuffer(i - 1)) Then
                        CompareCorrect = False
                        Exit For
                    End If
                Next i
                If (CompareCorrect = True) Then
                    txtCardResponse.Text = txtCardResponse.Text & "."
                Else
                    txtCardResponse.Text = txtCardResponse.Text & "X"
                    ErrorNo = ErrorNo + 1
                    Print #1, "Loop: " & ErrorNo & vbCrLf
                    strTmp = ""
                    For i = 0 To Val(txtLength.Text) - 1
                        strTmp = strTmp & Byte2Char(refBuffer(i)) & " "
                    Next i
                    Print #1, "Update Binary " & strTmp
                    
                    strTmp = ""
                    For i = 0 To Val(txtLength.Text) - 1
                        strTmp = strTmp & Byte2Char(bReplyBuffer(i)) & " "
                    Next i
                    Print #1, "Read Binary " & strTmp
                    Print #1, vbCrLf
                   
                    
                End If
            
            End If
        End If
        
        lblHistory.Caption = "Loop: " & TotalNo & "   " & "Error: " & ErrorNo
        
        If (Len(txtCardResponse.Text) = 400) Then
            txtCardResponse.Text = ""
        End If
            
        If (ErrorNo <> 0) Then
            lngresult = SCardDisconnect(lngCard, SCARD_UNPOWER_CARD)
            
            Exit Do
        End If
        
            
        DoEvents
    Loop
    
        
        
    Close #1
'    lngResult = SCardEndTransaction(lngCard, SCARD_LEAVE_CARD)

End Sub

Private Sub Form_Load()
StressStop = True
End Sub
