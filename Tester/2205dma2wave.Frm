VERSION 5.00
Begin VB.Form dma2205_2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FF00&
   Caption         =   "2205/2206 AI DMA"
   ClientHeight    =   9240
   ClientLeft      =   1020
   ClientTop       =   5625
   ClientWidth     =   13215
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9240
   ScaleMode       =   0  'User
   ScaleWidth      =   13215
   Begin VB.Frame dmaFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "DMA Input(A/D)"
      ForeColor       =   &H00FF0000&
      Height          =   4815
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   4680
      Width           =   11595
      Begin VB.PictureBox picScreen 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   3675
         Index           =   1
         Left            =   720
         ScaleHeight     =   243
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   679
         TabIndex        =   11
         Top             =   480
         Width           =   10215
      End
      Begin VB.Label lblLeft 
         Caption         =   "0"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label lblRight 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   195
         Index           =   1
         Left            =   5520
         TabIndex        =   18
         Top             =   3120
         Width           =   915
      End
      Begin VB.Label labValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "-10.0"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   17
         Top             =   3840
         Width           =   600
      End
      Begin VB.Label labValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "-5.00"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   600
      End
      Begin VB.Label labValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "0.00"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label labValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "5.00"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label labValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "10.0"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   5
         Left            =   45
         TabIndex        =   13
         Top             =   420
         Width           =   600
      End
      Begin VB.Label labVots 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Volt"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   195
         Width           =   435
      End
   End
   Begin VB.Frame dmaFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "DMA Input(A/D)"
      ForeColor       =   &H00FF0000&
      Height          =   4815
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11595
      Begin VB.PictureBox picScreen 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   3675
         Index           =   0
         Left            =   720
         ScaleHeight     =   243
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   679
         TabIndex        =   1
         Top             =   480
         Width           =   10215
      End
      Begin VB.Label labVots 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Volt"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   195
         Width           =   435
      End
      Begin VB.Label labValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "10.0"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   8
         Top             =   420
         Width           =   600
      End
      Begin VB.Label labValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "5.00"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label labValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "0.00"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label labValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "-5.00"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   600
      End
      Begin VB.Label labValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "-10.0"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   3840
         Width           =   600
      End
      Begin VB.Label lblRight 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   5520
         TabIndex        =   3
         Top             =   3120
         Width           =   915
      End
      Begin VB.Label lblLeft 
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   2280
         Width           =   795
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit!"
   End
End
Attribute VB_Name = "dma2205_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit
 

 
Private Sub btnStart_Click()

End Sub


Private Sub cobTrigL_Level_Click()
 ' If CInt(cobTrigL_Level.Text) > trigLevel Then
 '   MsgBox "L_TrigLevel can't be larger than H_TrigLevel."
'    cobTrigL_Level.ListIndex = trigLLevel / 8
    'cobTrigL_Level.Text = CStr(trigLLevel)
 ' Else
 '   DrawTrigLevel trigLevel, CInt(cobTrigLevel.Text), trigLLevel, CInt(cobTrigL_Level.Text)
'    trigLLevel = CInt(cobTrigL_Level.Text)
'End If
 End Sub

Public Sub ShowData(ByVal startidx As Long, ByVal acq_count As Long)
    Dim i As Long, k As Long, n As Long
    Dim w As Long, h As Long
    Dim X As Long, Y As Long
    Dim y1 As Long, last_y1 As Long
    Dim last_x As Long, last_y As Long
    Dim samp As Integer, trig_pos As Long
    Dim samp1 As Integer
    Dim Data As Double
    picScreen(0).Cls
    picScreen(1).Cls
    'DrawTrigLevel -1, trigLevel
    DrawTrigLevel -1, 184, -1, 0
   ' lblLeft.Caption = CStr(CLng(scrollView.value) * DISP_COUNT)
   ' lblRight.Caption = CStr(CLng(lblLeft) + DISP_COUNT - 1)
    w = picScreen(0).ScaleWidth
    h = picScreen(0).ScaleHeight
 
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
        samp = gnBuffer(i)
        
    
        Y = CInt((CSng(32767 - CLng(samp)) / CSng(65535)) * h)
       
   '
     '   Debug.Print x, y, gnBuffer(i)
         If last_y <> -1 Then
             picScreen(0).Line (last_x, last_y)-(X, Y), &HFFFF&
             
         End If
        last_x = X
        last_y = Y
        last_y1 = y1
    Next
    
   
If RecordModeMP3 = 21 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_21(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen(0).Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If


If RecordModeMP3 = 22 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_22(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen(0).Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If


If RecordModeMP3 = 23 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_23(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen(0).Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If
 
'======================================================
'
'
'
'======================================================


  
    'DrawTrigLevel -1, trigLevel
   ' DrawTrigLevel -1, 184, -1, 0
   ' lblLeft.Caption = CStr(CLng(scrollView.value) * DISP_COUNT)
   ' lblRight.Caption = CStr(CLng(lblLeft) + DISP_COUNT - 1)
    w = picScreen(1).ScaleWidth
    h = picScreen(1).ScaleHeight
 
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
        samp = gnBufferWMA(i)
        
    
        Y = CInt((CSng(32767 - CLng(samp)) / CSng(65535)) * h)
       
   '
     '   Debug.Print x, y, gnBuffer(i)
         If last_y <> -1 Then
             picScreen(1).Line (last_x, last_y)-(X, Y), &HFFFF&
             
         End If
        last_x = X
        last_y = Y
        last_y1 = y1
    Next
 
 
 If RecordModeWMA = 24 Then
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3WMA_23(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen(1).Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If
 
 
 If RecordModeWMA = 241 Then
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3WMA_231(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen(1).Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If
  
 
If RecordModeWMA = 242 Then
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3WMA_232(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen(1).Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If
 
End Sub

Private Sub Form_Load()
ShowData 0, 100
End Sub

Private Sub Form_Paint()
ShowData 0, 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim result As Long
  Card2214Exist = 0
   If Card2214 >= 0 Then
      D2K_Release_Card (Card2214)
   End If
End Sub

Private Sub mnuExit_Click()
  If card >= 0 Then D2K_Release_Card (card)
  End
End Sub


'Private Sub DrawTrigLevel(ByVal old_level As Integer, ByVal level As Integer)
 '   Dim w As Long, h As Long
  '  Dim Y As Long
   ' w = picScreen.ScaleWidth
    'h = picScreen.ScaleHeight
    'If old_level <> -1 Then
     '   Y = h * (255 - old_level) / 255
      '  picScreen.Line (0, Y)-(w, Y), &H0
    'End If
    'Y = h * (255 - level) / 255
    'picScreen.Line (0, Y)-(w, Y), &H909090
'End Sub
Private Sub DrawTrigLevel(ByVal old_level As Integer, ByVal level As Integer, ByVal old_llevel As Integer, ByVal llevel As Integer)
    Dim w As Long, h As Long
    Dim Y As Long

    w = picScreen(0).ScaleWidth
    h = picScreen(0).ScaleHeight
    If old_level <> -1 Then
        Y = h * (255 - old_level) / 255
        picScreen(0).Line (0, Y)-(w, Y), &H0
    End If
    If old_llevel <> -1 Then
        Y = h * (255 - old_llevel) / 255
        picScreen(0).Line (0, Y)-(w, Y), &H0
    End If
    Y = h * (255 - level) / 255
    picScreen(0).Line (0, Y)-(w, Y), &H909090
    Y = h * (255 - llevel) / 255
    picScreen(0).Line (0, Y)-(w, Y), &HBBBBBB
End Sub


Private Function Min(ByVal v1 As Long, ByVal v2 As Long) As Long
  If v1 < v2 Then
    Min = v1
  Else
    Min = v2
  End If
End Function

