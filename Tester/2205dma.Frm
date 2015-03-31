VERSION 5.00
Begin VB.Form dma2205 
   Appearance      =   0  '平面
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FF00&
   Caption         =   "2205/2206 AI DMA"
   ClientHeight    =   5970
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
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   5970
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   13215
   Begin VB.Frame dmaFrame 
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0C0&
      Caption         =   "DMA Input(A/D)"
      ForeColor       =   &H00FF0000&
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11595
      Begin VB.PictureBox picScreen 
         Appearance      =   0  '平面
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   4395
         Left            =   720
         ScaleHeight     =   291
         ScaleMode       =   3  '像素
         ScaleWidth      =   679
         TabIndex        =   1
         Top             =   360
         Width           =   10215
      End
      Begin VB.Label lblLeft 
         Caption         =   "0"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label lblRight 
         Alignment       =   1  '靠右對齊
         Caption         =   "0"
         Height          =   195
         Left            =   5520
         TabIndex        =   8
         Top             =   3120
         Width           =   915
      End
      Begin VB.Label labValue 
         Alignment       =   1  '靠右對齊
         Appearance      =   0  '平面
         BackColor       =   &H00C0C0C0&
         Caption         =   "-10.0"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   4560
         Width           =   600
      End
      Begin VB.Label labValue 
         Alignment       =   1  '靠右對齊
         Appearance      =   0  '平面
         BackColor       =   &H00C0C0C0&
         Caption         =   "-5.00"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   6
         Top             =   3360
         Width           =   600
      End
      Begin VB.Label labValue 
         Alignment       =   1  '靠右對齊
         Appearance      =   0  '平面
         BackColor       =   &H00C0C0C0&
         Caption         =   "0.00"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   5
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label labValue 
         Alignment       =   1  '靠右對齊
         Appearance      =   0  '平面
         BackColor       =   &H00C0C0C0&
         Caption         =   "5.00"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label labValue 
         Alignment       =   1  '靠右對齊
         Appearance      =   0  '平面
         BackColor       =   &H00C0C0C0&
         Caption         =   "10.0"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   3
         Top             =   420
         Width           =   600
      End
      Begin VB.Label labVots 
         Appearance      =   0  '平面
         BackColor       =   &H00C0C0C0&
         Caption         =   "Volt"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   195
         Width           =   435
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   11760
      TabIndex        =   11
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label txtXfer 
      Alignment       =   1  '靠右對齊
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   6000
      Width           =   2115
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit!"
   End
End
Attribute VB_Name = "dma2205"
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

Public Sub ShowDataAU6476(ByVal startidx As Long, ByVal acq_count As Long)
    Dim i As Long, k As Long, n As Long
    Dim w As Long, h As Long
    Dim X As Long, Y As Long
    Dim y1 As Long, last_y1 As Long
    Dim last_x As Long, last_y As Long
    Dim samp As Integer, trig_pos As Long
    Dim samp1 As Integer
    Dim Data As Double
    Dim VSum As Single
    Dim VData As Single
    picScreen.Cls
    'DrawTrigLevel -1, trigLevel
    DrawTrigLevel -1, 184, -1, 0
   ' lblLeft.Caption = CStr(CLng(scrollView.value) * DISP_COUNT)
   ' lblRight.Caption = CStr(CLng(lblLeft) + DISP_COUNT - 1)
    w = picScreen.ScaleWidth
    h = picScreen.ScaleHeight
 
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
        samp = ADCBuffer(i)
         VData = CSng(samp + 1) / CSng(32768) * 5#
         VSum = VSum + VData
     
        Debug.Print VData
    
        Y = CInt((CSng(32767 - CLng(samp)) / CSng(65535)) * h)
       
   '
     '   Debug.Print x, y, gnBuffer(i)
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y)-(X, Y), &HFFFF&
             
         End If
        last_x = X
        last_y = Y
        last_y1 = y1
    Next
    
 ' Label1.Caption = CStr(VSum / acq_count / CSng(32768) * 5#)
  Label1.Caption = CStr(VSum / acq_count)
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
    picScreen.Cls
    'DrawTrigLevel -1, trigLevel
    DrawTrigLevel -1, 184, -1, 0
   ' lblLeft.Caption = CStr(CLng(scrollView.value) * DISP_COUNT)
   ' lblRight.Caption = CStr(CLng(lblLeft) + DISP_COUNT - 1)
    w = picScreen.ScaleWidth
    h = picScreen.ScaleHeight
 
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
        samp = gnBuffer(i)
        
    
        Y = CInt((CSng(32767 - CLng(samp)) / CSng(65535)) * h)
       
   '
     '   Debug.Print x, y, gnBuffer(i)
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y)-(X, Y), &HFFFF&
             
         End If
        last_x = X
        last_y = Y
        last_y1 = y1
    Next
    
   
If RecordMode = 0 Then
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If
 
 If RecordMode = 2 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_1(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If
    
    
 If RecordMode = 3 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_2(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If

 If RecordMode = 4 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_3(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If


If RecordMode = 5 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_4(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If


If RecordMode = 6 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_5(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If


If RecordMode = 7 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_6(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If

If RecordMode = 8 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_7(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If

If RecordMode = 9 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_8(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If


If RecordMode = 10 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_9(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If

'================================================================
    
If RecordMode = 11 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_10(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If


'==============================================================


If RecordMode = 12 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_11(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If


If RecordMode = 13 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_12(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If


If RecordMode = 14 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_13(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If

 If RecordMode = 15 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_15(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If
   
  If RecordMode = 16 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_16(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If

 If RecordMode = 161 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_161(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If
   
   
   If RecordMode = 17 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_17(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If


   
If RecordMode = 18 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_18(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If


If RecordMode = 19 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_19(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If


If RecordMode = 20 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_20(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If


If RecordMode = 21 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_21(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If


If RecordMode = 22 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_22(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If


If RecordMode = 23 Then  ' FOR au3130 2ND MODE
     
     For i = 0 To acq_count - 1
         
     '    n = (startidx + i + startPos) Mod count1
        X = (i * w) / (acq_count)
         
        samp1 = MP3Data_23(i)
    
       
        y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
   '
        
         If last_y <> -1 Then
             picScreen.Line (last_x, last_y1)-(X, y1), &HFF00&
             
         End If
        last_x = X
      
        last_y1 = y1
    Next
End If
    'Set trigger position
  '  Select Case cobTrigMode.ListIndex
  '    Case DAQ2K_AI_TRGMOD_POST / 8
  '      trig_pos = 0
  '    Case DAQ2K_AI_TRGMOD_PRE / 8
  '      trig_pos = count1 - 1
  '    Case DAQ2K_AI_TRGMOD_DELAY / 8
  '      trig_pos = -1
  '    Case DAQ2K_AI_TRGMOD_MIDL / 8
  '      trig_pos = count1 - CLng(txtPostCount.Text) * ch_cnt
  '  End Select
  '  If CLng(lblLeft) <= trig_pos And CLng(lblRight) >= trig_pos Then
  '    TriggerPos.Visible = True
  '    TriggerPos.Left = (picScreen.Left - 60) + ((trig_pos - CLng(lblLeft)) * w * Screen.TwipsPerPixelX) / (DISP_COUNT - 1)
  '  Else
  '    TriggerPos.Visible = False
  '  End If
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

    w = picScreen.ScaleWidth
    h = picScreen.ScaleHeight
    If old_level <> -1 Then
        Y = h * (255 - old_level) / 255
        picScreen.Line (0, Y)-(w, Y), &H0
    End If
    If old_llevel <> -1 Then
        Y = h * (255 - old_llevel) / 255
        picScreen.Line (0, Y)-(w, Y), &H0
    End If
    Y = h * (255 - level) / 255
    picScreen.Line (0, Y)-(w, Y), &H909090
    Y = h * (255 - llevel) / 255
    picScreen.Line (0, Y)-(w, Y), &HBBBBBB
End Sub


Private Function Min(ByVal v1 As Long, ByVal v2 As Long) As Long
  If v1 < v2 Then
    Min = v1
  Else
    Min = v2
  End If
End Function

