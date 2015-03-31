VERSION 5.00
Begin VB.Form dma2205_4 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FF00&
   Caption         =   "mp3 + WMA R-L Noise"
   ClientHeight    =   9630
   ClientLeft      =   3015
   ClientTop       =   5625
   ClientWidth     =   13410
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
   ScaleHeight     =   9630
   ScaleMode       =   0  'User
   ScaleWidth      =   13410
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   4320
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   13215
   End
   Begin VB.CommandButton Enlarge 
      Caption         =   ">>"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Shrink 
      Caption         =   "<<"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2115
      Index           =   3
      Left            =   120
      ScaleHeight     =   139
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   879
      TabIndex        =   3
      Top             =   7440
      Width           =   13215
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2115
      Index           =   2
      Left            =   120
      ScaleHeight     =   139
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   879
      TabIndex        =   2
      Top             =   5280
      Width           =   13215
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2115
      Index           =   1
      Left            =   120
      ScaleHeight     =   139
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   879
      TabIndex        =   1
      Top             =   2160
      Width           =   13215
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1995
      Index           =   0
      Left            =   120
      ScaleHeight     =   131
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   879
      TabIndex        =   0
      Top             =   120
      Width           =   13215
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit!"
   End
End
Attribute VB_Name = "dma2205_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit
Dim BeginPoint As Long
Dim DataCount As Long
 

 
Private Sub btnStart_Click()

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
    Dim j As Integer
    Dim jj As Integer
    
    '===========================
    ' draw triggle level
    '===========================
    DrawTrigLevel -1, 184, -1, 0
   
 
    '1. for picture 0
    
    For j = 0 To 1  ' for MP3 R-L, WMA R-L
    
    w = picScreen(j).ScaleWidth
    h = picScreen(j).ScaleHeight
        
        
        '===========================
        ' draw raw data
        '===========================
         For i = 0 To acq_count - 1
          
            X = (i * w) / (acq_count)
            samp = MP3RawData(j, i)    ' j:mean picture index, i: data array
        
            Y = CInt((CSng(32767 - CLng(samp)) / CSng(65535)) * h)
    
             If last_y <> -1 Then
                 picScreen(j).Line (last_x, last_y)-(X, Y), &HFFFF&
                 
             End If
            last_x = X
            last_y = Y
            last_y1 = y1
        Next
    
          
        '===========================
        ' draw Audio Pattern
        '===========================
         For i = 0 To acq_count - 1
          
            X = (i * w) / (acq_count)
            samp1 = MP3Pattern(MP3MatchIndex, j, i)
          '    samp1 = SineCurve(j, i)
            y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
         
             If last_y <> -1 Then
                 picScreen(j).Line (last_x, last_y1)-(X, y1), &HFF00&
                 
             End If
            last_x = X
          
            last_y1 = y1
        Next
        
    
    
    Next
  
  
  
  '=================================================================
  '
  '   WMA
  '
  '================================================================
  
  
      '1. for picture 0
    
    For j = 0 To 1  ' for WMA R-L, WMA R-L
    jj = j + 2
    w = picScreen(jj).ScaleWidth
    h = picScreen(jj).ScaleHeight
        
        
        '===========================
        ' draw raw data
        '===========================
         For i = 0 To acq_count - 1
          
            X = (i * w) / (acq_count)
            samp = WMARawData(j, i)  ' jj:mean picture index, i: data array
        
            Y = CInt((CSng(32767 - CLng(samp)) / CSng(65535)) * h)
    
             If last_y <> -1 Then
                 picScreen(jj).Line (last_x, last_y)-(X, Y), &HFFFF&
                 
             End If
            last_x = X
            last_y = Y
            last_y1 = y1
        Next
    
          
        '===========================
        ' draw Audio Pattern
        '===========================
         For i = 0 To acq_count - 1
          
            X = (i * w) / (acq_count)
             samp1 = WMAPattern(WMAMatchIndex, j, i)
           ' samp1 = SineCurve2(j, i)
            y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
         
             If last_y <> -1 Then
                 picScreen(jj).Line (last_x, last_y1)-(X, y1), &HFF00&
                 
             End If
            last_x = X
          
            last_y1 = y1
        Next
        
    
    
    Next
  
  
 
End Sub
Public Sub ShowData3(ByVal startidx As Long, ByVal acq_count As Long)
    Dim i As Long, k As Long, n As Long
    Dim w As Long, h As Long
    Dim X As Long, Y As Long
    Dim y1 As Long, last_y1 As Long
    Dim last_x As Long, last_y As Long
    Dim samp As Integer, trig_pos As Long
    Dim samp1 As Integer
    Dim Data As Double
    Dim j As Integer
    Dim jj As Integer
    
    '===========================
    ' draw triggle level
    '===========================
    DrawTrigLevel -1, 184, -1, 0
   
 
    '1. for picture 0
    
    For j = 0 To 1  ' for MP3 R-L, WMA R-L
    
    w = picScreen(j).ScaleWidth
    h = picScreen(j).ScaleHeight
        
        
        '===========================
        ' draw raw data
        '===========================
         For i = 0 To acq_count - 1
          
            X = (i * w) / (acq_count)
            samp = MP3RawData(j, i)    ' j:mean picture index, i: data array
        
            Y = CInt((CSng(32767 - CLng(samp)) / CSng(65535)) * h)
    
             If last_y <> -1 Then
                 picScreen(j).Line (last_x, last_y)-(X, Y), &HFFFF&
                 
             End If
            last_x = X
            last_y = Y
            last_y1 = y1
        Next
    
          
        '===========================
        ' draw Audio Pattern
        '===========================
         For i = 0 To acq_count - 1
          
            X = (i * w) / (acq_count)
       '     samp1 = MP3Pattern(MP3MatchIndex, j, i)
              samp1 = SineCurve(j, i)
            y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
         
             If last_y <> -1 Then
                 picScreen(j).Line (last_x, last_y1)-(X, y1), &HFF00&
                 
             End If
            last_x = X
          
            last_y1 = y1
        Next
        
    
    
    Next
  
  
  
  '=================================================================
  '
  '   WMA
  '
  '================================================================
  
  
      '1. for picture 0
    
    For j = 0 To 1  ' for WMA R-L, WMA R-L
    jj = j + 2
    w = picScreen(jj).ScaleWidth
    h = picScreen(jj).ScaleHeight
        
        
        '===========================
        ' draw raw data
        '===========================
         For i = 0 To acq_count - 1
          
            X = (i * w) / (acq_count)
            samp = WMARawData(j, i)  ' jj:mean picture index, i: data array
        
            Y = CInt((CSng(32767 - CLng(samp)) / CSng(65535)) * h)
    
             If last_y <> -1 Then
                 picScreen(jj).Line (last_x, last_y)-(X, Y), &HFFFF&
                 
             End If
            last_x = X
            last_y = Y
            last_y1 = y1
        Next
    
          
        '===========================
        ' draw Audio Pattern
        '===========================
         For i = 0 To acq_count - 1
          
            X = (i * w) / (acq_count)
         '   samp1 = WMAPattern(WMAMatchIndex, j, i)
            samp1 = SineCurve2(j, i)
            y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
         
             If last_y <> -1 Then
                 picScreen(jj).Line (last_x, last_y1)-(X, y1), &HFF00&
                 
             End If
            last_x = X
          
            last_y1 = y1
        Next
        
    
    
    Next
  
  
 
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

Public Sub ShowData2(ByVal startidx As Long, ByVal acq_count As Long)

   If (startidx + acq_count) > 6000 Then
   MsgBox "out space!!"
   Exit Sub
   End If
 
    Dim i As Long, k As Long, n As Long
    Dim w As Long, h As Long
    Dim X As Long, Y As Long
    Dim y1 As Long, last_y1 As Long
    Dim last_x As Long, last_y As Long
    Dim samp As Integer, trig_pos As Long
    Dim samp1 As Integer
    Dim Data As Double
    Dim j As Integer
    Dim jj As Integer
    
    '===========================
    ' draw triggle level
    '===========================
    DrawTrigLevel -1, 184, -1, 0
   
 
    '1. for picture 0
    
    For j = 0 To 1  ' for MP3 R-L, WMA R-L
    
    w = picScreen(j).ScaleWidth
    h = picScreen(j).ScaleHeight
        
        
        '===========================
        ' draw raw data
        '===========================
         For i = 0 To acq_count - 1
                      
                      
            X = (i * w) / (acq_count)
            samp = MP3RawData(j, startidx + i)  ' j:mean picture index, i: data array
        
            Y = CInt((CSng(32767 - CLng(samp)) / CSng(65535)) * h)
    
             If last_y <> -1 Then
                 picScreen(j).Line (last_x, last_y)-(X, Y), &HFFFF&
                 
             End If
            last_x = X
            last_y = Y
            last_y1 = y1
        Next
    
          last_y = -1
        '===========================
        ' draw Audio Pattern
        '===========================
         For i = 0 To acq_count - 1
          
            X = (i * w) / (acq_count)
        '    samp1 = MP3Pattern(MP3MatchIndex, j, startidx + i)
            samp1 = SineCurve(j, i)
            y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
         
                If i > 0 Then
                 picScreen(j).Line (last_x, last_y1)-(X, y1), &HFF00&
                 
             End If
            last_x = X
          
            last_y1 = y1
        Next
        
    
    
    Next
  
  
  
  '=================================================================
  '
  '   WMA
  '
  '================================================================
  
  
      '1. for picture 0
    
    For j = 0 To 1  ' for WMA R-L, WMA R-L
    jj = j + 2
    w = picScreen(jj).ScaleWidth
    h = picScreen(jj).ScaleHeight
        
        last_y = -1
        '===========================
        ' draw raw data
        '===========================
         For i = 0 To acq_count - 1
          
            X = (i * w) / (acq_count)
            samp = WMARawData(j, startidx + i) ' jj:mean picture index, i: data array
        
            Y = CInt((CSng(32767 - CLng(samp)) / CSng(65535)) * h)
    
              If i > 0 Then
                 picScreen(jj).Line (last_x, last_y)-(X, Y), &HFFFF&
                 
             End If
            last_x = X
            last_y = Y
            last_y1 = y1
        Next
    
          
        '===========================
        ' draw Audio Pattern
        '===========================
         For i = 0 To acq_count - 1
          
            X = (i * w) / (acq_count)
           ' samp1 = WMAPattern(WMAMatchIndex, j, startidx + i)
            samp1 = SineCurve2(j, i)
            y1 = CInt((CSng(32767 - CLng(samp1)) / CSng(65535)) * h)
         
             If i > 0 Then
                 picScreen(jj).Line (last_x, last_y1)-(X, y1), &HFF00&
                 
             End If
            last_x = X
          
            last_y1 = y1
        Next
        
    
    
    Next
  
  
 
End Sub

Private Sub Enlarge_Click()
DataCount = DataCount * 2
ShowData2 BeginPoint, DataCount
Text2.Text = DataCount
End Sub

Private Sub Form_Load()
ShowData 0, 100
BeginPoint = 0
DataCount = 6000
End Sub

Private Sub Form_Paint()
'ShowData 0, 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim result As Long
  Card2214Exist = 0
   If Card2214 >= 0 Then
      D2K_Release_Card (Card2214)
   End If
End Sub

Private Sub HScroll1_Change()
BeginPoint = HScroll1.value / 32767 * 6000
ShowData2 BeginPoint, DataCount
Text1.Text = BeginPoint
End Sub

Private Sub mnuExit_Click()
  If Card2214 >= 0 Then D2K_Release_Card (Card2214)
  End
End Sub


 Private Sub DrawTrigLevel(ByVal old_level As Integer, ByVal level As Integer, ByVal old_llevel As Integer, ByVal llevel As Integer)
    Dim w As Long, h As Long
    Dim Y As Long
    Dim j As Integer
    For j = 0 To 3
        picScreen(j).Cls
        w = picScreen(j).ScaleWidth
        h = picScreen(j).ScaleHeight
        If old_level <> -1 Then
            Y = h * (255 - old_level) / 255
            picScreen(j).Line (0, Y)-(w, Y), &H0
        End If
        If old_llevel <> -1 Then
            Y = h * (255 - old_llevel) / 255
            picScreen(j).Line (0, Y)-(w, Y), &H0
        End If
        Y = h * (255 - level) / 255
        picScreen(j).Line (0, Y)-(w, Y), &H909090
        Y = h * (255 - llevel) / 255
        picScreen(j).Line (0, Y)-(w, Y), &HBBBBBB
    Next
End Sub


Private Function Min(ByVal v1 As Long, ByVal v2 As Long) As Long
  If v1 < v2 Then
    Min = v1
  Else
    Min = v2
  End If
End Function

Private Sub Shrink_Click()
DataCount = DataCount / 2
ShowData2 BeginPoint, DataCount
Text2.Text = DataCount
End Sub
