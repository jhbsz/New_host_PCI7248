VERSION 5.00
Begin VB.Form Alarm_NonUPT2 
   Caption         =   "Please Change OS"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   13110
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   735
      Left            =   4080
      TabIndex        =   3
      Top             =   5640
      Width           =   4335
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   12000
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   3960
      Picture         =   "Alarm_NonUPT2.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   4515
      TabIndex        =   1
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Please Reset PC! Select UPT2 (DOS) environment!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   11175
   End
   Begin VB.Label Label1 
      Caption         =   "請重新開機! 並選擇UPT2 (DOS)環境!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   36
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   12255
   End
End
Attribute VB_Name = "Alarm_NonUPT2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TimerFlag As Boolean

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Activate()
    Call Timer1_Timer
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    If TimerFlag = True Then
        TimerFlag = False
        Label1.ForeColor = &H80000012
        Label2.ForeColor = &H80000012
        Exit Sub
    End If
    
    If TimerFlag = False Then
        TimerFlag = True
        Label1.ForeColor = &HFF&
        Label2.ForeColor = &HFF&
        Exit Sub
    End If
        
End Sub

