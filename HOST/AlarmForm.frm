VERSION 5.00
Begin VB.Form Alarm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   11790
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "關閉視窗(close)"
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確認 (Check)"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   10215
   End
End
Attribute VB_Name = "Alarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i

Private Sub Command1_Click()
 i = 1
End Sub

Private Sub Command2_Click()
If i = 0 Then
  Exit Sub
End If
i = 0
AlarmCtrl = 0
Unload Me
End Sub

Private Sub Form_Activate()
Dim OldTime
 
Do
    DoEvents
    If (Timer - OldTime) > 1 Then
        OldTime = Timer
        
        If BackColor = &HFF& Then
           BackColor = &HFF00&
        Else
            BackColor = &HFF&
        End If
    End If
    
    If i = 1 Then
    Exit Do
    End If
Loop While (1)
End Sub

