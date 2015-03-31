VERSION 5.00
Begin VB.Form DebugForm 
   Caption         =   "Debug"
   ClientHeight    =   3750
   ClientLeft      =   10725
   ClientTop       =   3630
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   4710
   Begin VB.TextBox SiteOnOff 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Height          =   285
      Index           =   3
      Left            =   2640
      TabIndex        =   14
      Text            =   "S4"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox SiteOnOff 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Height          =   285
      Index           =   2
      Left            =   2160
      TabIndex        =   13
      Text            =   "S3"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox SiteOnOff 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   12
      Text            =   "S2"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Log"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox textS4Bin 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      TabIndex        =   10
      Text            =   "PASS"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox textS3Bin 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Text            =   "PASS"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox textS2Bin 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      TabIndex        =   6
      Text            =   "PASS"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox textS1Bin 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "PASS"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox SiteOnOff 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Text            =   "S1"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox textState 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "Site4 Bin"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Site3 Bin"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Site2 Bin"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Site1 Bin"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Site On/Off"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "DebugForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    textState.Text = ""
End Sub
Private Sub SiteOnOff_Click(Index As Integer)
    If SiteOnOff(Index).BackColor = DebugSiteOn Then
        SiteOnOff(Index).BackColor = DebugSiteOff
    Else
        SiteOnOff(Index).BackColor = DebugSiteOn
    End If
End Sub
