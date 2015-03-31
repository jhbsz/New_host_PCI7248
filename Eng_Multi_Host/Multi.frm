VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Host 
   Caption         =   "Multi "
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12975
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   14.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   12975
   StartUpPosition =   3  '系統預設值
   Begin MSCommLib.MSComm MSComm3 
      Left            =   9720
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   9720
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Mode 
      Caption         =   "Mode"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   120
      Top             =   720
      Visible         =   0   'False
      Width           =   3615
      Begin VB.OptionButton Option3 
         Caption         =   "二號機"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   123
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "一號機"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   122
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "雙機"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   121
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame HandlerType 
      Caption         =   "HandlerType"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9600
      TabIndex        =   118
      Top             =   120
      Width           =   3255
      Begin VB.OptionButton Dummy 
         Caption         =   "Dummy"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   126
         Top             =   360
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton Handler_SRM 
         Caption         =   "SRM"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   125
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Handler_NS 
         Caption         =   "NS"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   124
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Handler_Blank 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   119
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   1320
      TabIndex        =   45
      Top             =   2040
      Width           =   735
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   1320
      TabIndex        =   44
      Top             =   2040
      Width           =   735
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   1320
      TabIndex        =   43
      Top             =   2040
      Width           =   735
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   1320
      TabIndex        =   42
      Top             =   2040
      Width           =   735
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1320
      TabIndex        =   41
      Top             =   2040
      Width           =   735
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1320
      TabIndex        =   40
      Top             =   2040
      Width           =   735
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   39
      Top             =   2040
      Width           =   735
   End
   Begin VB.Timer Blank_Timer 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   9120
      Top             =   1800
   End
   Begin VB.ComboBox ChipNameCombo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   29
      Text            =   "Ver"
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton EndBtn 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11760
      TabIndex        =   28
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CheckBox ReportCheck 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   27
      Top             =   480
      Width           =   1215
   End
   Begin VB.CheckBox OneCycleCheck 
      Caption         =   "OneCycle"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   26
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton StopBtn 
      Caption         =   "STOP"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   18
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox ChipNameCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   17
      Text            =   "ChipName"
      Top             =   120
      Width           =   2055
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton BeginBtn 
      Caption         =   "Begin"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   14
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CheckBox ResetPC 
      Caption         =   "Bin2 Fail >5, Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6600
      TabIndex        =   13
      Top             =   120
      Width           =   2655
   End
   Begin VB.CheckBox OffLineCheck 
      Caption         =   "Off Line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   12
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox SiteCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   0
      Left            =   10440
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   1
      Left            =   11160
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   2
      Left            =   11880
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   3
      Left            =   12600
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   4
      Left            =   10440
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   5
      Left            =   11160
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   6
      Left            =   11880
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   7
      Left            =   12600
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   117
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   1320
      TabIndex        =   116
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   1320
      TabIndex        =   115
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1320
      TabIndex        =   114
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1320
      TabIndex        =   113
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1320
      TabIndex        =   112
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label ContFailTitle 
      BackColor       =   &H80000013&
      Caption         =   "Binning"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   111
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   110
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   109
      Top             =   8400
      Width           =   855
   End
   Begin VB.Line Line2 
      X1              =   1080
      X2              =   1080
      Y1              =   1560
      Y2              =   8880
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   108
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   107
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   106
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   105
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   104
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   103
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   102
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   101
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   100
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   99
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   98
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   97
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   96
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   95
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   94
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   93
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   92
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   91
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   90
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   89
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   88
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   87
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   86
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   85
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   84
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   83
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   82
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   81
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   80
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   79
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   78
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   77
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   76
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   75
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   74
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   73
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   72
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   71
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   70
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   69
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   68
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   67
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   66
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   65
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   64
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   63
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   62
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   61
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   60
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   59
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   58
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   57
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   56
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   55
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   54
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   53
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   52
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   51
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   50
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   49
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   48
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   47
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   46
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Bin4Title 
      Caption         =   "Bin4"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   38
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Bin2Title 
      Caption         =   "Bin2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   37
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   36
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   35
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   34
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   33
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   32
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   31
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   30
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label TestCycleTimeLbl 
      Caption         =   "TestCycleTime"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   25
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label SiteLbl 
      Caption         =   "Sites"
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
      Left            =   2400
      TabIndex        =   24
      Top             =   120
      Width           =   855
   End
   Begin VB.Label YieldTitle 
      Caption         =   "Yield"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   23
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label TotalFailTitle 
      Caption         =   "Fail"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   22
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label TimeOutTitle 
      Caption         =   "TimeOut"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   21
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Bin5Title 
      Caption         =   "Bin5"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   20
      Top             =   5760
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   13320
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Bin3Title 
      Caption         =   "Bin3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   15
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Bin1Title 
      Caption         =   "Bin1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   11
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label avgTestTimeLbl 
      Caption         =   "AVG TestTime"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   855
   End
End
Attribute VB_Name = "Host"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const HandlerStartTimeOut = 5
Const PCReadyTimeOut = 5

Const PCResetTimeOut = 100
Const PC_RESET_TIME = 0.2

Dim AllenStop As Byte
Dim AllenIdleStateStop As Byte
Dim SelectFailFlag As Byte
Dim PrintEnable As Byte
Dim StopFlag As Byte
Dim PCIFlag As Byte
Dim MinGetStartTime
Dim TotalCycleTime

Dim TmpBuf As String
Dim AU6254Msg As String

Const PCI7248_EOT = &H1   'for 7248 card
Const PCI7248_PASS = &HFD 'for 7248 card  11111101
Const PCI7248_BIN2 = &HFB 'for 7248 card  11111011
Const PCI7248_BIN3 = &HF7 'for 7248 card  11110111
Const PCI7248_BIN4 = &HEF 'for 7248 card  11101111
Const PCI7248_BIN5 = &HDF 'for 7248 card  11011111

Const PCI7296_EOT = &H1             'for 7248 card
Const PCI7296_PASS = &HBD           'for 7248 card  10 111 101
Const PCI7296_BIN2 = &HBB           'for 7248 card  10 111 011
Const PCI7296_BIN3 = &HB7           'for 7248 card  10 110 111
Const PCI7296_BIN4 = &HAF           'for 7248 card  10 101 111
Const PCI7296_BIN5 = &H9F           'for 7248 card  10 011 111
Const PCI7296_CLEAR_START = &HBF    'for 7248 card  10 111 111
Const PCI7296_PC_RESET = &H7F       'for 7248 card  01 111 111

Const GREEN_COLOR = &HFF00&
Const RED_COLOR = &H8080FF
Const YELLOW_COLOR = &HFFFF&
Const BACK_COLOR = &H8000000F

Dim TestTimeOut As Single

Public ChipName As String

Dim StartFlag(0 To 7) As Byte
Dim GetStart(0 To 7) As Byte
Dim BinFlag(0 To 7) As Byte
Dim OneCycleFlag As Byte
Dim SiteCounter As Byte
Dim StartCounter As Byte
Dim BinCounter As Byte
Dim OffLineSiteCounter As Byte

Dim State(0 To 7) As String
Dim Fire(0 To 7)
Dim PassTime(0 To 7)
Dim OldState(0 To 7) As String
Dim Buf(0 To 7) As String  ' to get Comm buf
Dim TestResult(0 To 7) As String
Dim GetStartTime(0 To 7)
Dim RealTestTime(0 To 7)
Dim CycleTestTime(0 To 7)
Dim TimeOutCounter(0 To 7) As Long
Dim Bin2ContiFail(0 To 7) As Long
Dim ContiFailCounter(0 To 7) As Integer

Const IdleState = "IDLE"
Const HandlerStartState = "START"
Const PCReadyState = "READY"
Const PCResetState = "RESET"
Const PCResetFailState = "RESET_FAIL"
Const BinState = "BINING"
Const PCIState = "PCIBin"
 
Dim CPort(0 To 7) As Byte

Dim FirstTimeFlag As Byte
 
Dim Channel As Byte
Dim i As Byte
Dim result

Public First_site As Boolean
Public Second_site As Boolean

Public TestResult1 As String
Public TestResult2 As String

Dim value_a(0 To 1) As Long, value_b(0 To 1) As Long, value_cu(0 To 1) As Long, value_cl(0 To 1) As Long

Public WaitForStart
Public TestMode As Byte
Public WaitStartTime

Public TotalRealTestTime
Public OldTotalRealTestTime

Public DI_P As Long
Public DO_P As Long

Public buf1
Public buf2

Public WAIT_START_TIME_OUT  As Single   ' wait start signal time out condition
Public WAIT_TEST_CYCLE_OUT As Single    ' wait test time out cycle time
Public POWER_ON_TIME As Single
Public UNLOAD_DRIVER As Single
Public CAPACTOR_CHARGE As Single
Public NO_CARD_TEST_TIME As Single
Public Need_GPIB As Byte

Public NewPowerOnTime As Single

Public TesterStatus1
Public TesterStatus2

Public TesterReady1 As Byte
Public TesterReady2 As Byte
                                   
Public ResetCounter1 As Byte
Public ResetCounter2 As Byte

Public TesterDownCount1 As Byte
Public TesterDownCount2 As Byte

Dim TesterDownCountTimer1   ' timer
Dim TesterDownCountTimer2

Public FirstRun As Byte
Public WaitForReady         ' timer

Public WaitForPowerOn1      ' timer
Public WaitForPowerOn2

Public VB6_Flag As Boolean

Dim GPIBInquiryTime
Dim OldRealTestTime

Public WaitStartCounter As Integer
Public WaitStartTimeOutCounter As Integer
Public WaitStartTimeOut As Integer

Dim WaitTestTimeOut1 As Integer
Dim WaitTestTimeOut2 As Integer

Dim WaitTestTimeOutCounter1 As Integer
Dim WaitTestTimeOutCounter2 As Integer

Public continuefail1 As Integer
'Public continuefail1_bin2 As Integer
'Public continuefail1_bin3 As Integer
'Public continuefail1_bin4 As Integer
'Public continuefail1_bin5 As Integer
Public continuefail2 As Integer
'Public continuefail2_bin2 As Integer
'Public continuefail2_bin3 As Integer
'Public continuefail2_bin4 As Integer
'Public continuefail2_bin5 As Integer

Dim GetGPIBStatus(1) As Boolean
Dim GPIBReady(1) As Boolean

Public NoCardTestResult1 As String
Public NoCardTestResult2 As String

Public NoCardTestCycleTime1
Public NoCardTestCycleTime2

Public NoCardTestStop1 As Byte
Public NoCardTestStop2 As Byte

Public NoCardWaitForTest1
Public NoCardWaitForTest2

Public WaitForTest1
Public WaitForTest2

Public TestStop1 As Byte
Public TestStop2 As Byte

Dim TestCycleTime1
Dim TestCycleTime2
Dim TestCounter As Integer

Const AlarmLimit = 5

Public GreaTekChipName As String

Public PassCounter1 As Long
Public PassCounter2 As Long

Public OffLPassCounter1 As Long
Public OffLPassCounter2 As Long

' ================= for prevent multi open host fail ============
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, _
  ByVal lpWindowName As String) As Long
  
Private Declare Function ShowWindow Lib "user32" _
  (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
  
Private Declare Function SetForegroundWindow Lib "user32" _
  (ByVal hwnd As Long) As Long
' ================================================================

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

Sub PrintReportSummary2()
On Error Resume Next
Dim i As Byte

Dim TestedSite(0 To 7) As Long
Dim TestedTotal As Long
Dim TestedPercent As Single
Dim PassTotal As Long
Dim PassPercent As Single
Dim FailTotal As Long
Dim FailSite(0 To 7) As Long
Dim FailPercent As Single
Dim Bin2Total As Long
Dim Bin2Percent As Single
Dim Bin3Total As Long
Dim Bin3Percent As Single
Dim Bin4Total As Long
Dim Bin4Percent As Single
Dim Bin5Total As Long
Dim Bin5Percent As Single

Dim TestedSite1 As Long     ' for dual site only
Dim TestedSite2 As Long     ' for dual site only
 
Dim PassSite1 As Long       ' for dual site only
Dim PassSite2 As Long       ' for dual site only

Dim FailSite1 As Long       ' for dual site only
Dim FailSite2 As Long       ' for dual site only

    If ReportCheck.value = 0 Then
        Exit Sub
    End If
    
    If DataBaseDebug = 1 Then
        For i = 0 To 7
            Bin1Counter(i) = 2903
            Bin2Counter(i) = 10
            Bin3Counter(i) = 6
            Bin4Counter(i) = 1
            Bin5Counter(i) = 1
        Next
    End If
    
    OutFileName = RunCardNO & "_" & ProcessIDSum & "_" & Left(EndDay, 4) & Mid(EndDay, 6, 2) & Right(EndDay, 2)
    OutFileName = OutFileName & Left(EndSecond, 2) & Mid(EndSecond, 4, 2) & Right(EndSecond, 2) & "Sum.txt"
  
    ' calculate summary
    Call GetReportSummarySub2
    
    If No8PCard Then
        TestedSite1 = Bin1Site1Sum + Bin2Site1Sum + Bin3Site1Sum + Bin4Site1Sum + Bin5Site1Sum
        TestedSite2 = Bin1Site2Sum + Bin2Site2Sum + Bin3Site2Sum + Bin4Site2Sum + Bin5Site2Sum
        TestedTotal = TestedSite1 + TestedSite2
        
        TestedPercent = 1
        
        PassSite1 = Bin1Site1Sum
        PassSite2 = Bin1Site2Sum
        PassTotal = PassSite1 + PassSite2
        PassPercent = CSng(PassTotal / TestedTotal)
        
        FailSite1 = Bin2Site1Sum + Bin3Site1Sum + Bin4Site1Sum + Bin5Site1Sum
        FailSite2 = Bin2Site2Sum + Bin3Site2Sum + Bin4Site2Sum + Bin5Site2Sum
        FailTotal = FailSite1 + FailSite2
        FailPercent = CSng(FailTotal / TestedTotal)
        
        Bin2Total = Bin2Site1Sum + Bin2Site2Sum
        Bin2Percent = CSng(Bin2Total / TestedTotal)
            
        Bin3Total = Bin3Site1Sum + Bin3Site2Sum
        Bin3Percent = CSng(Bin3Total / TestedTotal)
                
        Bin4Total = Bin4Site1Sum + Bin4Site2Sum
        Bin4Percent = CSng(Bin4Total / TestedTotal)
                    
                
        Bin5Total = Bin5Site1Sum + Bin5Site2Sum
        Bin5Percent = CSng(Bin5Total / TestedTotal)
        
    Else
    
        For i = 0 To 7
            TestedSite(i) = Bin1Sum(i) + Bin2Sum(i) + Bin3Sum(i) + Bin4Sum(i) + Bin5Sum(i)
        Next i
         
        For i = 0 To 7
            TestedTotal = TestedTotal + TestedSite(i)
        Next
        TestedPercent = 1
        
        For i = 0 To 7
            PassTotal = PassTotal + Bin1Sum(i)
        Next i
        PassPercent = CSng(PassTotal / TestedTotal)
        
        For i = 0 To 7
            FailSite(i) = Bin2Sum(i) + Bin3Sum(i) + Bin4Sum(i) + Bin5Sum(i)
        Next i
        
        For i = 0 To 7
            FailTotal = FailTotal + FailSite(i)
        Next
        FailPercent = CSng(FailTotal / TestedTotal)
        
        For i = 0 To 7
            Bin2Total = Bin2Total + Bin2Sum(i)
        Next
        Bin2Percent = CSng(Bin2Total / TestedTotal)
            
        For i = 0 To 7
            Bin3Total = Bin3Total + Bin3Sum(i)
        Next i
        Bin3Percent = CSng(Bin3Total / TestedTotal)
        
        For i = 0 To 7
            Bin4Total = Bin4Total + Bin4Sum(i)
        Next i
        Bin4Percent = CSng(Bin4Total / TestedTotal)
                
        For i = 0 To 7
            Bin5Total = Bin5Total + Bin5Sum(i)
        Next i
        Bin5Percent = CSng(Bin5Total / TestedTotal)
    
    End If
    
    Open "D:\SLT Summary\" & OutFileName For Output As #1
    
    Print #1, "#####################################################"
    Print #1, "Name of PC: " & NameofPC
    Print #1, "Program Name: " & ProgramName
    Print #1, "Program Version Code: " & ProgramRevisionCode
    Print #1, "Device ID: " & DeviceID
    Print #1, "Run Card NO: " & RunCardNO
    Print #1, "Lot ID: " & LotID
    Print #1, "Process: " & ProcessIDSum
    Print #1, "Start at: " & StartAtMin
    Print #1, "End at: " & EndAtMax
    Print #1, "HandlerID: " & HandlerID
    Print #1, "Operator Name: " & OperatorName

    If No8PCard Then
        Print #1, "-------------------------------------------------------"
        Print #1, Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Total  " & Space(3) & "Total"
        Print #1, Space(13) & "COUNT  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percent"
        Print #1, "-------------------------------------------------------"
    Else
    
        Print #1, "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Print #1, Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Site 3 " & Space(3) & "Site 4 " & Space(3) & "Site 5 " & Space(3) & "Site 6 " & Space(3) & "Site 7 " & Space(3) & "Site 8 " & Space(3) & "Total  " & Space(3) & "Total"
        Print #1, Space(13) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percent"
        Print #1, "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    End If

    If No8PCard Then
        ' file output
        
        Print #1, "TESTED" & Space(7) & Space(7 - Len(Format(TestedSite1, "#######"))) & Format(TestedSite1, "#######") _
                                & Space(3) & Space(7 - Len(Format(TestedSite2, "#######"))) & Format(TestedSite2, "#######") _
                                & Space(3) & Space(7 - Len(Format(TestedTotal, "#######"))) & Format(TestedTotal, "#######") _
                                & Space(3) & Space(7 - Len(Format(TestedPercent, "0.00%"))) & Format(TestedPercent, "0.00%")
        
        Print #1, "PASS" & Space(9) & Space(7 - Len(Format(PassSite1, "#######"))) & Format(PassSite1, "#######") _
                               & Space(3) & Space(7 - Len(Format(PassSite2, "#######"))) & Format(PassSite2, "#######") _
                               & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                               & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
        
        Print #1, "FAIL" & Space(9) & Space(7 - Len(Format(FailSite1, "#######"))) & Format(FailSite1, "#######") _
                               & Space(3) & Space(7 - Len(Format(FailSite2, "#######"))) & Format(FailSite2, "#######") _
                               & Space(3) & Space(7 - Len(Format(FailTotal, "#######"))) & Format(FailTotal, "#######") _
                               & Space(3) & Space(7 - Len(Format(FailPercent, "0.00%"))) & Format(FailPercent, "0.00%")
        
        ' file output
        Print #1, "1 PASS" & Space(7) & Space(7 - Len(Format(PassSite1, "#######"))) & Format(PassSite1, "#######") _
                               & Space(3) & Space(7 - Len(Format(PassSite2, "#######"))) & Format(PassSite2, "#######") _
                               & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                               & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
        
        Print #1, "2 BIN2" & Space(7) & Space(7 - Len(Format(Bin2Site1Sum, "#######"))) & Format(Bin2Site1Sum, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin2Site2Sum, "#######"))) & Format(Bin2Site2Sum, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin2Total, "#######"))) & Format(Bin2Total, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin2Percent, "0.00%"))) & Format(Bin2Percent, "0.00%")
        
        Print #1, "3 BIN3" & Space(7) & Space(7 - Len(Format(Bin3Site1Sum, "#######"))) & Format(Bin3Site1Sum, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin3Site2Sum, "#######"))) & Format(Bin3Site2Sum, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin3Total, "#######"))) & Format(Bin3Total, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin3Percent, "0.00%"))) & Format(Bin3Percent, "0.00%")
        
        Print #1, "4 BIN4" & Space(7) & Space(7 - Len(Format(Bin4Site1Sum, "#######"))) & Format(Bin4Site1Sum, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin4Site2Sum, "#######"))) & Format(Bin4Site2Sum, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin4Total, "#######"))) & Format(Bin4Total, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin4Percent, "0.00%"))) & Format(Bin4Percent, "0.00%")
        
        Print #1, "5 BIN5" & Space(7) & Space(7 - Len(Format(Bin5Site1Sum, "#######"))) & Format(Bin5Site1Sum, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin5Site2Sum, "#######"))) & Format(Bin5Site2Sum, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin5Total, "#######"))) & Format(Bin5Total, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin5Percent, "0.00%"))) & Format(Bin5Percent, "0.00%")
            
    Else

        '================ file output
                                       
        Print #1, "TESTED" & Space(7) & Space(7 - Len(Format(TestedSite(0), "#######"))) & Format(TestedSite(0), "#######") _
                           & Space(3) & Space(7 - Len(Format(TestedSite(1), "#######"))) & Format(TestedSite(1), "#######") _
                           & Space(3) & Space(7 - Len(Format(TestedSite(2), "#######"))) & Format(TestedSite(2), "#######") _
                           & Space(3) & Space(7 - Len(Format(TestedSite(3), "#######"))) & Format(TestedSite(3), "#######") _
                           & Space(3) & Space(7 - Len(Format(TestedSite(4), "#######"))) & Format(TestedSite(4), "#######") _
                           & Space(3) & Space(7 - Len(Format(TestedSite(5), "#######"))) & Format(TestedSite(5), "#######") _
                           & Space(3) & Space(7 - Len(Format(TestedSite(6), "#######"))) & Format(TestedSite(6), "#######") _
                           & Space(3) & Space(7 - Len(Format(TestedSite(7), "#######"))) & Format(TestedSite(7), "#######") _
                           & Space(3) & Space(7 - Len(Format(TestedTotal, "#######"))) & Format(TestedTotal, "#######") _
                           & Space(3) & Space(7 - Len(Format(TestedPercent, "0.00%"))) & Format(TestedPercent, "0.00%")
        
        Print #1, "PASS" & Space(9) & Space(7 - Len(Format(Bin1Sum(0), "#######"))) & Format(Bin1Sum(0), "#######") _
                         & Space(3) & Space(7 - Len(Format(Bin1Sum(1), "#######"))) & Format(Bin1Sum(1), "#######") _
                         & Space(3) & Space(7 - Len(Format(Bin1Sum(2), "#######"))) & Format(Bin1Sum(2), "#######") _
                         & Space(3) & Space(7 - Len(Format(Bin1Sum(3), "#######"))) & Format(Bin1Sum(3), "#######") _
                         & Space(3) & Space(7 - Len(Format(Bin1Sum(4), "#######"))) & Format(Bin1Sum(4), "#######") _
                         & Space(3) & Space(7 - Len(Format(Bin1Sum(5), "#######"))) & Format(Bin1Sum(5), "#######") _
                         & Space(3) & Space(7 - Len(Format(Bin1Sum(6), "#######"))) & Format(Bin1Sum(6), "#######") _
                         & Space(3) & Space(7 - Len(Format(Bin1Sum(7), "#######"))) & Format(Bin1Sum(7), "#######") _
                         & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                         & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
        
        Print #1, "FAIL" & Space(9) & Space(7 - Len(Format(FailSite(0), "#######"))) & Format(FailSite(0), "#######") _
                         & Space(3) & Space(7 - Len(Format(FailSite(1), "#######"))) & Format(FailSite(1), "#######") _
                         & Space(3) & Space(7 - Len(Format(FailSite(2), "#######"))) & Format(FailSite(2), "#######") _
                         & Space(3) & Space(7 - Len(Format(FailSite(3), "#######"))) & Format(FailSite(3), "#######") _
                         & Space(3) & Space(7 - Len(Format(FailSite(4), "#######"))) & Format(FailSite(4), "#######") _
                         & Space(3) & Space(7 - Len(Format(FailSite(5), "#######"))) & Format(FailSite(5), "#######") _
                         & Space(3) & Space(7 - Len(Format(FailSite(6), "#######"))) & Format(FailSite(6), "#######") _
                         & Space(3) & Space(7 - Len(Format(FailSite(7), "#######"))) & Format(FailSite(7), "#######") _
                         & Space(3) & Space(7 - Len(Format(FailTotal, "#######"))) & Format(FailTotal, "#######") _
                         & Space(3) & Space(7 - Len(Format(FailPercent, "0.00%"))) & Format(FailPercent, "0.00%")
        
        '=============== file output
        Print #1, "1 PASS" & Space(7) & Space(7 - Len(Format(Bin1Sum(0), "#######"))) & Format(Bin1Sum(0), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin1Sum(1), "#######"))) & Format(Bin1Sum(1), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin1Sum(2), "#######"))) & Format(Bin1Sum(2), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin1Sum(3), "#######"))) & Format(Bin1Sum(3), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin1Sum(4), "#######"))) & Format(Bin1Sum(4), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin1Sum(5), "#######"))) & Format(Bin1Sum(5), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin1Sum(6), "#######"))) & Format(Bin1Sum(6), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin1Sum(7), "#######"))) & Format(Bin1Sum(7), "#######") _
                           & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                           & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
        
        
        Print #1, "2 BIN2" & Space(7) & Space(7 - Len(Format(Bin2Sum(0), "#######"))) & Format(Bin2Sum(0), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Sum(1), "#######"))) & Format(Bin2Sum(1), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Sum(2), "#######"))) & Format(Bin2Sum(2), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Sum(3), "#######"))) & Format(Bin2Sum(3), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Sum(4), "#######"))) & Format(Bin2Sum(4), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Sum(5), "#######"))) & Format(Bin2Sum(5), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Sum(6), "#######"))) & Format(Bin2Sum(6), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Sum(7), "#######"))) & Format(Bin2Sum(7), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Total, "#######"))) & Format(Bin2Total, "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Percent, "0.00%"))) & Format(Bin2Percent, "0.00%")
        
        
        Print #1, "3 BIN3" & Space(7) & Space(7 - Len(Format(Bin3Sum(0), "#######"))) & Format(Bin3Sum(0), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Sum(1), "#######"))) & Format(Bin3Sum(1), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Sum(2), "#######"))) & Format(Bin3Sum(2), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Sum(3), "#######"))) & Format(Bin3Sum(3), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Sum(4), "#######"))) & Format(Bin3Sum(4), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Sum(5), "#######"))) & Format(Bin3Sum(5), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Sum(6), "#######"))) & Format(Bin3Sum(6), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Sum(7), "#######"))) & Format(Bin3Sum(7), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Total, "#######"))) & Format(Bin3Total, "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Percent, "0.00%"))) & Format(Bin3Percent, "0.00%")
        
        
        Print #1, "4 BIN4" & Space(7) & Space(7 - Len(Format(Bin4Sum(0), "#######"))) & Format(Bin4Sum(0), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Sum(1), "#######"))) & Format(Bin4Sum(1), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Sum(2), "#######"))) & Format(Bin4Sum(2), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Sum(3), "#######"))) & Format(Bin4Sum(3), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Sum(4), "#######"))) & Format(Bin4Sum(4), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Sum(5), "#######"))) & Format(Bin4Sum(5), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Sum(6), "#######"))) & Format(Bin4Sum(6), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Sum(7), "#######"))) & Format(Bin4Sum(7), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Total, "#######"))) & Format(Bin4Total, "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Percent, "0.00%"))) & Format(Bin4Percent, "0.00%")
        
        
        Print #1, "5 BIN5" & Space(7) & Space(7 - Len(Format(Bin5Sum(0), "#######"))) & Format(Bin5Sum(0), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Sum(1), "#######"))) & Format(Bin5Sum(1), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Sum(2), "#######"))) & Format(Bin5Sum(2), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Sum(3), "#######"))) & Format(Bin5Sum(3), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Sum(4), "#######"))) & Format(Bin5Sum(4), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Sum(5), "#######"))) & Format(Bin5Sum(5), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Sum(6), "#######"))) & Format(Bin5Sum(6), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Sum(7), "#######"))) & Format(Bin5Sum(7), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Total, "#######"))) & Format(Bin5Total, "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Percent, "0.00%"))) & Format(Bin5Percent, "0.00%")
    End If
    
    Close #1

    '============================ printer section ===========================
    'Printer.CurrentX = 300

    If ReportDebug = 1 Then
        Exit Sub
    End If
    
    If (MsgBox("是否列印報表? ", vbYesNo + vbQuestion + vbDefaultButton2, "Comform Stop") = vbYes) Then

        Printer.Orientation = 2     ' 1 for 直印 2 for 橫印
        Printer.FontSize = 14
        Printer.Font = "標楷體"
        Printer.Print "#####################################################"
        Printer.Print "Name of PC: " & NameofPC
        Printer.Print "Program Name: " & ProgramName
        Printer.Print "Program Version Code: " & ProgramRevisionCode
        Printer.Print "Device ID: " & DeviceID
        Printer.Print "Run Card NO: " & RunCardNO
        Printer.Print "Lot ID: " & LotID
        Printer.Print "Process: " & ProcessIDSum
        Printer.Print "Start at: " & StartAtMin
        Printer.Print "End at: " & EndAtMax
        Printer.Print "HandlerID: " & HandlerID
        Printer.Print "Operator Name: " & OperatorName
        Printer.Print
        
        If No8PCard Or SiteCombo.Text = "2" Then
            Printer.Print "-------------------------------------------------------"
            Printer.Print Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Total  " & Space(3) & "Total"
            Printer.Print Space(13) & "COUNT  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percent"
            Printer.Print "-------------------------------------------------------"
        ElseIf SiteCombo.Text = "4" Then
            Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Site 3 " & Space(3) & "Site 4 " & Space(3) & "Total  " & Space(3) & "Total"
            Printer.Print Space(13) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percent"
            Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------"
        ElseIf SiteCombo.Text = "6" Then
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Site 3 " & Space(3) & "Site 4 " & Space(3) & "Site 5 " & Space(3) & "Site 6 " & Space(3) & "Total  " & Space(3) & "Total"
            Printer.Print Space(13) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percent"
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Else
            Printer.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Site 3 " & Space(3) & "Site 4 " & Space(3) & "Site 5 " & Space(3) & "Site 6 " & Space(3) & "Site 7 " & Space(3) & "Site 8 " & Space(3) & "Total  " & Space(3) & "Total"
            Printer.Print Space(13) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percent"
            Printer.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        End If
    
        If No8PCard Then
            Printer.Print "TESTED" & Space(7) & Space(7 - Len(Format(TestedSite1, "#######"))) & Format(TestedSite1, "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite2, "#######"))) & Format(TestedSite2, "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedTotal, "#######"))) & Format(TestedTotal, "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedPercent, "0.00%"))) & Format(TestedPercent, "0.00%")
            
            Printer.Print "PASS" & Space(9) & Space(7 - Len(Format(PassSite1, "#######"))) & Format(PassSite1, "#######") _
                                 & Space(3) & Space(7 - Len(Format(PassSite2, "#######"))) & Format(PassSite2, "#######") _
                                 & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                                 & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
            
            
            Printer.Print "FAIL" & Space(9) & Space(7 - Len(Format(FailSite1, "#######"))) & Format(FailSite1, "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite2, "#######"))) & Format(FailSite2, "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailTotal, "#######"))) & Format(FailTotal, "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailPercent, "0.00%"))) & Format(FailPercent, "0.00%")
            
            Printer.Print "1 PASS" & Space(7) & Space(7 - Len(Format(PassSite1, "#######"))) & Format(PassSite1, "#######") _
                                   & Space(3) & Space(7 - Len(Format(PassSite2, "#######"))) & Format(PassSite2, "#######") _
                                   & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                                   & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
            
            Printer.Print "2 BIN2" & Space(7) & Space(7 - Len(Format(Bin2Site1Sum, "#######"))) & Format(Bin2Site1Sum, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Site2Sum, "#######"))) & Format(Bin2Site2Sum, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Total, "#######"))) & Format(Bin2Total, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Percent, "0.00%"))) & Format(Bin2Percent, "0.00%")
            
            Printer.Print "3 BIN3" & Space(7) & Space(7 - Len(Format(Bin3Site1Sum, "#######"))) & Format(Bin3Site1Sum, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Site2Sum, "#######"))) & Format(Bin3Site2Sum, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Total, "#######"))) & Format(Bin3Total, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Percent, "0.00%"))) & Format(Bin3Percent, "0.00%")
            
            Printer.Print "4 BIN4" & Space(7) & Space(7 - Len(Format(Bin4Site1Sum, "#######"))) & Format(Bin4Site1Sum, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Site2Sum, "#######"))) & Format(Bin4Site2Sum, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Total, "#######"))) & Format(Bin4Total, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Percent, "0.00%"))) & Format(Bin4Percent, "0.00%")
            
            Printer.Print "5 BIN5" & Space(7) & Space(7 - Len(Format(Bin5Site1Sum, "#######"))) & Format(Bin5Site1Sum, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Site2Sum, "#######"))) & Format(Bin5Site2Sum, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Total, "#######"))) & Format(Bin5Total, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Percent, "0.00%"))) & Format(Bin5Percent, "0.00%")
                               
            Printer.EndDoc
    
        ElseIf SiteCombo.Text = "2" Or SiteCombo.Text = "4" Then
    
            Printer.Print "TESTED" & Space(7) & Space(7 - Len(Format(TestedSite(0), "#######"))) & Format(TestedSite(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite(1), "#######"))) & Format(TestedSite(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite(2), "#######"))) & Format(TestedSite(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite(3), "#######"))) & Format(TestedSite(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedTotal, "#######"))) & Format(TestedTotal, "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedPercent, "0.00%"))) & Format(TestedPercent, "0.00%")
            
            Printer.Print "PASS" & Space(9) & Space(7 - Len(Format(Bin1Sum(0), "#######"))) & Format(Bin1Sum(0), "#######") _
                                 & Space(3) & Space(7 - Len(Format(Bin1Sum(1), "#######"))) & Format(Bin1Sum(1), "#######") _
                                 & Space(3) & Space(7 - Len(Format(Bin1Sum(2), "#######"))) & Format(Bin1Sum(2), "#######") _
                                 & Space(3) & Space(7 - Len(Format(Bin1Sum(3), "#######"))) & Format(Bin1Sum(3), "#######") _
                                 & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                                 & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
            
            Printer.Print "FAIL" & Space(9) & Space(7 - Len(Format(FailSite(0), "#######"))) & Format(FailSite(0), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite(1), "#######"))) & Format(FailSite(1), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite(2), "#######"))) & Format(FailSite(2), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite(3), "#######"))) & Format(FailSite(3), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailTotal, "#######"))) & Format(FailTotal, "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailPercent, "0.00%"))) & Format(FailPercent, "0.00%")
            
            ' file output
            Printer.Print "1 PASS" & Space(7) & Space(7 - Len(Format(Bin1Sum(0), "#######"))) & Format(Bin1Sum(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin1Sum(1), "#######"))) & Format(Bin1Sum(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin1Sum(2), "#######"))) & Format(Bin1Sum(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin1Sum(3), "#######"))) & Format(Bin1Sum(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                                   & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
            
            Printer.Print "2 BIN2" & Space(7) & Space(7 - Len(Format(Bin2Sum(0), "#######"))) & Format(Bin2Sum(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Sum(1), "#######"))) & Format(Bin2Sum(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Sum(2), "#######"))) & Format(Bin2Sum(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Sum(3), "#######"))) & Format(Bin2Sum(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Total, "#######"))) & Format(Bin2Total, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Percent, "0.00%"))) & Format(Bin2Percent, "0.00%")
            
            Printer.Print "3 BIN3" & Space(7) & Space(7 - Len(Format(Bin3Sum(0), "#######"))) & Format(Bin3Sum(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Sum(1), "#######"))) & Format(Bin3Sum(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Sum(2), "#######"))) & Format(Bin3Sum(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Sum(3), "#######"))) & Format(Bin3Sum(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Total, "#######"))) & Format(Bin3Total, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Percent, "0.00%"))) & Format(Bin3Percent, "0.00%")
            
            Printer.Print "4 BIN4" & Space(7) & Space(7 - Len(Format(Bin4Sum(0), "#######"))) & Format(Bin4Sum(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Sum(1), "#######"))) & Format(Bin4Sum(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Sum(2), "#######"))) & Format(Bin4Sum(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Sum(3), "#######"))) & Format(Bin4Sum(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Total, "#######"))) & Format(Bin4Total, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Percent, "0.00%"))) & Format(Bin4Percent, "0.00%")
            
            Printer.Print "5 BIN5" & Space(7) & Space(7 - Len(Format(Bin5Sum(0), "#######"))) & Format(Bin5Sum(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Sum(1), "#######"))) & Format(Bin5Sum(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Sum(2), "#######"))) & Format(Bin5Sum(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Sum(3), "#######"))) & Format(Bin5Sum(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Total, "#######"))) & Format(Bin5Total, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Percent, "0.00%"))) & Format(Bin5Percent, "0.00%")
            Printer.EndDoc
            
        ElseIf SiteCombo.Text = "6" Then
    
            Printer.Print "TESTED" & Space(7) & Space(7 - Len(Format(TestedSite(0), "#######"))) & Format(TestedSite(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite(1), "#######"))) & Format(TestedSite(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite(2), "#######"))) & Format(TestedSite(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite(3), "#######"))) & Format(TestedSite(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite(4), "#######"))) & Format(TestedSite(4), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite(5), "#######"))) & Format(TestedSite(5), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedTotal, "#######"))) & Format(TestedTotal, "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedPercent, "0.00%"))) & Format(TestedPercent, "0.00%")
            
            Printer.Print "PASS" & Space(9) & Space(7 - Len(Format(Bin1Sum(0), "#######"))) & Format(Bin1Sum(0), "#######") _
                                 & Space(3) & Space(7 - Len(Format(Bin1Sum(1), "#######"))) & Format(Bin1Sum(1), "#######") _
                                 & Space(3) & Space(7 - Len(Format(Bin1Sum(2), "#######"))) & Format(Bin1Sum(2), "#######") _
                                 & Space(3) & Space(7 - Len(Format(Bin1Sum(3), "#######"))) & Format(Bin1Sum(3), "#######") _
                                 & Space(3) & Space(7 - Len(Format(Bin1Sum(4), "#######"))) & Format(Bin1Sum(4), "#######") _
                                 & Space(3) & Space(7 - Len(Format(Bin1Sum(5), "#######"))) & Format(Bin1Sum(5), "#######") _
                                 & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                                 & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
            
            Printer.Print "FAIL" & Space(9) & Space(7 - Len(Format(FailSite(0), "#######"))) & Format(FailSite(0), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite(1), "#######"))) & Format(FailSite(1), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite(2), "#######"))) & Format(FailSite(2), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite(3), "#######"))) & Format(FailSite(3), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite(4), "#######"))) & Format(FailSite(4), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite(5), "#######"))) & Format(FailSite(5), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailTotal, "#######"))) & Format(FailTotal, "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailPercent, "0.00%"))) & Format(FailPercent, "0.00%")
            
            ' file output
            Printer.Print "1 PASS" & Space(7) & Space(7 - Len(Format(Bin1Sum(0), "#######"))) & Format(Bin1Sum(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin1Sum(1), "#######"))) & Format(Bin1Sum(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin1Sum(2), "#######"))) & Format(Bin1Sum(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin1Sum(3), "#######"))) & Format(Bin1Sum(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin1Sum(4), "#######"))) & Format(Bin1Sum(4), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin1Sum(5), "#######"))) & Format(Bin1Sum(5), "#######") _
                                   & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                                   & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
            
            Printer.Print "2 BIN2" & Space(7) & Space(7 - Len(Format(Bin2Sum(0), "#######"))) & Format(Bin2Sum(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Sum(1), "#######"))) & Format(Bin2Sum(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Sum(2), "#######"))) & Format(Bin2Sum(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Sum(3), "#######"))) & Format(Bin2Sum(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Sum(4), "#######"))) & Format(Bin2Sum(4), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Sum(5), "#######"))) & Format(Bin2Sum(5), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Percent, "0.00%"))) & Format(Bin2Percent, "0.00%")
            
            Printer.Print "3 BIN3" & Space(7) & Space(7 - Len(Format(Bin3Sum(0), "#######"))) & Format(Bin3Sum(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Sum(1), "#######"))) & Format(Bin3Sum(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Sum(2), "#######"))) & Format(Bin3Sum(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Sum(3), "#######"))) & Format(Bin3Sum(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Sum(4), "#######"))) & Format(Bin3Sum(4), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Sum(5), "#######"))) & Format(Bin3Sum(5), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Total, "#######"))) & Format(Bin3Total, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Percent, "0.00%"))) & Format(Bin3Percent, "0.00%")
            
            Printer.Print "4 BIN4" & Space(7) & Space(7 - Len(Format(Bin4Sum(0), "#######"))) & Format(Bin4Sum(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Sum(1), "#######"))) & Format(Bin4Sum(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Sum(2), "#######"))) & Format(Bin4Sum(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Sum(3), "#######"))) & Format(Bin4Sum(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Sum(4), "#######"))) & Format(Bin4Sum(4), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Sum(5), "#######"))) & Format(Bin4Sum(5), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Total, "#######"))) & Format(Bin4Total, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Percent, "0.00%"))) & Format(Bin4Percent, "0.00%")
            
            Printer.Print "5 BIN5" & Space(7) & Space(7 - Len(Format(Bin5Sum(0), "#######"))) & Format(Bin5Sum(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Sum(1), "#######"))) & Format(Bin5Sum(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Sum(2), "#######"))) & Format(Bin5Sum(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Sum(3), "#######"))) & Format(Bin5Sum(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Sum(4), "#######"))) & Format(Bin5Sum(4), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Sum(5), "#######"))) & Format(Bin5Sum(5), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Total, "#######"))) & Format(Bin5Total, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Percent, "0.00%"))) & Format(Bin5Percent, "0.00%")
            Printer.EndDoc
            
        Else
    
            Printer.Print "TESTED" & Space(7) & Space(7 - Len(Format(TestedSite(0), "#######"))) & Format(TestedSite(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite(1), "#######"))) & Format(TestedSite(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite(2), "#######"))) & Format(TestedSite(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite(3), "#######"))) & Format(TestedSite(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite(4), "#######"))) & Format(TestedSite(4), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite(5), "#######"))) & Format(TestedSite(5), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite(6), "#######"))) & Format(TestedSite(6), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedSite(7), "#######"))) & Format(TestedSite(7), "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedTotal, "#######"))) & Format(TestedTotal, "#######") _
                                   & Space(3) & Space(7 - Len(Format(TestedPercent, "0.00%"))) & Format(TestedPercent, "0.00%")
            
            Printer.Print "PASS" & Space(9) & Space(7 - Len(Format(Bin1Sum(0), "#######"))) & Format(Bin1Sum(0), "#######") _
                                 & Space(3) & Space(7 - Len(Format(Bin1Sum(1), "#######"))) & Format(Bin1Sum(1), "#######") _
                                 & Space(3) & Space(7 - Len(Format(Bin1Sum(2), "#######"))) & Format(Bin1Sum(2), "#######") _
                                 & Space(3) & Space(7 - Len(Format(Bin1Sum(3), "#######"))) & Format(Bin1Sum(3), "#######") _
                                 & Space(3) & Space(7 - Len(Format(Bin1Sum(4), "#######"))) & Format(Bin1Sum(4), "#######") _
                                 & Space(3) & Space(7 - Len(Format(Bin1Sum(5), "#######"))) & Format(Bin1Sum(5), "#######") _
                                 & Space(3) & Space(7 - Len(Format(Bin1Sum(6), "#######"))) & Format(Bin1Sum(6), "#######") _
                                 & Space(3) & Space(7 - Len(Format(Bin1Sum(7), "#######"))) & Format(Bin1Sum(7), "#######") _
                                 & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                                 & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
            
            Printer.Print "FAIL" & Space(9) & Space(7 - Len(Format(FailSite(0), "#######"))) & Format(FailSite(0), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite(1), "#######"))) & Format(FailSite(1), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite(2), "#######"))) & Format(FailSite(2), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite(3), "#######"))) & Format(FailSite(3), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite(4), "#######"))) & Format(FailSite(4), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite(5), "#######"))) & Format(FailSite(5), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite(6), "#######"))) & Format(FailSite(6), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailSite(7), "#######"))) & Format(FailSite(7), "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailTotal, "#######"))) & Format(FailTotal, "#######") _
                                 & Space(3) & Space(7 - Len(Format(FailPercent, "0.00%"))) & Format(FailPercent, "0.00%")
            
            ' file output
            Printer.Print "1 PASS" & Space(7) & Space(7 - Len(Format(Bin1Sum(0), "#######"))) & Format(Bin1Sum(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin1Sum(1), "#######"))) & Format(Bin1Sum(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin1Sum(2), "#######"))) & Format(Bin1Sum(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin1Sum(3), "#######"))) & Format(Bin1Sum(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin1Sum(4), "#######"))) & Format(Bin1Sum(4), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin1Sum(5), "#######"))) & Format(Bin1Sum(5), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin1Sum(6), "#######"))) & Format(Bin1Sum(6), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin1Sum(7), "#######"))) & Format(Bin1Sum(7), "#######") _
                                   & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                                   & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
            
            Printer.Print "2 BIN2" & Space(7) & Space(7 - Len(Format(Bin2Sum(0), "#######"))) & Format(Bin2Sum(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Sum(1), "#######"))) & Format(Bin2Sum(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Sum(2), "#######"))) & Format(Bin2Sum(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Sum(3), "#######"))) & Format(Bin2Sum(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Sum(4), "#######"))) & Format(Bin2Sum(4), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Sum(5), "#######"))) & Format(Bin2Sum(5), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Sum(6), "#######"))) & Format(Bin2Sum(6), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Sum(7), "#######"))) & Format(Bin2Sum(7), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Total, "#######"))) & Format(Bin2Total, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin2Percent, "0.00%"))) & Format(Bin2Percent, "0.00%")
            
            Printer.Print "3 BIN3" & Space(7) & Space(7 - Len(Format(Bin3Sum(0), "#######"))) & Format(Bin3Sum(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Sum(1), "#######"))) & Format(Bin3Sum(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Sum(2), "#######"))) & Format(Bin3Sum(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Sum(3), "#######"))) & Format(Bin3Sum(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Sum(4), "#######"))) & Format(Bin3Sum(4), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Sum(5), "#######"))) & Format(Bin3Sum(5), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Sum(6), "#######"))) & Format(Bin3Sum(6), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Sum(7), "#######"))) & Format(Bin3Sum(7), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Total, "#######"))) & Format(Bin3Total, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin3Percent, "0.00%"))) & Format(Bin3Percent, "0.00%")
            
            Printer.Print "4 BIN4" & Space(7) & Space(7 - Len(Format(Bin4Sum(0), "#######"))) & Format(Bin4Sum(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Sum(1), "#######"))) & Format(Bin4Sum(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Sum(2), "#######"))) & Format(Bin4Sum(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Sum(3), "#######"))) & Format(Bin4Sum(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Sum(4), "#######"))) & Format(Bin4Sum(4), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Sum(5), "#######"))) & Format(Bin4Sum(5), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Sum(6), "#######"))) & Format(Bin4Sum(6), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Sum(7), "#######"))) & Format(Bin4Sum(7), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Total, "#######"))) & Format(Bin4Total, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin4Percent, "0.00%"))) & Format(Bin4Percent, "0.00%")
            
            Printer.Print "5 BIN5" & Space(7) & Space(7 - Len(Format(Bin5Sum(0), "#######"))) & Format(Bin5Sum(0), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Sum(1), "#######"))) & Format(Bin5Sum(1), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Sum(2), "#######"))) & Format(Bin5Sum(2), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Sum(3), "#######"))) & Format(Bin5Sum(3), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Sum(4), "#######"))) & Format(Bin5Sum(4), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Sum(5), "#######"))) & Format(Bin5Sum(5), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Sum(6), "#######"))) & Format(Bin5Sum(6), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Sum(7), "#######"))) & Format(Bin5Sum(7), "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Total, "#######"))) & Format(Bin5Total, "#######") _
                                   & Space(3) & Space(7 - Len(Format(Bin5Percent, "0.00%"))) & Format(Bin5Percent, "0.00%")
            Printer.EndDoc
        End If
    End If
End Sub

Sub GetReportSummarySub2()
Dim oDB As ADOX.Catalog
Dim sDBPAth As String
Dim sConStr As String
Dim oCn As ADODB.Connection
Dim oCM As ADODB.Command
Dim RS As ADODB.Recordset

Dim Cmstr11 As String
Dim Cmstr12 As String

Dim Cmstr21 As String
Dim Cmstr22 As String

Dim Cmstr31 As String
Dim Cmstr32 As String

Dim Cmstr41 As String
Dim Cmstr42 As String

Dim Cmstr51 As String
Dim Cmstr52 As String
Dim Cmstr0 As String

Dim cmstr As String
Dim tmp As String

    If ReportDebug = 1 Then
        RunCardNO = "2008CLOP0040"
    End If
    
    '-----------------------------
    ' set Path and connection string
    '---------------------------
    If No8PCard Then
        sDBPAth = "D:\SLT Summary\Summary.mdb"
        If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
            MsgBox "MDB no EXIST"
            Exit Sub
        End If
        sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\SLT.mdb"
    Else
        sDBPAth = "D:\SLT Summary\MultiSummary.mdb"
        If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
            MsgBox "MDB no EXIST"
            Exit Sub
        End If
        sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\MultiSLT.mdb"
    End If

    ' ------------------------
    ' Create New ADOX Object
    ' ------------------------
    Set oCn = New ADODB.Connection
    oCn.ConnectionString = sConStr
    oCn.Open
    
    Set oCM = New ADODB.Command
    oCM.ActiveConnection = oCn
    
    Set RS = New ADODB.Recordset
    
    cmstr = "SELECT Min(Summary.StartAt) as StartATMin,Max(Summary.EndAt) as EndAtMax  " & _
            "from Summary where RunCardNO= '" & RunCardNO & "' and ProcessID= '" & ProcessIDSum & "'"
            
    oCM.CommandText = cmstr
    Debug.Print cmstr
    Set RS = oCM.Execute
    
    StartAtMin = RS.Fields("StartATMin")
    Debug.Print StartAtMin
    
    EndAtMax = RS.Fields("EndAtMax")
    Debug.Print EndAtMax
    
    If No8PCard Then
        cmstr = "SELECT Sum(Summary.Bin1Site1) as Bin1Site1Sum ," & _
            "Sum(Summary.Bin2Site1) as Bin2Site1Sum ," & _
            "Sum(Summary.Bin3Site1) as Bin3Site1Sum ," & _
            "Sum(Summary.Bin4Site1) as Bin4Site1Sum ," & _
            "Sum(Summary.Bin5Site1) as Bin5Site1Sum ," & _
            "Sum(Summary.Bin1Site2) as Bin1Site2Sum ," & _
            "Sum(Summary.Bin2Site2) as Bin2Site2Sum ," & _
            "Sum(Summary.Bin3Site2) as Bin3Site2Sum ," & _
            "Sum(Summary.Bin4Site2) as Bin4Site2Sum ," & _
           "Sum(Summary.Bin5Site2) as Bin5Site2Sum " & _
            "from Summary where RunCardNO= '" & RunCardNO & "' and ProcessID= '" & ProcessIDSum & "' "
        oCM.CommandText = cmstr

        Set RS = oCM.Execute
        
        Bin1Site1Sum = RS.Fields("Bin1Site1Sum")
        Bin2Site1Sum = RS.Fields("Bin2Site1Sum")
        Bin3Site1Sum = RS.Fields("Bin3Site1Sum")
        Bin4Site1Sum = RS.Fields("Bin4Site1Sum")
        Bin5Site1Sum = RS.Fields("Bin5Site1Sum")
        Bin1Site2Sum = RS.Fields("Bin1Site2Sum")
        Bin2Site2Sum = RS.Fields("Bin2Site2Sum")
        Bin3Site2Sum = RS.Fields("Bin3Site2Sum")
        Bin4Site2Sum = RS.Fields("Bin4Site2Sum")
        Bin5Site2Sum = RS.Fields("Bin5Site2Sum")
    Else
        For i = 0 To 6
            Cmstr11 = Cmstr11 & "Sum(Summary.Bin1_" & CStr(i) & ") as Bin1Sum" & CStr(i) & " ,"
        Next i
        Cmstr12 = "Sum(Summary.Bin1_7) as Bin1Sum7,"
         
        For i = 0 To 6
            Cmstr21 = Cmstr21 & "Sum(Summary.Bin2_" & CStr(i) & ") as Bin2Sum" & CStr(i) & " ,"
        Next i
        Cmstr22 = "Sum(Summary.Bin2_7) as Bin2Sum7,"
        
        For i = 0 To 6
            Cmstr31 = Cmstr31 & "Sum(Summary.Bin3_" & CStr(i) & ") as Bin3Sum" & CStr(i) & " ,"
        Next i
        Cmstr32 = "Sum(Summary.Bin3_7) as Bin3Sum7,"
        
        For i = 0 To 6
            Cmstr41 = Cmstr41 & "Sum(Summary.Bin4_" & CStr(i) & ") as Bin4Sum" & CStr(i) & " ,"
        Next i
        Cmstr42 = "Sum(Summary.Bin4_7) as Bin4Sum7,"
        
        For i = 0 To 6
            Cmstr51 = Cmstr51 & "Sum(Summary.Bin5_" & CStr(i) & ") as Bin5Sum" & CStr(i) & " ,"
        Next i
        
        Cmstr52 = "Sum(Summary.Bin5_7) as Bin5Sum7"
        Cmstr0 = " from Summary where RunCardNO= '" & RunCardNO & "' and ProcessID= '" & ProcessIDSum & "' "
         
        oCM.CommandText = "SELECT " & Cmstr11 & Cmstr12 & Cmstr21 & Cmstr22 & Cmstr31 & Cmstr32 & Cmstr41 & Cmstr42 & Cmstr51 & Cmstr52 & Cmstr0
    
        Set RS = oCM.Execute
        
        For i = 0 To 7
            tmp = "Bin1Sum" & CStr(i)
            Bin1Sum(i) = RS.Fields(tmp)
            tmp = "Bin2Sum" & CStr(i)
            Bin2Sum(i) = RS.Fields(tmp)
            tmp = "Bin3Sum" & CStr(i)
            Bin3Sum(i) = RS.Fields(tmp)
            tmp = "Bin4Sum" & CStr(i)
            Bin4Sum(i) = RS.Fields(tmp)
            tmp = "Bin5Sum" & CStr(i)
            Bin5Sum(i) = RS.Fields(tmp)
        Next i
    End If
    
    RS.Close
    
    ' ------------------------
    ' Release / Destroy Objects
    ' ------------------------
    If Not oCM Is Nothing Then Set oCM = Nothing
    If Not oCn Is Nothing Then Set oCn = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    If Not RS Is Nothing Then Set RS = Nothing
    
    ' ------------------------
    ' Error Handling
    ' ------------------------
Err_Handler:

End Sub
Sub PrintReport()
On Error Resume Next
Dim i As Byte

Dim TestedSite(0 To 7) As Long
Dim TestedTotal As Long
Dim TestedPercent As Single
Dim PassTotal As Long
Dim PassPercent As Single
Dim FailTotal As Long
Dim FailSite(0 To 7) As Long
Dim FailPercent As Single
Dim Bin2Total As Long
Dim Bin2Percent As Single
Dim Bin3Total As Long
Dim Bin3Percent As Single
Dim Bin4Total As Long
Dim Bin4Percent As Single
Dim Bin5Total As Long
Dim Bin5Percent As Single

Dim TestedSite1 As Long     ' for dual site only
Dim TestedSite2 As Long     ' for dual site only
 
Dim PassSite1 As Long       ' for dual site only
Dim PassSite2 As Long       ' for dual site only

Dim FailSite1 As Long       ' for dual site only
Dim FailSite2 As Long       ' for dual site only

    If ReportCheck.value = 0 Then
        Exit Sub
    End If

    If DataBaseDebug = 1 Then
        For i = 0 To 7
            Bin1Counter(i) = 2903 + i
            Bin2Counter(i) = 10 + i
            Bin3Counter(i) = 6 + i
            Bin4Counter(i) = 1 + i
            Bin5Counter(i) = 1 + i
        Next
    End If
    
    ' time control
    EndSecond = Format(Now, "HH:MM:SS")
    EndDay = Format(Now, "YYYY/MM/DD")
        
    OutFileName = RunCardNO & "_" & ProcessID & "_" & Left(EndDay, 4) & Mid(EndDay, 6, 2) & Right(EndDay, 2)
    OutFileName = OutFileName & Left(EndSecond, 2) & Mid(EndSecond, 4, 2) & Right(EndSecond, 2) & ".txt"
        
    EndAt = EndDay & Space(1) & EndSecond
    
    If No8PCard Then
    
        TestedSite1 = Bin1Site1 + Bin2Site1 + Bin3Site1 + Bin4Site1 + Bin5Site1
        TestedSite2 = Bin1Site2 + Bin2Site2 + Bin3Site2 + Bin4Site2 + Bin5Site2
        TestedTotal = TestedSite1 + TestedSite2
        
        TestedPercent = 1
        
        PassSite1 = Bin1Site1
        PassSite2 = Bin1Site2
        PassTotal = PassSite1 + PassSite2
        PassPercent = CSng(PassTotal / TestedTotal)
        
        FailSite1 = Bin2Site1 + Bin3Site1 + Bin4Site1 + Bin5Site1
        FailSite2 = Bin2Site2 + Bin3Site2 + Bin4Site2 + Bin5Site2
        FailTotal = FailSite1 + FailSite2
        FailPercent = CSng(FailTotal / TestedTotal)
        
        Bin2Total = Bin2Site1 + Bin2Site2
        Bin2Percent = CSng(Bin2Total / TestedTotal)
            
        Bin3Total = Bin3Site1 + Bin3Site2
        Bin3Percent = CSng(Bin3Total / TestedTotal)
                
        Bin4Total = Bin4Site1 + Bin4Site2
        Bin4Percent = CSng(Bin4Total / TestedTotal)
                    
                
        Bin5Total = Bin5Site1 + Bin5Site2
        Bin5Percent = CSng(Bin5Total / TestedTotal)
        
        Call UpdateDB
    Else
        ' calculate summary
        
        For i = 0 To 7
            TestedSite(i) = Bin1Counter(i) + Bin2Counter(i) + Bin3Counter(i) + Bin4Counter(i) + Bin5Counter(i)
        Next i
         
        For i = 0 To 7
            TestedTotal = TestedTotal + TestedSite(i)
        Next
        TestedPercent = 1
        
        For i = 0 To 7
            PassTotal = PassTotal + Bin1Counter(i)
        Next i
        PassPercent = CSng(PassTotal / TestedTotal)
        
        For i = 0 To 7
            FailSite(i) = Bin2Counter(i) + Bin3Counter(i) + Bin4Counter(i) + Bin5Counter(i)
        Next i
        
        For i = 0 To 7
            FailTotal = FailTotal + FailSite(i)
        Next
        FailPercent = CSng(FailTotal / TestedTotal)
        
        For i = 0 To 7
            Bin2Total = Bin2Total + Bin2Counter(i)
        Next
        Bin2Percent = CSng(Bin2Total / TestedTotal)
            
        For i = 0 To 7
            Bin3Total = Bin3Total + Bin3Counter(i)
        Next i
        
        Bin3Percent = CSng(Bin3Total / TestedTotal)
        
        For i = 0 To 7
            Bin4Total = Bin4Total + Bin4Counter(i)
        Next i
        Bin4Percent = CSng(Bin4Total / TestedTotal)
                    
                
        For i = 0 To 7
            Bin5Total = Bin5Total + Bin5Counter(i)
        Next i
        Bin5Percent = CSng(Bin5Total / TestedTotal)
    
    End If
    

    Open "D:\SLT Summary\" & OutFileName For Output As #1
    
    Print #1, "#####################################################"
    Print #1, "Name of PC: " & NameofPC
    Print #1, "Program Name: " & ProgramName
    Print #1, "Program Version Code: " & ProgramRevisionCode
    Print #1, "Device ID: " & DeviceID
    Print #1, "Run Card NO: " & RunCardNO
    Print #1, "Lot ID: " & LotID
    Print #1, "Process: " & ProcessID
    Print #1, "Start at: " & StartAt
    Print #1, "End at: " & EndAt
    Print #1, "HandlerID: " & HandlerID
    Print #1, "Operator Name: " & OperatorName
    Print #1,
    
    If No8PCard Then
        Print #1, "-------------------------------------------------------"
        Print #1, Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Total  " & Space(3) & "Total"
        Print #1, Space(13) & "COUNT  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percent"
        Print #1, "-------------------------------------------------------"
    Else
        Print #1, "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Print #1, Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Site 3 " & Space(3) & "Site 4 " & Space(3) & "Site 5 " & Space(3) & "Site 6 " & Space(3) & "Site 7 " & Space(3) & "Site 8 " & Space(3) & "Total  " & Space(3) & "Total"
        Print #1, Space(13) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percent"
        Print #1, "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    End If
    
    
    If No8PCard Then
        ' file output
        Print #1, "TESTED" & Space(7) & Space(7 - Len(Format(TestedSite1, "#######"))) & Format(TestedSite1, "#######") _
                                & Space(3) & Space(7 - Len(Format(TestedSite2, "#######"))) & Format(TestedSite2, "#######") _
                                & Space(3) & Space(7 - Len(Format(TestedTotal, "#######"))) & Format(TestedTotal, "#######") _
                                & Space(3) & Space(7 - Len(Format(TestedPercent, "0.00%"))) & Format(TestedPercent, "0.00%")
        
        Print #1, "PASS" & Space(9) & Space(7 - Len(Format(PassSite1, "#######"))) & Format(PassSite1, "#######") _
                               & Space(3) & Space(7 - Len(Format(PassSite2, "#######"))) & Format(PassSite2, "#######") _
                               & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                               & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
        
        Print #1, "FAIL" & Space(9) & Space(7 - Len(Format(FailSite1, "#######"))) & Format(FailSite1, "#######") _
                               & Space(3) & Space(7 - Len(Format(FailSite2, "#######"))) & Format(FailSite2, "#######") _
                               & Space(3) & Space(7 - Len(Format(FailTotal, "#######"))) & Format(FailTotal, "#######") _
                               & Space(3) & Space(7 - Len(Format(FailPercent, "0.00%"))) & Format(FailPercent, "0.00%")
        
        ' file output
        Print #1, "1 PASS" & Space(7) & Space(7 - Len(Format(PassSite1, "#######"))) & Format(PassSite1, "#######") _
                               & Space(3) & Space(7 - Len(Format(PassSite2, "#######"))) & Format(PassSite2, "#######") _
                               & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                               & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
        
        Print #1, "2 BIN2" & Space(7) & Space(7 - Len(Format(Bin2Site1, "#######"))) & Format(Bin2Site1, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin2Site2, "#######"))) & Format(Bin2Site2, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin2Total, "#######"))) & Format(Bin2Total, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin2Percent, "0.00%"))) & Format(Bin2Percent, "0.00%")
        
        Print #1, "3 BIN3" & Space(7) & Space(7 - Len(Format(Bin3Site1, "#######"))) & Format(Bin3Site1, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin3Site2, "#######"))) & Format(Bin3Site2, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin3Total, "#######"))) & Format(Bin3Total, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin3Percent, "0.00%"))) & Format(Bin3Percent, "0.00%")
        
        Print #1, "4 BIN4" & Space(7) & Space(7 - Len(Format(Bin4Site1, "#######"))) & Format(Bin4Site1, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin4Site2, "#######"))) & Format(Bin4Site2, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin4Total, "#######"))) & Format(Bin4Total, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin4Percent, "0.00%"))) & Format(Bin4Percent, "0.00%")
        
        Print #1, "5 BIN5" & Space(7) & Space(7 - Len(Format(Bin5Site1, "#######"))) & Format(Bin5Site1, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin5Site2, "#######"))) & Format(Bin5Site2, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin5Total, "#######"))) & Format(Bin5Total, "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin5Percent, "0.00%"))) & Format(Bin5Percent, "0.00%")
    Else
        ' file output
        Print #1, "TESTED" & Space(7) & Space(7 - Len(Format(TestedSite(0), "#######"))) & Format(TestedSite(0), "#######") _
                                & Space(3) & Space(7 - Len(Format(TestedSite(1), "#######"))) & Format(TestedSite(1), "#######") _
                                & Space(3) & Space(7 - Len(Format(TestedSite(2), "#######"))) & Format(TestedSite(2), "#######") _
                                & Space(3) & Space(7 - Len(Format(TestedSite(3), "#######"))) & Format(TestedSite(3), "#######") _
                                & Space(3) & Space(7 - Len(Format(TestedSite(4), "#######"))) & Format(TestedSite(4), "#######") _
                                & Space(3) & Space(7 - Len(Format(TestedSite(5), "#######"))) & Format(TestedSite(5), "#######") _
                                & Space(3) & Space(7 - Len(Format(TestedSite(6), "#######"))) & Format(TestedSite(6), "#######") _
                                & Space(3) & Space(7 - Len(Format(TestedSite(7), "#######"))) & Format(TestedSite(7), "#######") _
                                & Space(3) & Space(7 - Len(Format(TestedTotal, "#######"))) & Format(TestedTotal, "#######") _
                                & Space(3) & Space(7 - Len(Format(TestedPercent, "0.00%"))) & Format(TestedPercent, "0.00%")
        
        Print #1, "PASS" & Space(9) & Space(7 - Len(Format(Bin1Counter(0), "#######"))) & Format(Bin1Counter(0), "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin1Counter(1), "#######"))) & Format(Bin1Counter(1), "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin1Counter(2), "#######"))) & Format(Bin1Counter(2), "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin1Counter(3), "#######"))) & Format(Bin1Counter(3), "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin1Counter(4), "#######"))) & Format(Bin1Counter(4), "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin1Counter(5), "#######"))) & Format(Bin1Counter(5), "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin1Counter(6), "#######"))) & Format(Bin1Counter(6), "#######") _
                               & Space(3) & Space(7 - Len(Format(Bin1Counter(7), "#######"))) & Format(Bin1Counter(7), "#######") _
                               & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                               & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
        
        Print #1, "FAIL" & Space(9) & Space(7 - Len(Format(FailSite(0), "#######"))) & Format(FailSite(0), "#######") _
                               & Space(3) & Space(7 - Len(Format(FailSite(1), "#######"))) & Format(FailSite(1), "#######") _
                               & Space(3) & Space(7 - Len(Format(FailSite(2), "#######"))) & Format(FailSite(2), "#######") _
                               & Space(3) & Space(7 - Len(Format(FailSite(3), "#######"))) & Format(FailSite(3), "#######") _
                               & Space(3) & Space(7 - Len(Format(FailSite(4), "#######"))) & Format(FailSite(4), "#######") _
                               & Space(3) & Space(7 - Len(Format(FailSite(5), "#######"))) & Format(FailSite(5), "#######") _
                               & Space(3) & Space(7 - Len(Format(FailSite(6), "#######"))) & Format(FailSite(6), "#######") _
                               & Space(3) & Space(7 - Len(Format(FailSite(7), "#######"))) & Format(FailSite(7), "#######") _
                               & Space(3) & Space(7 - Len(Format(FailTotal, "#######"))) & Format(FailTotal, "#######") _
                               & Space(3) & Space(7 - Len(Format(FailPercent, "0.00%"))) & Format(FailPercent, "0.00%")
        
        ' file output
        Print #1, "1 PASS" & Space(7) & Space(7 - Len(Format(Bin1Counter(0), "#######"))) & Format(Bin1Counter(0), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin1Counter(1), "#######"))) & Format(Bin1Counter(1), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin1Counter(2), "#######"))) & Format(Bin1Counter(2), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin1Counter(3), "#######"))) & Format(Bin1Counter(3), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin1Counter(4), "#######"))) & Format(Bin1Counter(4), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin1Counter(5), "#######"))) & Format(Bin1Counter(5), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin1Counter(6), "#######"))) & Format(Bin1Counter(6), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin1Counter(7), "#######"))) & Format(Bin1Counter(7), "#######") _
                           & Space(3) & Space(7 - Len(Format(PassTotal, "#######"))) & Format(PassTotal, "#######") _
                           & Space(3) & Space(7 - Len(Format(PassPercent, "0.00%"))) & Format(PassPercent, "0.00%")
        
        Print #1, "2 BIN2" & Space(7) & Space(7 - Len(Format(Bin2Counter(0), "#######"))) & Format(Bin2Counter(0), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Counter(1), "#######"))) & Format(Bin2Counter(1), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Counter(2), "#######"))) & Format(Bin2Counter(2), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Counter(3), "#######"))) & Format(Bin2Counter(3), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Counter(4), "#######"))) & Format(Bin2Counter(4), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Counter(5), "#######"))) & Format(Bin2Counter(5), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Counter(6), "#######"))) & Format(Bin2Counter(6), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Counter(7), "#######"))) & Format(Bin2Counter(7), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Total, "#######"))) & Format(Bin2Total, "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin2Percent, "0.00%"))) & Format(Bin2Percent, "0.00%")
        
        
        Print #1, "3 BIN3" & Space(7) & Space(7 - Len(Format(Bin3Counter(0), "#######"))) & Format(Bin3Counter(0), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Counter(1), "#######"))) & Format(Bin3Counter(1), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Counter(2), "#######"))) & Format(Bin3Counter(2), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Counter(3), "#######"))) & Format(Bin3Counter(3), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Counter(4), "#######"))) & Format(Bin3Counter(4), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Counter(5), "#######"))) & Format(Bin3Counter(5), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Counter(6), "#######"))) & Format(Bin3Counter(6), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Counter(7), "#######"))) & Format(Bin3Counter(7), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Total, "#######"))) & Format(Bin3Total, "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin3Percent, "0.00%"))) & Format(Bin3Percent, "0.00%")
        
        
        Print #1, "4 BIN4" & Space(7) & Space(7 - Len(Format(Bin4Counter(0), "#######"))) & Format(Bin4Counter(0), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Counter(1), "#######"))) & Format(Bin4Counter(1), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Counter(2), "#######"))) & Format(Bin4Counter(2), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Counter(3), "#######"))) & Format(Bin4Counter(3), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Counter(4), "#######"))) & Format(Bin4Counter(4), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Counter(5), "#######"))) & Format(Bin4Counter(5), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Counter(6), "#######"))) & Format(Bin4Counter(6), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Counter(7), "#######"))) & Format(Bin4Counter(7), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Total, "#######"))) & Format(Bin4Total, "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin4Percent, "0.00%"))) & Format(Bin4Percent, "0.00%")
        
        
        Print #1, "5 BIN5" & Space(7) & Space(7 - Len(Format(Bin5Counter(0), "#######"))) & Format(Bin5Counter(0), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Counter(1), "#######"))) & Format(Bin5Counter(1), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Counter(2), "#######"))) & Format(Bin5Counter(2), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Counter(3), "#######"))) & Format(Bin5Counter(3), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Counter(4), "#######"))) & Format(Bin5Counter(4), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Counter(5), "#######"))) & Format(Bin5Counter(5), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Counter(6), "#######"))) & Format(Bin5Counter(6), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Counter(7), "#######"))) & Format(Bin5Counter(7), "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Total, "#######"))) & Format(Bin5Total, "#######") _
                           & Space(3) & Space(7 - Len(Format(Bin5Percent, "0.00%"))) & Format(Bin5Percent, "0.00%")
    End If
    
    Close #1
    
    '========================= printer section ===========================

    Call GetProcessIDSub
    
End Sub

Sub ReportActive()
Dim winHwnd As Long

    If ReportCheck.value = 1 Then
        ReportForm.Show
        winHwnd = FindWindow(vbNullString, "報表設定")
        SetWindowPos winHwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags
    End If

End Sub

Public Sub MsecDelay(Msec As Single)
Dim start As Single
Dim pause As Single
start = Timer
    Do
        pause = Timer
    Loop Until pause - start >= Msec
End Sub

Sub UpdateDB()
Dim oDB As ADOX.Catalog
Dim sDBPAth As String
Dim sConStr As String
Dim oCn As ADODB.Connection
Dim oCM As ADODB.Command

Dim EndDay As String
Dim EndSecond As String
Dim SNow As String
Dim OutFileName As String

Dim cmstr As String
Dim Cmstr1 As String
Dim Cmstr2 As String
Dim Cmstr3 As String
Dim Cmstr4 As String
Dim Cmstr5 As String
Dim Cmstr0 As String
Dim i As Byte
     
    EndSecond = Format(Now, "HH:MM:SS")
    EndDay = Format(Now, "YYYY/MM/DD")
    EndAt = EndDay & Space(1) & EndSecond
    
    If No8PCard Then
        '-------------------------------
        ' set Path and connection string
        '-------------------------------
        sDBPAth = "D:\SLT Summary\Summary.mdb"
        If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
            MsgBox "MDB no EXIST"
            Exit Sub
        End If
    
        sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\SLT.mdb"
        
        ' ------------------------
        ' Create New ADOX Object
        ' ------------------------
        Set oCn = New ADODB.Connection
        oCn.ConnectionString = sConStr
        oCn.Open
        
        Set oCM = New ADODB.Command
        oCM.ActiveConnection = oCn
    
        cmstr = "UPDATE Summary SET " & _
        "EndAt= '" & EndAt & "' ," & _
        "Bin1Site1=" & CStr(Bin1Site1) & "," & _
        "Bin1Site2=" & CStr(Bin1Site2) & "," & _
        "Bin2Site1=" & CStr(Bin2Site1) & "," & _
        "Bin2Site2=" & CStr(Bin2Site2) & "," & _
        "Bin3Site1=" & CStr(Bin3Site1) & "," & _
        "Bin3Site2=" & CStr(Bin3Site2) & "," & _
        "Bin4Site1=" & CStr(Bin4Site1) & "," & _
        "Bin4Site2=" & CStr(Bin4Site2) & "," & _
        "Bin5Site1=" & CStr(Bin5Site1) & "," & _
        "Bin5Site2=" & CStr(Bin5Site2) & _
        " where StartAT= '" & StartAt & "'"
        oCM.CommandText = cmstr
        Debug.Print cmstr
        oCM.Execute
    
        ' ------------------------
        ' Release / Destroy Objects
        ' ------------------------
        If Not oCM Is Nothing Then Set oCM = Nothing
        If Not oCn Is Nothing Then Set oCn = Nothing
        If Not oDB Is Nothing Then Set oDB = Nothing
    Else
        '-----------------------------
        ' set Path and connection string
        '---------------------------
        sDBPAth = "D:\SLT Summary\MultiSummary.mdb"
        If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
            MsgBox "MDB no EXIST"
            Exit Sub
        End If
    
        sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\MultiSLT.mdb"
         
        ' ------------------------
        ' Create New ADOX Object
        ' ------------------------
        
        Set oCn = New ADODB.Connection
        oCn.ConnectionString = sConStr
        oCn.Open
        
        Set oCM = New ADODB.Command
        oCM.ActiveConnection = oCn
            
        cmstr = "UPDATE Summary SET EndAt= '" & EndAt & "' ,"
        
        For i = 0 To 7
        Cmstr1 = Cmstr1 & "Bin1_" & CStr(i) & "=" & CStr(Bin1Counter(i)) & ","
        Next i
        For i = 0 To 7
        Cmstr2 = Cmstr2 & "Bin2_" & CStr(i) & "=" & CStr(Bin2Counter(i)) & ","
        Next i
        For i = 0 To 7
        Cmstr3 = Cmstr3 & "Bin3_" & CStr(i) & "=" & CStr(Bin3Counter(i)) & ","
        Next i
        For i = 0 To 7
        Cmstr4 = Cmstr4 & "Bin4_" & CStr(i) & "=" & CStr(Bin4Counter(i)) & ","
        Next i
        For i = 0 To 6
        Cmstr5 = Cmstr5 & "Bin5_" & CStr(i) & "=" & CStr(Bin5Counter(i)) & ","
        Next i
        Cmstr0 = "Bin5_7" & "=" & CStr(Bin5Counter(7)) & _
        " where StartAT= '" & StartAt & "'"
        
        oCM.CommandText = cmstr & Cmstr1 & Cmstr2 & Cmstr3 & Cmstr4 & Cmstr5 & Cmstr0
        Debug.Print cmstr
        oCM.Execute
         
        ' ------------------------
        ' Release / Destroy Objects
        ' ------------------------
        If Not oCM Is Nothing Then Set oCM = Nothing
        If Not oCn Is Nothing Then Set oCn = Nothing
        If Not oDB Is Nothing Then Set oDB = Nothing
    End If
    
' ------------------------
' Error Handling
' ------------------------
Err_Handler:

    
End Sub

Sub GetProcessIDSub()

Dim oDB As ADOX.Catalog
Dim sDBPAth As String
Dim sConStr As String
Dim oCn As ADODB.Connection
Dim oCM As ADODB.Command
Dim RS As ADODB.Recordset
Dim cmstr As String

    If ReportDebug = 1 Then
        RunCardNO = "2008CLOP0040"
    End If
    
    If No8PCard Then
        '-------------------------------
        ' set Path and connection string
        '-------------------------------
        sDBPAth = "D:\SLT Summary\Summary.mdb"
        If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
            MsgBox "MDB no EXIST"
            Exit Sub
        End If
        
        sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\SLT.mdb"
    Else
        '-------------------------------
        ' set Path and connection string
        '-------------------------------
        sDBPAth = "D:\SLT Summary\MultiSummary.mdb"
    
        If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
            MsgBox "MDB no EXIST"
            Exit Sub
        End If
    
        sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\MultiSLT.mdb"
    End If
     
    ' ------------------------
    ' Create New ADOX Object
    ' ------------------------
    Set oCn = New ADODB.Connection
    oCn.ConnectionString = sConStr
    oCn.Open
    
    Set oCM = New ADODB.Command
    oCM.ActiveConnection = oCn
    
    Set RS = New ADODB.Recordset
    
    cmstr = "SELECT DISTINCT(Summary.ProcessID)   " & _
            "from Summary where RunCardNO= '" & RunCardNO & "'"
            
    oCM.CommandText = cmstr
    Debug.Print cmstr
    Set RS = oCM.Execute
    
    RS.MoveFirst
    
    ProcessIDSum = ProcessID
    Call PrintReportSummary2

    ' ------------------------
    ' Release / Destroy Objects
    ' ------------------------
    If Not oCM Is Nothing Then Set oCM = Nothing
    If Not oCn Is Nothing Then Set oCn = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    If Not RS Is Nothing Then Set RS = Nothing
    
    ' ------------------------
    ' Error Handling
    ' ------------------------
Err_Handler:

End Sub

Public Sub PCIBinSub(i As Byte)

On Error Resume Next

    Select Case i
        Case 0
            Channel = Channel_P1B
        Case 1
            Channel = Channel_P1C
        Case 2
            Channel = Channel_P2A
        Case 3
            Channel = Channel_P2B
        Case 4
            Channel = Channel_P2C
        Case 5
            Channel = Channel_P3A
        Case 6
            Channel = Channel_P3B
        Case 7
            Channel = Channel_P3C
    End Select
    
    Select Case TestResult(i)
        Case "PASS"
            Call PCI7296_bin(Channel, PCI7296_PASS)
        Case "UNKNOW", "bin2", "Bin2"
            Call PCI7296_bin(Channel, PCI7296_BIN2)
        Case "gponFail", "bin3", "Bin3", "SD_WF", "SD_RF", "CF_WF", "CF_RF"
            Call PCI7296_bin(Channel, PCI7296_BIN3)
        Case "XD_WF", "bin4", "Bin4", "XD_RF"
            Call PCI7296_bin(Channel, PCI7296_BIN4)
        Case "MS_WF", "bin5", "TimeOut", "Bin5", "MS_RF"
            Call PCI7296_bin(Channel, PCI7296_BIN5)
        Case Else
            Call PCI7296_bin(Channel, PCI7296_BIN2)
    End Select

End Sub

Public Sub PCIStateSub()
Dim k As Byte

    For k = 0 To 7
        If OffLineCheck.value = 0 Then
            If SiteCheck(k).value = 1 And GetStart(k) = 1 Then
                If State(k) <> PCIState Then
                    Exit Sub
                End If
            End If
        Else
            If SiteCheck(k).value = 1 Then
                If State(k) <> PCIState Then
                    Exit Sub
                End If
            End If
       End If
    Next k
    
    For k = 0 To 7
        If OffLineCheck.value = 0 Then
            If SiteCheck(k).value = 1 And GetStart(k) = 1 Then
                Call PCIBinSub(k)
                State(k) = IdleState
                MSComm1(k).InBufferCount = 0
                MSComm1(k).InputLen = 0
            End If
        Else
            If SiteCheck(k).value = 1 Then
                Call PCIBinSub(k)
                State(k) = IdleState
                MSComm1(k).InBufferCount = 0
                MSComm1(k).InputLen = 0
            End If
        End If
    Next k
    PCIFlag = 1

End Sub
Public Sub BinSub(i)
On Error Resume Next

Dim tmp As Single
Dim tmpFail As Long
Dim tmps As String
Dim result
Dim k As Byte
Dim CurBinning As String
Dim v As Integer
Dim CurOutputStr As String

    testTime = testTime + 1

    Select Case TestResult(i)
        Case "PASS"
            TestResultLbl(i).Caption = "PASS"
            TestResultLbl(i).BackColor = GREEN_COLOR
            Bin1Counter(i) = Bin1Counter(i) + 1
            Bin1(i).Caption = CStr(Bin1Counter(i))
            Bin2ContiFail(i) = 0
            ContiFailCounter(i) = 0
            CurBinning = 1
            
        Case "UNKNOW", "bin2", "Bin2"
            TestResultLbl(i).Caption = "bin2"
            TestResultLbl(i).BackColor = RED_COLOR
            Bin2Counter(i) = Bin2Counter(i) + 1
            Bin2(i).Caption = CStr(Bin2Counter(i))
            Bin2ContiFail(i) = Bin2ContiFail(i) + 1
            ContiFailCounter(i) = ContiFailCounter(i) + 1
            CurBinning = 2
           
            If Bin2ContiFail(i) > 4 And ResetPC.value = 1 Then
                result = DO_WritePort(card, Channel, PCI7296_PC_RESET)
                Call MsecDelay(PC_RESET_TIME)
                result = DO_WritePort(card, Channel, &HFF)
                Bin2ContiFail(i) = 0
            End If
                    
        Case "gponFail", "bin3", "Bin3", "SD_WF", "SD_RF", "CF_WF", "CF_RF"
            TestResultLbl(i).Caption = "bin3"
            TestResultLbl(i).BackColor = YELLOW_COLOR
            Bin3Counter(i) = Bin3Counter(i) + 1
            ContiFailCounter(i) = ContiFailCounter(i) + 1
            Bin3(i).Caption = CStr(Bin3Counter(i))
            CurBinning = 3

        Case "XD_WF", "bin4", "Bin4", "XD_RF"
            TestResultLbl(i).Caption = "bin4"
            TestResultLbl(i).BackColor = YELLOW_COLOR
            Bin4Counter(i) = Bin4Counter(i) + 1
            ContiFailCounter(i) = ContiFailCounter(i) + 1
            Bin4(i).Caption = CStr(Bin4Counter(i))
            CurBinning = 4
            
        Case "MS_WF", "bin5", "TimeOut", "Bin5", "MS_RF"
           TestResultLbl(i).Caption = "bin5"
            TestResultLbl(i).BackColor = YELLOW_COLOR
            Bin5Counter(i) = Bin5Counter(i) + 1
            ContiFailCounter(i) = ContiFailCounter(i) + 1
            Bin5(i).Caption = CStr(Bin5Counter(i))
            CurBinning = 5
            
        Case Else
            TestResultLbl(i).Caption = "bin2"
            TestResultLbl(i).BackColor = RED_COLOR
            Bin2Counter(i) = Bin2Counter(i) + 1
            Bin2(i).Caption = CStr(Bin2Counter(i))
            ContiFailCounter(i) = ContiFailCounter(i) + 1
            CurBinning = 2
            
    End Select
 
 
    Call UpdateDB

    TempNowStr(i) = CurBinning & TempNowStr(i)
    TempNowStr(i) = Left(TempNowStr(i), 20)
    
    CurOutputStr = ""
    
    For v = Len(TempNowStr(i)) To 1 Step -1
        If v Mod 10 = 0 Then
            CurOutputStr = Mid(TempNowStr(i), v, 1) & vbCrLf & CurOutputStr
        Else
            CurOutputStr = Mid(TempNowStr(i), v, 1) & CurOutputStr
        End If
    Next
    
    
    ContFail(i).Caption = CurOutputStr
    
    
    State(i) = PCIState
      
    tmpFail = (Bin2Counter(i) + Bin3Counter(i) + Bin4Counter(i) + Bin5Counter(i))
    
    TotalFail(i).Caption = CStr(tmpFail)
    tmp = CSng(Bin1Counter(i) / (tmpFail + Bin1Counter(i)))
    tmps = Format$(tmp, "0.00%")
    Yield(i).Caption = tmps
     
    If OffLineCheck.value = 0 Then
    
        BinFlag(i) = 1   ' set Bin flag
    
        OneCycleFlag = 1
        
        For k = 0 To 7
            If BinFlag(k) <> StartFlag(k) Then
                OneCycleFlag = 0
                Exit For
            End If
        Next k
        
        If OneCycleFlag = 1 Then   ' reset all flag
        
            For k = 0 To 7
                BinFlag(i) = 0
                StartFlag(i) = 0
            Next k
            StartCounter = 0
            TotalCycleTime = Timer - MinGetStartTime
            
            totalTestTime = totalTestTime + TotalCycleTime
            avgTestTime = totalTestTime / testTime
            Debug.Print avgTestTime
            avgTestTimeLbl.Caption = "AvgTestTime: " & avgTestTime
            
            TestCycleTimeLbl.Caption = "TestTime:" & CStr(TotalCycleTime)
        End If
    Else
        BinCounter = BinCounter + 1
        Debug.Print "BinCounter="; BinCounter
    
        If BinCounter = OffLineSiteCounter Then
            OneCycleFlag = 1
            TotalCycleTime = Timer - MinGetStartTime
            
            totalTestTime = totalTestTime + TotalCycleTime
            avgTestTime = totalTestTime / testTime
            Debug.Print avgTestTime
            avgTestTimeLbl.Caption = "AvgTestTime: " & avgTestTime
            
            TestCycleTimeLbl.Caption = "TestTime:" & CStr(TotalCycleTime)
        End If
    End If

End Sub

Sub PCI7296_bin(Channel As Byte, PCI7296bin As Byte)

    result = DO_WritePort(card, Channel, PCI7296bin)
    Call Timer_1ms(12)
    
    result = DO_WritePort(card, Channel, PCI7296bin - PCI7296_EOT)
    Call Timer_1ms(7)
            
    result = DO_WritePort(card, Channel, PCI7296bin)
    Call Timer_1ms(7)
           
    result = DO_WritePort(card, Channel, &HFF)
  
End Sub

Function Parser(InputStr As String) As String
Dim TmpStr As String
 
    AU6254Msg = ""
    TmpStr = Trim$(InputStr)
    
    If (InStr(TmpStr, "Rea") = 1) Or _
       (InStr(TmpStr, "ead") = 1) Or _
       (InStr(TmpStr, "ady") = 1) Or _
       (InStr(TmpStr, "dyR") = 1) Or _
       (InStr(TmpStr, "yRe") = 1) Then
        
        Parser = ""
        Exit Function
    End If
    
    If InStr(TmpStr, "Rea") Then
        Parser = Left(TmpStr, Len(TmpStr) - InStr(TmpStr, "Rea") + 1)
    ElseIf InStr(TmpStr, "ead") Then
        Parser = Left(TmpStr, Len(TmpStr) - InStr(TmpStr, "ead") + 1)
    ElseIf InStr(TmpStr, "ady") Then
        Parser = Left(TmpStr, Len(TmpStr) - InStr(TmpStr, "ady") + 1)
    ElseIf InStr(TmpStr, "dyR") Then
        Parser = Left(TmpStr, Len(TmpStr) - InStr(TmpStr, "dyR") + 1)
    ElseIf InStr(TmpStr, "yRe") Then
        Parser = Left(TmpStr, Len(TmpStr) - InStr(TmpStr, "yRe") + 1)
    Else
        Parser = TmpStr
    End If
        
    Select Case Left(TmpStr, 4)
    
    Case "bin2"
        Parser = "bin2"
           
    Case "bin3"
        Parser = "bin3"
             
    Case "UNKN"
        Parser = "UNKNOW"
          
    Case "SD_W"
        Parser = "SD_WF"
           
    Case "SD_R"
        Parser = "SD_RF"
           
    Case "CF_W"
        Parser = "CF_WF"
           
    Case "CF_R"
        Parser = "CF_RF"
    
    Case "XD_W"
        Parser = "XD_WF"
           
    Case "XD_R"
        Parser = "XD_RF"
            
    Case "MS_W"
        Parser = "MS_WF"
           
    Case "MS_R"
        Parser = "MS_RF"
            
    Case "0x90"
        Select Case Mid(TmpStr, 5, 2)
         
        ' Check upstream connect
        Case "F0"
            Parser = "PASS"
                
        Case "E0"
            Parser = "bin2"
            AU6254Msg = "connect error"
         
        Case "E1"
            Parser = "bin2"
            AU6254Msg = "speed error"
            'Hub enumeration
              
        Case "E2"
            Parser = "bin2"
            AU6254Msg = "unkown device error"
              
        ' Check downstream port1-->(connect hub64 module)
        Case "E3"
            Parser = "bin3"
            AU6254Msg = "port1 connect error"
              
        Case "E4"
            Parser = "bin3"
            AU6254Msg = "port1 speed error"
         
        Case "E5"
            Parser = "bin3"
            AU6254Msg = "TT error"
              
        ' Check downstream port2-->(connect high speed module)
        Case "E6"
            Parser = "bin4"
            AU6254Msg = "port2 connect error"
         
        Case "E7"
            Parser = "bin4"
            AU6254Msg = "port2 speed error"
              
        Case "E8"
            Parser = "bin4"
            AU6254Msg = "Port2 control error"
         
        Case "E9"
            Parser = "bin4"
            AU6254Msg = "Port2 bulk error"
              
        Case "EA"
            Parser = "bin4"
            AU6254Msg = "Port2 interrupt error"
         
        Case "EB"
            Parser = "bin4"
            AU6254Msg = "Port2 isochronous error"
              
        'Check downstream port3-->(connect full speed module)
        Case "EC"
            Parser = "bin4"
            AU6254Msg = "port3 connect error"
         
        Case "ED"
            Parser = "bin4"
            AU6254Msg = "port3 speed error"
          
        Case "EE"
            Parser = "bin4"
            AU6254Msg = "Port3 control error"
         
        Case "EF"
            Parser = "bin4"
            AU6254Msg = "Port3 bulk error"
              
        Case "0E"
            AU6254Msg = "Port3 interrupt error"
            Parser = "bin4"
         
        Case "1E"
            Parser = "bin4"
            AU6254Msg = "Port3 isochronous error"
               
        'Check downstream port4-->(connect low speed module)
        Case "2E"
            AU6254Msg = "port4 connect error "
            Parser = "bin4"
        
        Case "3E"
            AU6254Msg = "port4 speed error"
            Parser = "bin4"
        
        Case "4E"
            AU6254Msg = "Port4 control error "
            Parser = "bin4"
        
        Case "5E"
            AU6254Msg = "Port4 bulk error "
            Parser = "bin4"
               
        Case "6E"
            AU6254Msg = "Port4 interrupt error "
            Parser = "bin4"
        
        Case "7E"
            AU6254Msg = "Port4 isochronous error "
            Parser = "bin4"
               
        'Suspend and Resume and Current
        Case "8E"
            AU6254Msg = "global suspend current error"
            Parser = "bin5"
        
        Case "9E"
            AU6254Msg = "global resume error"
            Parser = "bin5"
            
        'Check Disconnect
        Case "AE"
            AU6254Msg = "port1 disconnect error "
            Parser = "bin5"
        Case "BE"
            AU6254Msg = "port2 disconnect error "
            Parser = "bin5"
            
        Case "CE"
            AU6254Msg = "port3 disconnect error "
            Parser = "bin5"
               
        Case "DE"
            AU6254Msg = "port4 disconnect error "
               
        End Select
                 
    Case Else
        If InStr(TmpStr, "PAS") Then
            Parser = "PASS"
        End If
            
    End Select

End Function
Public Sub FirstTimeSub()
Dim k As Byte
Dim tmp As Byte

    SelectFailFlag = 0
    If ChipName = "" Then
        MsgBox "please select Chip"
        SelectFailFlag = 1
        Exit Sub
    End If

    If FirstTimeFlag = 0 Then   ' 僅第一次進入sub才會執行
        BinCounter = 0
        StartCounter = 0
        OffLineSiteCounter = 0
        
        For k = 0 To 7
            StartFlag(k) = 0
            BinFlag(k) = 0
            State(k) = IdleState
        Next k
        ' card initial for 7248 or 7296
        If No8PCard = False Then
            tmp = PCI7296_CLEAR_START                       ' 10 111 111
            result = DO_WritePort(card, Channel_P1B, tmp)   ' 1
            result = DO_WritePort(card, Channel_P1C, tmp)   ' 2
            result = DO_WritePort(card, Channel_P2A, tmp)   ' 5
            result = DO_WritePort(card, Channel_P2B, tmp)   ' 6
            result = DO_WritePort(card, Channel_P2C, tmp)   ' 7
            result = DO_WritePort(card, Channel_P3A, tmp)   ' 10
            result = DO_WritePort(card, Channel_P3B, tmp)   ' 11
            result = DO_WritePort(card, Channel_P3C, tmp)   ' 12
            
            Call Timer_1ms(10)
               
            tmp = &HFF
            result = DO_WritePort(card, Channel_P1B, tmp)
            result = DO_WritePort(card, Channel_P1C, tmp)
            result = DO_WritePort(card, Channel_P2A, tmp)
            result = DO_WritePort(card, Channel_P2B, tmp)
            result = DO_WritePort(card, Channel_P2C, tmp)
            result = DO_WritePort(card, Channel_P3A, tmp)
            result = DO_WritePort(card, Channel_P3B, tmp)
            result = DO_WritePort(card, Channel_P3C, tmp)
        Else
            result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
            result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
            If ChipName = "AU6366S4" Or ChipName = "AU66S4_F" Or ChipName = "AU9520" Then
                result = DIO_PortConfig(card, Channel_P1CH, INPUT_PORT)
                result = DIO_PortConfig(card, Channel_P1CL, INPUT_PORT)
            Else
                result = DIO_PortConfig(card, Channel_P1CH, OUTPUT_PORT)
                result = DIO_PortConfig(card, Channel_P1CL, OUTPUT_PORT)
            End If
            result = DIO_PortConfig(card, Channel_P2A, OUTPUT_PORT)
            result = DIO_PortConfig(card, Channel_P2B, OUTPUT_PORT)
            result = DIO_PortConfig(card, Channel_P2CH, INPUT_PORT)
            result = DIO_PortConfig(card, Channel_P2CL, INPUT_PORT)
        End If
        
        If OffLineCheck.value = 0 Then
            For i = 0 To CInt(SiteCombo.Text) - 1   ' auto mode, set選取的port數
                Status(i).Caption = IdleState
                'SiteCheck(i).value = 1
            Next i
        Else
            For i = 0 To 7                          ' off line mode, set全部port
                If SiteCheck(i).value = 1 Then
                    OffLineSiteCounter = OffLineSiteCounter + 1
                    Status(i).Caption = IdleState
                End If
            Next i
                    
            If OffLineSiteCounter = 0 Then
                MsgBox "Please select site"
                SelectFailFlag = 1
                Exit Sub
            End If
        End If
            
        For i = 0 To 7
            Status(i).BackColor = BACK_COLOR
        Next
    
        For i = 0 To 7
            Bin1Counter(i) = 0
            Bin2Counter(i) = 0
            Bin3Counter(i) = 0
            Bin4Counter(i) = 0
            Bin5Counter(i) = 0
        Next
    
        FirstTimeFlag = 1
    
    End If

End Sub
Public Function PassTimeFcn(i As Byte) As Single

    PassTimeFcn = Timer - Fire(i)
    
    If PassTimeFcn < 0 Then
        PassTimeFcn = 86400 + PassTimeFcn
    End If

End Function
Private Sub Initial_7248()
Dim i As Integer, j As Integer
Dim result As Integer
  
    For i = 0 To 1  'Initial status is Output for all channels
        result = DIO_PortConfig(card, i * 5 + Channel_P1A, OUTPUT_PORT)
        value_a(i) = &HFF
        result = DO_WritePort(card, i * 5 + Channel_P1A, value_a(i))
        '===================================================================
        result = DIO_PortConfig(card, i * 5 + Channel_P1B, OUTPUT_PORT)
        value_b(i) = &HFF
        result = DO_WritePort(card, i * 5 + Channel_P1B, value_b(i))
        '===================================================================
        result = DIO_PortConfig(card, i * 5 + Channel_P1CH, OUTPUT_PORT)
        value_cu(i) = &HF
        result = DO_WritePort(card, i * 5 + Channel_P1CH, value_cu(i))
        '===================================================================
        result = DIO_PortConfig(card, i * 5 + Channel_P1CL, OUTPUT_PORT)
        value_cl(i) = &HF
        result = DO_WritePort(card, i * 5 + Channel_P1CL, value_cl(i))
    Next i
End Sub
Private Sub Initial_7296()
  
Dim result As Integer
Dim tmp As Byte

    result = DIO_PortConfig(card, Channel_P1A, INPUT_PORT)
    result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
    result = DIO_PortConfig(card, Channel_P1C, OUTPUT_PORT)
    result = DIO_PortConfig(card, Channel_P2A, OUTPUT_PORT)
    result = DIO_PortConfig(card, Channel_P2B, OUTPUT_PORT)
    result = DIO_PortConfig(card, Channel_P2C, OUTPUT_PORT)
    result = DIO_PortConfig(card, Channel_P3A, OUTPUT_PORT)
    result = DIO_PortConfig(card, Channel_P3B, OUTPUT_PORT)
    result = DIO_PortConfig(card, Channel_P3C, OUTPUT_PORT)
    result = DIO_PortConfig(card, Channel_P4A, OUTPUT_PORT)
    result = DIO_PortConfig(card, Channel_P4B, OUTPUT_PORT)
    result = DIO_PortConfig(card, Channel_P4C, OUTPUT_PORT)
        
    tmp = PCI7296_CLEAR_START
    result = DO_WritePort(card, Channel_P1B, tmp)
         
    result = DO_WritePort(card, Channel_P1C, tmp)
    result = DO_WritePort(card, Channel_P2A, tmp)
    result = DO_WritePort(card, Channel_P2B, tmp)
    result = DO_WritePort(card, Channel_P2C, tmp)
    result = DO_WritePort(card, Channel_P3A, tmp)
    result = DO_WritePort(card, Channel_P3B, tmp)
    result = DO_WritePort(card, Channel_P3C, tmp)
    result = DO_WritePort(card, Channel_P4A, tmp)
    result = DO_WritePort(card, Channel_P4B, tmp)
    result = DO_WritePort(card, Channel_P4C, tmp)
    
    Call Timer_1ms(10)
    
    tmp = &HFF
    result = DO_WritePort(card, Channel_P1B, tmp)
         
    result = DO_WritePort(card, Channel_P1C, tmp)
    result = DO_WritePort(card, Channel_P2A, tmp)
    result = DO_WritePort(card, Channel_P2B, tmp)
    result = DO_WritePort(card, Channel_P2C, tmp)
    result = DO_WritePort(card, Channel_P3A, tmp)
    result = DO_WritePort(card, Channel_P3B, tmp)
    result = DO_WritePort(card, Channel_P3C, tmp)
    result = DO_WritePort(card, Channel_P4A, tmp)
    result = DO_WritePort(card, Channel_P4B, tmp)
    result = DO_WritePort(card, Channel_P4C, tmp)

End Sub
Sub SetTimer_1ms()
Dim err As Integer
    err = CTR_Setup(card, 1, RATE_GENERATOR, 200, BINTimer)
    err = CTR_Setup(card, 2, RATE_GENERATOR, 10, BINTimer)
End Sub
Private Sub Timer_1ms(ms As Integer)
Dim result
Dim old_value1
Dim old_value2
 
Dim i As Integer
 
result = CTR_Read(0, 2, old_value1)

   For i = 1 To ms
        Do
            DoEvents
            result = CTR_Read(0, 2, old_value2)
        Loop Until old_value1 <> old_value2 Or AllenStop = 1

        Do
            DoEvents
            result = CTR_Read(0, 2, old_value2)
        Loop Until old_value1 = old_value2 Or AllenStop = 1
    Next
     
End Sub

Private Sub BeginBtn_Click()
Dim i As Byte

    avgTestTime = 0
    totalTestTime = 0
    testTime = 0
    
    If Dummy.value = True Then                      ' 無8port卡就無視handler type
        If No8PCard = False Then
            MsgBox ("Please Select Handler Type")
            Blank_Timer.Enabled = True
            Exit Sub
        End If
    End If
        
    PrintEnable = 0
    StopFlag = 0
    AllenStop = 0
    
'    ' for summary initialize
'    For i = 0 To 7
'        TestedSite(i) = 0
'        FailSite(i) = 0
'    Next i
'
'    TestedTotal = 0
'    TestedPercent = 0
'    PassTotal = 0
'    PassPercent = 0
'    FailTotal = 0
'    FailPercent = 0
'    Bin2Total = 0
'    Bin2Percent = 0
'    Bin3Total = 0
'    Bin3Percent = 0
'    Bin4Total = 0
'    Bin4Percent = 0
'    Bin5Total = 0
'    Bin5Percent = 0
'    TestedSite1 = 0
'    TestedSite2 = 0
'    PassSite1 = 0
'    PassSite2 = 0
'    FailSite1 = 0
'    FailSite2 = 0
    
    Call FirstTimeSub
    
    If SelectFailFlag = 1 Then
        Exit Sub
    End If

    Call LockOption

    HubTestEnd = 0
    SiteCheckCount = 0
    MPChipName = ""
    RealChipName = ""
    
    If PreviousChipName = "" Then
        PreviousChipName = ChipName
    Else
        If PreviousChipName <> ChipName Then
            ReportForm.DeviceIDText = ""
            ReportForm.RunCardNOText = ""
            ReportForm.LotIDText = ""
            ReportForm.OperatorNameText = ""
            ReportForm.ProcessIDCombo = ""
            PreviousChipName = ChipName
            ProgramName = Trim(ChipNameCombo.Text) & Trim(ChipNameCombo2.Text)
        End If
    End If
    
    For i = 0 To 7
        EnCheck(i) = False
        GetBinning(i) = False
        ContiFailCounter(i) = 0
    Next
    
    ReportBegin = 0
    
    Call ReportActive
    
    If ReportCheck.value = 1 Then
        Do
            DoEvents
        Loop While ReportBegin = 0 And AllenStop = 0
    End If
    
    ReportBegin = 0
    GetFirstStart = False
    
    Do
MultiLoop:
        If No8PCard = False Then
            For i = 0 To 7
                
                DoEvents
                
                If SiteCheck(i).value = 1 Then
                    ' assign Channel
                    Select Case i
                        Case 0
                            Channel = Channel_P1B
                        Case 1
                            Channel = Channel_P1C
                        Case 2
                            Channel = Channel_P2A
                        Case 3
                            Channel = Channel_P2B
                        Case 4
                            Channel = Channel_P2C
                        Case 5
                            Channel = Channel_P3A
                        Case 6
                            Channel = Channel_P3B
                        Case 7
                            Channel = Channel_P3C
                    End Select
        
                    ' assign State
                    Debug.Print "q"; i; State(i)
                    Select Case State(i)
                        Case IdleState
                            Debug.Print "1"
                            Call HandlerStartSub(i)             ' 1. wait start  2. change to HandlerStartState
                            Status(i).BackColor = GREEN_COLOR
                        Case HandlerStartState
                            Debug.Print "2"
                            Call PCReadySub(i)                  ' 1. wait for PC reader
                            Status(i).BackColor = GREEN_COLOR
                        Case PCResetState
                            Debug.Print "3"
                            Call PCResetSub(i)
                            Status(i).BackColor = RED_COLOR
                        Case PCResetFailState
                            Debug.Print "4"
                            Call PCResetFailSub(i)
                            Status(i).BackColor = RED_COLOR
                        Case PCReadyState
                            Debug.Print "5"
                            Call TestingSub(i)                  ' next state
                            Status(i).BackColor = YELLOW_COLOR
                        Case BinState
                            Debug.Print "6"
                            Call BinSub(i)                      ' back to idle state
                            Status(i).BackColor = GREEN_COLOR
                        
                            If (Bin2ContiFail(i) >= 3) And (HubNonUPT2Flag = False) Then
                                ResetUPT2_Flag = True
                            End If
                        Case PCIState
                            Call PCIStateSub
                            If PCIFlag = 1 Then
                                If (OneCycleCheck.value = 1) Then
                                    AllenStop = 1
                                    MsgBox "one cycle test Finish"
                                    FirstTimeFlag = 0
                                End If
          
                                If StopFlag = 1 Then
                                    AllenStop = 1
                                    MsgBox "STOP test Finish"
                                    FirstTimeFlag = 0
                                End If
                                PCIFlag = 0
                                GoTo MultiLoop
                            End If
                    End Select
                End If
            
                If UPT2TestFlag = True Then
                    
                    If (State(i) = IdleState) Then
                        GetBinning(i) = False
                    End If
                    
                    If (GetBinning(i) = False) And (State(i) = BinState) Then
                        GetBinning(i) = True
                        RealSiteCount = RealSiteCount + 1
                    End If
            
                    If (RealSiteCount <> 0) And (RealSiteCount = SiteCheckCount) Then
                        HubTestEnd = 1
                        GetFirstStart = False
                    End If
            
                End If
            
            Next
          
            For i = 0 To 7
                Status(i).Caption = State(i)
            Next i
    
            
            '20130327 modify reset UPT2 when test end (orginal: HubEnaOn = 1)
            If (ResetUPT2_Flag = True) And (HubTestEnd = 1) Then
                If Handler_NS.value = True Then
                    result = DO_WritePort(card, Channel_P3A, &HFC)
                    Call MsecDelay(0.2)
                    result = DO_WritePort(card, Channel_P3A, &HFF)
                End If
                        
                If Handler_SRM.value = True Then
                    result = DO_WritePort(card, Channel_P4A, &HFC)
                    Call MsecDelay(0.2)
                    result = DO_WritePort(card, Channel_P4A, &HFF)
                End If
                    
                ResetUPT2_Flag = False
            End If
          
            If (UPT2TestFlag = True) And (HubTestEnd = 1) Then
                If ((HubEnaOn = 1) And (Host.Handler_NS = True)) Then
                    result = DO_WritePort(card, Channel_P3B, &HFF)      ' Site1 ~ Site4 (bit1~bit4) Ena Off
                    HubEnaOn = 0
                    HubTestEnd = 0
                    RealSiteCount = 0
                    SiteCheckCount = 0
                    For i = 0 To 7
                        EnCheck(i) = False
                        GetBinning(i) = False
                    Next
                End If
            
                If ((HubEnaOn = 1) And (Host.Handler_SRM = True)) Then
                    result = DO_WritePort(card, Channel_P4B, &HFF)      ' Site1 ~ Site8 (bit1~bit8) Ena OFF
                    HubEnaOn = 0
                    HubTestEnd = 0
                    RealSiteCount = 0
                    SiteCheckCount = 0
                    For i = 0 To 7
                        EnCheck(i) = False
                        GetBinning(i) = False
                    Next
                End If
            End If
          
            If StopFlag = 1 Then
                AllenStop = 1
                FirstTimeFlag = 0
            End If
            
            OneCycleFlag = 0
        
        Else
            For i = 0 To 1
                Status(i).Caption = "Testing"
            Next i
        
            testTime = testTime + 1
            
            MinGetStartTime = Timer
            Call DualCommonTest
            Call DualTestResult
            TotalCycleTime = Timer - MinGetStartTime
            
            totalTestTime = totalTestTime + TotalCycleTime
            avgTestTime = totalTestTime / testTime
            Debug.Print avgTestTime
            avgTestTimeLbl.Caption = "AvgTestTime: " & avgTestTime
            
            TestCycleTimeLbl.Caption = "TestTime:" & CStr(TotalCycleTime)
            
            If OneCycleCheck.value = 1 Then
                AllenStop = 1
            End If
            
        End If
        
    Loop While AllenStop = 0
        
    Call UnlockOption
    
    If OneCycleCheck.value = 0 Then
        If No8PCard Then
            For i = 0 To 1
                Status(i).Caption = "Done"
            Next i
        End If
        
        Call PrintReport
    End If

End Sub
Public Sub HandlerStartSub(i)

'Step1: wait for Handle Start

Dim result
Dim k As Byte
Dim j As Integer
Dim ReadStartSignal
Dim UPT2DetectStartTime
 
    TestResult(i) = ""
    
    If OffLineCheck.value = 0 Then
    
        If (UPT2TestFlag = True) And (SiteCheckCount = 0) Then          ' 選了ver就會設定UPT2TestFlag為true
            
            result = DO_ReadPort(card, Channel_P1A, ReadStartSignal)    ' P1A為INPUT_PORT
            
            If ReadStartSignal <> &HFF Then
                For j = 0 To 7
                    OldState(j) = State(j)
                    GetStart(j) = 0
                    
                    'Change to Next State
                    If OldState(j) <> State(j) Then
                        Fire(j) = Timer
                    End If
                Next
                
                UPT2DetectStartTime = Timer
                
                Do
                    For j = 0 To 7
                        If (CAndValue(ReadStartSignal, CPort(j)) = 0) And (EnCheck(j) = False) Then
                        
                            If GetFirstStart = False Then
                                GetFirstStart = True
                            End If
                            
                            GetStart(j) = 1
                            State(j) = HandlerStartState
                            CycleTestTime(j) = Timer
                            StartFlag(i) = 1
                            EnCheck(j) = True
                        
                            If StartCounter = 0 Then
                                OneCycleFlag = 0
                            End If
                    
                            SiteCheckCount = SiteCheckCount + 1
                            StartCounter = StartCounter + 1
                    
                            If StartCounter = 1 Then
                                OneCycleFlag = 0
                                MinGetStartTime = Timer
                            End If
                    
                            If SiteCheckCount = 1 Then
                                If ((HubEnaOn = 0) And (Host.Handler_NS = True)) Then
                                    result = DO_WritePort(card, Channel_P3B, &HF0)  'Site1 ~ Site4 (bit1~bit4) Ena ON
                                    HubEnaOn = 1
                                End If
     
                                If ((HubEnaOn = 0) And (Host.Handler_SRM = True)) Then
                                    result = DO_WritePort(card, Channel_P4B, &H0)   'Site1 ~ Site8 (bit1~bit8) Ena ON
                                    HubEnaOn = 1
                                End If
                            End If
                            
                            PassTime(j) = PassTimeFcn(CByte(j))
                            
                            If PassTime(j) > HandlerStartTimeOut Then
                                TestResultLbl(j).Caption = "Start Time Out"
                            End If
                            
                            Debug.Print "StartCounter="; StartCounter
                        End If
                    Next
                Loop Until (Timer - UPT2DetectStartTime > 0.5) Or (SiteCheckCount = CInt(SiteCombo))
                
            End If
                
        Else
            result = DO_ReadPort(card, Channel_P1A, ReadStartSignal)
        
            If OldState(i) <> State(i) Then
                Fire(i) = Timer
            End If
        
            OldState(i) = State(i)
            GetStart(i) = 0
            'Change to Next State
            
            If CAndValue(ReadStartSignal, CPort(i)) = 0 Then
                GetStart(i) = 1
                State(i) = HandlerStartState
                CycleTestTime(i) = Timer
                StartFlag(i) = 1
              
                If StartCounter = 0 Then
                    OneCycleFlag = 0
                 End If
              
                StartCounter = StartCounter + 1
                If StartCounter = 1 Then
                    OneCycleFlag = 0
                    MinGetStartTime = Timer
                End If
          
                Debug.Print "StartCounter="; StartCounter
                Exit Sub
            End If
    
            ' Time Out Conditions
            PassTime(i) = PassTimeFcn(CByte(i))
            If PassTime(i) > HandlerStartTimeOut Then
                TestResultLbl(i).Caption = "Start Time Out"
            End If
        End If
    Else
        If OffLineSiteCounter = BinCounter Then
            OneCycleFlag = 0
            StartCounter = 0
            BinCounter = 0
                    
            For k = 0 To 7
                StartFlag(k) = 0
                BinFlag(k) = 0
            Next k
        End If
            
        If StartFlag(i) = 1 Then
            Exit Sub
        End If
               
        StartFlag(i) = 1
               
        Call MsecDelay(0.1)
        State(i) = HandlerStartState
        CycleTestTime(i) = Timer
        
        StartCounter = StartCounter + 1
        
        If StartCounter = 1 Then
            MinGetStartTime = Timer
        End If
        Debug.Print "StartCounter="; StartCounter; "State"; State(i)
               
    End If

End Sub

Public Sub PCReadySub(i As Byte)

Dim j As Integer
 
    '1
    If OldState(i) <> State(i) Then
       Fire(i) = Timer
    End If
    OldState(i) = State(i)

    '2 Action
    If (UPT2TestFlag = True) And (HubNonUPT2Flag = False) Then
        MSComm1(i).Output = "~"
    End If
     
    TmpBuf = MSComm1(i).Input
    Buf(i) = Buf(i) & TmpBuf
    
    If InStr(1, Buf(i), "Rea") <> 0 Then
        If (UPT2TestFlag = True) And (HubNonUPT2Flag = False) Then
            For j = 1 To Len(ChipName)
                MSComm1(i).Output = Mid(ChipName, j, 1)
                Call MsecDelay(0.02)
            Next
        Else
            
            If ((ContiFailCounter(0) >= 5) Or _
                (ContiFailCounter(1) >= 5) Or _
                (ContiFailCounter(2) >= 5) Or _
                (ContiFailCounter(3) >= 5) Or _
                (ContiFailCounter(4) >= 5) Or _
                (ContiFailCounter(5) >= 5) Or _
                (ContiFailCounter(6) >= 5) Or _
                (ContiFailCounter(7) >= 5)) And (InStr(ChipName, "U69") = 2) And (Len(ChipName) = 14) And (Mid(ChipName, 12, 1) <> "U") Then
                SendMP_Flag = True
                If (RealChipName = "") And (MPChipName = "") Then
                    RealChipName = Trim(ChipNameCombo.Text) & Trim(ChipNameCombo2.Text)
                    MPChipName = Left(ChipName, 10) & "M" & Right(ChipName, 3)
                End If
                ChipName = MPChipName
            ElseIf (InStr(ChipName, "U69") = 2) Then
                SendMP_Flag = False
                ChipName = Trim(ChipNameCombo.Text) & Trim(ChipNameCombo2.Text)
            End If
            
            MSComm1(i).Output = ChipName
            
            If SendMP_Flag = True Then
                ChipName = RealChipName
            End If
            
        End If
       
        MSComm1(i).InBufferCount = 0
        MSComm1(i).InputLen = 0
        Buf(i) = ""
        State(i) = PCReadyState
        Exit Sub
    End If
    
    ' Time Out Condition
    PassTime(i) = PassTimeFcn(i)
    
    If PassTime(i) > PCReadyTimeOut Then
        TestResultLbl(i).Caption = "PC Ready Time Out"
        'Reset PC
        State(i) = PCResetState
        result = DO_WritePort(card, Channel, PCI7296_PC_RESET)
        Call MsecDelay(PC_RESET_TIME)
        result = DO_WritePort(card, Channel, &HFF)
    End If

End Sub

Public Sub PCResetFailSub(i As Byte)

    'Step2: wait for PC Ready
    
    '1
    If OldState(i) <> State(i) Then
        Fire(i) = Timer
    End If
    OldState(i) = State(i)
    
    '2 Action
    TmpBuf = MSComm1(i).Input
    Buf(i) = Buf(i) & TmpBuf
    If InStr(1, Buf(i), "Rea") <> 0 Then
        MSComm1(i).Output = ChipName
        MSComm1(i).InBufferCount = 0
        MSComm1(i).InputLen = 0
        State(i) = PCReadyState
        Buf(i) = ""
        Exit Sub
    End If
    
    'Time Out Condition
    PassTime(i) = PassTimeFcn(i)
    
    If PassTime(i) > PCResetTimeOut Then
        TestResultLbl(i).Caption = "PC Reset Fail 2"
        result = DO_WritePort(card, Channel, PCI7296_PC_RESET)
        Call MsecDelay(PC_RESET_TIME)
        result = DO_WritePort(card, Channel, &HFF)
        State(i) = PCResetState ' Reset PC
    End If
 
End Sub
Public Sub PCResetSub(i As Byte)

    'Step2: wait for PC Ready
    
    '1
    If OldState(i) <> State(i) Then
       Fire(i) = Timer
    End If
    OldState(i) = State(i)
    
    '2 Action
    TmpBuf = MSComm1(i).Input
    Buf(i) = Buf(i) & TmpBuf
    If InStr(1, Buf(i), "Rea") <> 0 Then
        MSComm1(i).Output = ChipName
        MSComm1(i).InBufferCount = 0
        MSComm1(i).InputLen = 0
        State(i) = PCReadyState
        Buf(i) = ""
        Exit Sub
    End If
    
    PassTime(i) = PassTimeFcn(i)
    
    
    If PassTime(i) > PCResetTimeOut Then
        TestResultLbl(i).Caption = "PC Reset Fail 1"
        result = DO_WritePort(card, Channel, PCI7296_PC_RESET)
        Call MsecDelay(PC_RESET_TIME)
        result = DO_WritePort(card, Channel, &HFF)
        State(i) = PCResetFailState ' Reset PC
    End If

End Sub

Public Sub TestingSub(i As Byte)
    '1
    If OldState(i) <> State(i) Then
        Fire(i) = Timer
    End If
    OldState(i) = State(i)
    
    '2 Action
    If (UPT2TestFlag = True) And (HubNonUPT2Flag = False) Then
        MSComm1(i).Output = "~"
    End If
    
    If MSComm1(i).InBufferCount >= 3 Then
        TmpBuf = MSComm1(i).Input
        TestResult(i) = TestResult(i) & TmpBuf
    End If
     
    TestResult(i) = Parser(TestResult(i))
     
    If TestResult(i) <> "" Then
        State(i) = BinState
        Exit Sub
    End If
    
    ' Time Out Condition
    PassTime(i) = PassTimeFcn(i)
    
    If PassTime(i) > TestTimeOut Then
        TestResultLbl(i).Caption = "Test Time Out"
        TimeOutCounter(i) = TimeOutCounter(i) + 1
        TimeOut(i).Caption = CStr(TimeOutCounter(i))
        'Reset PC
        State(i) = BinState
    End If

End Sub

Private Sub Blank_Timer_Timer()
    If Handler_Blank.BackColor = System_Color Then
        Handler_Blank.BackColor = Blank_Color
        Handler_NS.BackColor = Blank_Color
        Handler_SRM.BackColor = Blank_Color
    Else
        Handler_Blank.BackColor = System_Color
        Handler_NS.BackColor = System_Color
        Handler_SRM.BackColor = System_Color
    End If
End Sub

Private Sub ChipNameCombo_Click()
Dim DB_Path As String
Dim MPTesterRS As New ADODB.Recordset
Dim TesterRS As New ADODB.Recordset
Dim ConnMPTesterDB As New ADODB.Connection
Dim ConnTesterDB As New ADODB.Connection
Dim MpTesterDB_Path As String
Dim TesterDB_Path As String
Dim NameTmp As String
Dim DbPath As String
Dim IC_Name As String

    IC_Name = Trim(ChipNameCombo)
    ChipNameCombo2.Clear
    ChipNameCombo2.Enabled = False

    '=============================================
        'connection to MPTester.mdb
    '=============================================
    MpTesterDB_Path = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\PGM_ListDB\MPTester_" & LastMPTesterDateCode & ".mdb"
    ConnMPTesterDB.Open MpTesterDB_Path
    MPTesterRS.CursorLocation = adUseClient
    MPTesterRS.Open "MPTester", ConnMPTesterDB, adOpenKeyset, adLockPessimistic
    MPTesterRS.MoveFirst
    Set MPTesterRS = ConnMPTesterDB.Execute("Select * From [MPTester] Where [PGM_Name] LIKE '" & IC_Name & "%" & "' Order By [PGM_Name]")
    
    Do Until MPTesterRS.EOF = True
        If MPTesterRS.Fields(0) = 1 Then
            ChipNameCombo2.AddItem Mid(MPTesterRS.Fields(1), 7, Len(MPTesterRS.Fields(1)))
        End If
        MPTesterRS.MoveNext
    Loop
    
    '=============================================
        'connection to Tester.mdb
    '=============================================
    
    TesterDB_Path = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\PGM_ListDB\Tester_" & LastTesterDateCode & ".mdb"
    ConnTesterDB.Open TesterDB_Path
    TesterRS.CursorLocation = adUseClient
    TesterRS.Open "Tester", ConnTesterDB, adOpenKeyset, adLockPessimistic
    TesterRS.MoveFirst
    Set TesterRS = ConnTesterDB.Execute("Select * From [Tester] Where [PGM_Name] LIKE '" & IC_Name & "%" & "' Order By [PGM_Name]")
    
    Do Until TesterRS.EOF = True
        If TesterRS.Fields(0) = 1 Then
            ChipNameCombo2.AddItem Mid(TesterRS.Fields(1), 7, Len(TesterRS.Fields(1)))
        End If
        TesterRS.MoveNext
    Loop
    
    TestTimeOut = 0
    
    If ChipNameCombo2.ListCount = 1 And ChipNameCombo2.List(0) = "" Then
        Call ChipNameCombo2_Click
        Exit Sub
    End If
    
    ChipNameCombo2.Enabled = True
    
End Sub

Private Sub Command1_Click()

Dim ReadStartSignal
Dim i As Byte
Dim result
Dim Channel
Dim x As Byte

    Call Timer_1ms(10)
    Debug.Print Timer

    x = &H0
    x = &HAA
    For i = 0 To 7
    
        Select Case i
                Case 0
                    Channel = Channel_P1A
                Case 1
                    Channel = Channel_P1B
                Case 2
                    Channel = Channel_P1C
                Case 3
                    Channel = Channel_P2A
                Case 4
                    Channel = Channel_P2B
                Case 5
                    Channel = Channel_P2C
                Case 6
                    Channel = Channel_P3A
                Case 7
                    Channel = Channel_P3B
            End Select
            
        Debug.Print i
        result = DO_WritePort(card, Channel, x)
    Next
    
    Debug.Print Timer
    Call Timer_1ms(1000)
    Debug.Print Timer
     
    result = DO_ReadPort(card, Channel_P3C, ReadStartSignal)
    Debug.Print "h"; Hex(ReadStartSignal)
    
    For i = 0 To 7
        Debug.Print Hex(CAndValue(ReadStartSignal, CPort(i)))
    Next

End Sub

Private Sub ChipNameCombo2_Click()
Dim DB_Path As String
Dim MPTesterRS As New ADODB.Recordset
Dim TesterRS As New ADODB.Recordset
Dim ConnMPTesterDB As New ADODB.Connection
Dim ConnTesterDB As New ADODB.Connection
Dim MpTesterDB_Path As String
Dim TesterDB_Path As String
Dim NameTmp As String
Dim DbPath As String
 
    If Trim(ChipNameCombo2.Text = "") Then
        ChipName = Trim(ChipNameCombo.Text)
    Else
        ChipName = Trim(ChipNameCombo.Text) & Trim(ChipNameCombo2.Text)
    End If

    HubNonUPT2Flag = False
    
    If (ChipName = "AU6350BL_1Port") Or _
       (ChipName = "AU6350GL_2Port") Or _
       (ChipName = "AU6350CF_3Port") Or _
       (ChipName = "AU6350AL_4Port") Or _
       (ChipName = "AU6256XLS1A") Then
        
        UPT2TestFlag = True
        
        If (ChipName = "AU6256XLS1A") Then
            HubNonUPT2Flag = True
        End If
        
    End If
    
    '=============================================
    '       connect to MPTester.mdb
    '=============================================
    MpTesterDB_Path = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\PGM_ListDB\MPTester_" & LastMPTesterDateCode & ".mdb"
    ConnMPTesterDB.Open MpTesterDB_Path
    MPTesterRS.CursorLocation = adUseClient
    MPTesterRS.Open "MPTester", ConnMPTesterDB, adOpenKeyset, adLockPessimistic
    
    NameTmp = ""
    MPTesterRS.MoveFirst
        
    '=============================================
    '           connect to Tester.mdb
    '=============================================
    TesterDB_Path = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\PGM_ListDB\Tester_" & LastTesterDateCode & ".mdb"
    ConnTesterDB.Open TesterDB_Path
    TesterRS.CursorLocation = adUseClient
    TesterRS.Open "Tester", ConnTesterDB, adOpenKeyset, adLockPessimistic
    
    MPTesterRS.Find "PGM_Name = '" & ChipName & "'"
        
    If MPTesterRS.EOF = True Then
        TesterRS.Find "PGM_Name = '" & ChipName & "'"
            
        If TesterRS.EOF = True Then
            MsgBox "*.HPC file error 2"
            End
        Else
            DB_Path = "Tester"
        End If
    Else
        DB_Path = "MPTester"
    End If
        
    If DB_Path = "MPTester" Then
        TestTimeOut = MPTesterRS.Fields(4)
        ResetPC.value = MPTesterRS.Fields(14)
        ReportCheck.value = MPTesterRS.Fields(15)
        
        ' for dual site
        WAIT_START_TIME_OUT = MPTesterRS.Fields(3)
        WAIT_TEST_CYCLE_OUT = MPTesterRS.Fields(4)
        POWER_ON_TIME = MPTesterRS.Fields(5)
        UNLOAD_DRIVER = MPTesterRS.Fields(7)
        CAPACTOR_CHARGE = MPTesterRS.Fields(9)
        NO_CARD_TEST_TIME = MPTesterRS.Fields(12)
        ReportCheck.value = MPTesterRS.Fields(15)
        Need_GPIB = MPTesterRS.Fields(17)
        
    End If
        
    If DB_Path = "Tester" Then
        TestTimeOut = TesterRS.Fields(4)
        ResetPC.value = TesterRS.Fields(14)
        ReportCheck.value = TesterRS.Fields(15)
        
        ' for dual site
        WAIT_START_TIME_OUT = TesterRS.Fields(3)
        WAIT_TEST_CYCLE_OUT = TesterRS.Fields(4)
        POWER_ON_TIME = TesterRS.Fields(5)
        UNLOAD_DRIVER = TesterRS.Fields(7)
        CAPACTOR_CHARGE = TesterRS.Fields(9)
        NO_CARD_TEST_TIME = TesterRS.Fields(12)
        ReportCheck.value = TesterRS.Fields(15)
        Need_GPIB = TesterRS.Fields(17)
        
    End If
    
    If SPIL_Flag Then
        If Dir(App.Path & "\" & ChipName) <> ChipName Then
            MsgBox ("Please Check Program Name !!")
            End
        End If
    End If
    
    'CycleTimeLbl.Caption = "MAX Cycle Time :" & CStr(TestTimeOut) & " S"
    BeginBtn.Enabled = True
    
    If (ChipName = "AU9540CSF21") Or (ChipName = "AU9540CSF20") Or (ChipName = "AU9525GLF20") Or Right(ChipName, 4) = "Port" Then
        VB6_Flag = False
    Else
        VB6_Flag = True
    End If

End Sub

Private Sub EndBtn_Click()
    End
End Sub

Private Sub ComCheck()
On Error Resume Next

Dim COM As Integer
Dim CheckCom As Integer

    '判斷 8 port卡是否存在，連續6個err number = 8002 即判斷無8埠卡
    For COM = 2 To 9
        If MSComm1(COM - 2).PortOpen = True Then
            MSComm1(COM - 2).PortOpen = False
        End If
        
        MSComm1(COM - 2).Settings = "9600,N,8,1"
        MSComm1(COM - 2).CommPort = COM
        MSComm1(COM - 2).PortOpen = True

        If err.Number = 8002 Then
            CheckCom = CheckCom + 1
        End If
        
        If CheckCom > 6 Then
            No8PCard = True
        End If
        MSComm1(COM - 2).PortOpen = False
    Next
End Sub

Private Sub Form_Load()
Dim BinLeft As Integer
Dim BinHeight As Integer
Dim BinTop As Integer
Dim BinWidth As Integer
Dim TitleWidth As Integer

'for DB
Dim MPTesterRS As New ADODB.Recordset
Dim TesterRS As New ADODB.Recordset
Dim ConnMPTesterDB As New ADODB.Connection
Dim ConnTesterDB As New ADODB.Connection
Dim MpTesterDB_Path As String
Dim TesterDB_Path As String
Dim NameTmp As String
Dim FS As New FileSystemObject
Dim FD As Folder
Dim ff As File
Dim TesterCounter As Integer
Dim MPTesterCounter As Integer
Dim VerifyCount As Integer
Dim CurrentName As String
Dim ComboExistFlag As Boolean

    If App.EXEName <> "Multi" Or CheckMe(Me) Then
        End
    End If

    ' for combine host
    First_site = False
    Second_site = False
    
    Call ComCheck               ' 判斷port數 (No8pcard = true 表示無8port卡)
   
    If No8PCard Then
        SiteCombo.AddItem "2"
        SiteCombo = 2
    Else
        SiteCombo.AddItem "2"
        SiteCombo.AddItem "4"
        SiteCombo.AddItem "6"
        SiteCombo.AddItem "8"
        SiteCombo = 4
    End If
    
    Call SiteCombo_Click        ' 依ComCheck決定Dual test是從哪個port輸出
    
    HubEnaOn = 0                ' initial HUB ENA pin control

    BinLeft = 1200
    BinHeight = 400
    BinTop = 1800
    BinWidth = 1400
    TitleWidth = 700
    SPIL_Flag = False

    If Dir("C:\Documents and Settings\User\桌面\SPIL.PC") = "SPIL.PC" Then
        SPIL_Flag = True
    End If

    ' set object position
    For i = 0 To 7
        SiteCheck(i).Left = BinLeft + i * (BinWidth + 100)
        SiteCheck(i).Height = BinHeight
        SiteCheck(i).Width = BinWidth
        SiteCheck(i).Top = BinTop
    
        TestResultLbl(i).Left = BinLeft + i * (BinWidth + 100)
        TestResultLbl(i).Height = BinHeight
        TestResultLbl(i).Width = BinWidth
        TestResultLbl(i).Top = BinTop + 500
        
        Status(i).Left = BinLeft + i * (BinWidth + 100)
        Status(i).Height = BinHeight
        Status(i).Width = BinWidth
        Status(i).Top = BinTop + 500 * 2

        Bin1(i).Left = BinLeft + i * (BinWidth + 100)
        Bin1(i).Height = BinHeight
        Bin1(i).Width = BinWidth
        Bin1(i).Top = BinTop + 500 * 3
    
        Bin1Title.Left = 100
        Bin1Title.Height = BinHeight
        Bin1Title.Top = Bin1(i).Top
        Bin1Title.Width = TitleWidth
    
        Bin2(i).Left = BinLeft + i * (BinWidth + 100)
        Bin2(i).Height = BinHeight
        Bin2(i).Width = BinWidth
        Bin2(i).Top = BinTop + 500 * 4
    
        Bin2Title.Left = 100
        Bin2Title.Height = BinHeight
        Bin2Title.Top = Bin2(i).Top
        Bin2Title.Width = TitleWidth
     
        Bin3(i).Left = BinLeft + i * (BinWidth + 100)
        Bin3(i).Height = BinHeight
        Bin3(i).Width = BinWidth
        Bin3(i).Top = BinTop + 500 * 5
    
        Bin3Title.Left = 100
        Bin3Title.Height = BinHeight
        Bin3Title.Top = Bin3(i).Top
        Bin3Title.Width = TitleWidth
    
        Bin4(i).Left = BinLeft + i * (BinWidth + 100)
        Bin4(i).Height = BinHeight
        Bin4(i).Width = BinWidth
        Bin4(i).Top = BinTop + 500 * 6
    
        Bin4Title.Left = 100
        Bin4Title.Height = BinHeight
        Bin4Title.Top = Bin4(i).Top
        Bin4Title.Width = TitleWidth
    
        Bin5(i).Left = BinLeft + i * (BinWidth + 100)
        Bin5(i).Height = BinHeight
        Bin5(i).Width = BinWidth
        Bin5(i).Top = BinTop + 500 * 7
    
        Bin5Title.Left = 100
        Bin5Title.Height = BinHeight
        Bin5Title.Top = Bin5(i).Top
        Bin5Title.Width = TitleWidth
    
        TimeOut(i).Left = BinLeft + i * (BinWidth + 100)
        TimeOut(i).Height = BinHeight
        TimeOut(i).Width = BinWidth
        TimeOut(i).Top = BinTop + 500 * 8
    
        TimeOutTitle.Left = 100
        TimeOutTitle.Height = BinHeight
        TimeOutTitle.Top = TimeOut(i).Top
        TimeOutTitle.Width = TitleWidth
    
        TotalFail(i).Left = BinLeft + i * (BinWidth + 100)
        TotalFail(i).Height = BinHeight
        TotalFail(i).Width = BinWidth
        TotalFail(i).Top = BinTop + 500 * 9
    
        TotalFailTitle.Left = 100
        TotalFailTitle.Height = BinHeight
        TotalFailTitle.Top = TotalFail(i).Top
        TotalFailTitle.Width = TitleWidth
    
        Yield(i).Left = BinLeft + i * (BinWidth + 100)
        Yield(i).Height = BinHeight
        Yield(i).Width = BinWidth
        Yield(i).Top = BinTop + 500 * 10
    
        YieldTitle.Left = 100
        YieldTitle.Height = BinHeight
        YieldTitle.Top = Yield(i).Top
        YieldTitle.Width = TitleWidth
        
        ContFail(i).Left = BinLeft + i * (BinWidth + 100)
        ContFail(i).Height = BinHeight
        ContFail(i).Width = BinWidth - 450
        ContFail(i).Top = BinTop + 500 * 11

        ContFailTitle.Left = 100
        ContFailTitle.Height = BinHeight
        ContFailTitle.Top = ContFail(i).Top
        ContFailTitle.Width = TitleWidth
        
    Next i

    ' set idle state
    For i = 0 To 7
        State(i) = IdleState
        Status(i).Caption = "FREE"
        SiteCheck(i).Caption = "Site" & CStr(i + 1)
    Next i

    CPort(0) = &H1
    CPort(1) = &H2
    CPort(2) = &H4
    CPort(3) = &H8
    CPort(4) = &H10
    CPort(5) = &H20
    CPort(6) = &H40
    CPort(7) = &H80

    ' read Program List from DB
    BeginBtn.Enabled = False
    
    ChipNameCombo.Clear
    ChipNameCombo.Text = "IC型號"
    ChipNameCombo2.Text = "程式版別"
    ChipNameCombo2.Enabled = False
    
    Set FD = FS.GetFolder(App.Path & "\PGM_ListDB\")
    
    If Not FS.FolderExists(App.Path & "\PGM_ListDB\Backup") Then
        FS.CreateFolder (App.Path & "\PGM_ListDB\Backup")
    End If
    
    MPTesterCounter = 0
    TesterCounter = 0
        
    For Each ff In FD.Files
                   
        If InStr(ff.Name, "_") > 0 Then
            If Len(ff.Name) = 21 And Left(ff.Name, 8) = "MPTester" Then     'MPTester
                ff.Copy App.Path & "\PGM_ListDB\Backup\MPTester.mdb"
                LastMPTesterDateCode = Right(ff.Name, 12)
                LastMPTesterDateCode = Left(LastMPTesterDateCode, 8)
                ff.Delete
            ElseIf Len(ff.Name) = 19 And Left(ff.Name, 6) = "Tester" Then   'Tester
                ff.Copy App.Path & "\PGM_ListDB\Backup\Tester.mdb"
                LastTesterDateCode = Right(ff.Name, 12)
                LastTesterDateCode = Left(LastTesterDateCode, 8)
                ff.Delete
            End If
        End If
    
    Next ff
    
    If LastMPTesterDateCode > LastTesterDateCode Then
        LastDateCode = LastMPTesterDateCode
    Else
        LastDateCode = LastTesterDateCode
    End If
    
    FS.CopyFile App.Path & "\PGM_ListDB\Backup\MPTester.mdb", App.Path & "\PGM_ListDB\MPTester_" & LastMPTesterDateCode & ".mdb"
    FS.CopyFile App.Path & "\PGM_ListDB\Backup\Tester.mdb", App.Path & "\PGM_ListDB\Tester_" & LastTesterDateCode & ".mdb"
            
    '=============================================
    '       connection to MPTester.mdb
    '=============================================
    MpTesterDB_Path = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\PGM_ListDB\MPTester_" & LastMPTesterDateCode & ".mdb"
    ConnMPTesterDB.Open MpTesterDB_Path
    MPTesterRS.CursorLocation = adUseClient
    MPTesterRS.Open "MPTester", ConnMPTesterDB, adOpenKeyset, adLockPessimistic
    Set MPTesterRS = ConnMPTesterDB.Execute("Select *  From [MPTester] Where [Visible] = 1 Order By [PGM_Name]")

    NameTmp = ""
    MPTesterRS.MoveFirst
    
    Do Until MPTesterRS.EOF
        CurrentName = Left(MPTesterRS.Fields(1), 6)
                
        If NameTmp <> CurrentName Then
            ComboExistFlag = False
            For VerifyCount = 0 To ChipNameCombo.ListCount
                If CurrentName = ChipNameCombo.List(VerifyCount) Then
                    ComboExistFlag = True
                    Exit For
                End If
            Next
            
            If Not ComboExistFlag Then
                ChipNameCombo.AddItem CurrentName
            End If
            NameTmp = CurrentName
        End If
                
        MPTesterRS.MoveNext
    Loop
    
    '=============================================
    '       connection to Tester.mdb
    '=============================================
    TesterDB_Path = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\PGM_ListDB\Tester_" & LastTesterDateCode & ".mdb"
    ConnTesterDB.Open TesterDB_Path
    TesterRS.CursorLocation = adUseClient
    TesterRS.Open "Tester", ConnTesterDB, adOpenKeyset, adLockPessimistic
    Set TesterRS = ConnTesterDB.Execute("Select *  From [Tester] Where [Visible] = 1 Order By [PGM_Name]")

    TesterRS.MoveFirst
    
    Do Until TesterRS.EOF
        CurrentName = Left(TesterRS.Fields(1), 6)
        
        If NameTmp <> CurrentName Then
            ComboExistFlag = False
            For VerifyCount = 0 To ChipNameCombo.ListCount
                If CurrentName = ChipNameCombo.List(VerifyCount) Then
                    ComboExistFlag = True
                    Exit For
                End If
            Next
            
            If Not ComboExistFlag Then
                ChipNameCombo.AddItem CurrentName
            End If
            NameTmp = CurrentName
        End If

        TesterRS.MoveNext
    Loop
    
    
    ConnMPTesterDB.Close
    Set ConnMPTesterDB = Nothing
    
    ConnTesterDB.Close
    Set ConnTesterDB = Nothing
    
    Host.Caption = Host.Caption & LastDateCode
   
    PreviousChipName = ""

    '============= FOR PCI_7296 ==============
    If AllenDebug = 0 Then
        card = Register_Card(PCI_7296, 0)
        Call SetTimer_1ms
        If card < 0 Then                        '若無7296才判斷是否有7248
            ' ========== For 7248 ===========
            card = Register_Card(PCI_7248, 0)
            Call SetTimer_1ms
            If card = 0 Then
               Initial_7248
               Exit Sub
            End If
             ' ==============================
            MsgBox "Register Card Failed"
            End
        End If
        Initial_7296
    End If

End Sub

Private Sub handler_NS_Click()
    Blank_Timer.Enabled = False
    Handler_Blank.BackColor = System_Color
    Handler_NS.BackColor = System_Color
    Handler_SRM.BackColor = System_Color
End Sub

Private Sub handler_SRM_Click()
    Blank_Timer.Enabled = False
    Handler_Blank.BackColor = System_Color
    Handler_NS.BackColor = System_Color
    Handler_SRM.BackColor = System_Color
End Sub

Private Sub OffLineCheck_Click()
    If OffLineCheck.value = 0 Then
        TestMode = 0
    Else
        TestMode = 1
    End If
End Sub

Private Sub Option1_Click()
    First_site = True
    Second_site = True
    SiteCheck(0).value = 1
    SiteCheck(1).value = 1
    SiteCheck(0).Enabled = False
    SiteCheck(1).Enabled = False
End Sub

Private Sub Option2_Click()
    First_site = True
    Second_site = False
    SiteCheck(0).value = 1
    SiteCheck(1).value = 0
    SiteCheck(0).Enabled = False
    SiteCheck(1).Enabled = False
End Sub

Private Sub Option3_Click()
    First_site = False
    Second_site = True
    SiteCheck(0).value = 0
    SiteCheck(1).value = 1
    SiteCheck(0).Enabled = False
    SiteCheck(1).Enabled = False
End Sub

Private Sub SiteCombo_Click()

On Error Resume Next

    ' 1. 接在主機版上
    If SiteCombo.Text = "2" And No8PCard = True Then
        ' 可用 com1 or com2
        Mode.Visible = True
        
        If Option1.value = False And Option2.value = False And Option3.value = False Then
            Option1.value = True
        End If
        
        If No8PCard Then
        
            MSComm2.CommPort = 1
            MSComm2.Settings = "9600,N,8,1"
            If MSComm2.PortOpen = False Then
                MSComm2.PortOpen = True
            End If
            MSComm2.InBufferCount = 0
            MSComm2.InputLen = 0
            
            MSComm3.CommPort = 2
            MSComm3.Settings = "9600,N,8,1"
            If MSComm3.PortOpen = False Then
                MSComm3.PortOpen = True
            End If
            MSComm3.InBufferCount = 0
            MSComm3.InputLen = 0

        End If

        For i = 2 To 7
            SiteCheck(i).Enabled = False
        Next
        
        Exit Sub
    ' 2. 接在8 port卡上
    Else
        MSComm2.CommPort = 1
        If MSComm2.PortOpen Then
            MSComm2.PortOpen = False
        End If
        
        MSComm3.CommPort = 2
        If MSComm3.PortOpen Then
            MSComm3.PortOpen = False
        End If
        
        Mode.Visible = False
        First_site = False
        Second_site = False
    End If
    
    For i = 0 To 7
        MSComm1(i).CommPort = i + 2
        MSComm1(i).Settings = "9600,N,8,1"
        
        If MSComm1(i).PortOpen = False Then
            MSComm1(i).PortOpen = True
        End If
        
        MSComm1(i).InBufferCount = 0
        MSComm1(i).InputLen = 0
    Next
    
End Sub

Private Sub StopBtn_Click()
On Error Resume Next

    If (MsgBox("Stop Test ?", vbYesNo + vbQuestion + vbDefaultButton2, "Comform Stop") = vbYes) Then
        AllenStop = 1
        Exit Sub
    End If
    StopFlag = 1

End Sub
Sub LockOption()
    ChipNameCombo.Enabled = False
    ChipNameCombo2.Enabled = False
    SiteCombo.Enabled = False
    OffLineCheck.Enabled = False
    ResetPC.Enabled = False
    OneCycleCheck.Enabled = False
    ReportCheck.Enabled = False
    BeginBtn.Enabled = False
    EndBtn.Enabled = False
    Handler_NS.Enabled = False
    Handler_SRM.Enabled = False
    Option1.Enabled = False
    Option2.Enabled = False
    Option3.Enabled = False
End Sub
Sub UnlockOption()
    ChipNameCombo.Enabled = True
    ChipNameCombo2.Enabled = True
    SiteCombo.Enabled = True
    OffLineCheck.Enabled = True
    ResetPC.Enabled = True
    OneCycleCheck.Enabled = True
    ReportCheck.Enabled = True
    BeginBtn.Enabled = True
    EndBtn.Enabled = True
    Handler_NS.Enabled = True
    Handler_SRM.Enabled = True
    Option1.Enabled = True
    Option2.Enabled = True
    Option3.Enabled = True
End Sub

Sub DualTestResult()
Dim k As Integer
Dim tmp As Single
Dim tmpFail As Long
Dim tmps As String

Dim CurBinning As String
Dim v As Integer
Dim CurOutputStr1 As String
Dim CurOutputStr2 As String

    k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off"& " SetSITE2 CDN high" Channel_P2A = 255
    k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off"& " SetSITE1 CDN high" Channel_P2A = 255
    
    If First_site = True Then
        Select Case TestResult1
            Case "PASS"
                TestResult1 = "PASS"
                
                TestResultLbl(0).Caption = "PASS"
                TestResultLbl(0).BackColor = GREEN_COLOR
                Bin1Counter(0) = Bin1Counter(0) + 1
                Bin1(0).Caption = CStr(Bin1Counter(0))
                Bin2ContiFail(0) = 0
                ContiFailCounter(0) = 0
        
                CurBinning = 1
        
            Case "UNKNOW", "bin2", "Bin2"
                TestResult1 = "bin2"
                
                TestResultLbl(0).Caption = "bin2"
                TestResultLbl(0).BackColor = RED_COLOR
                Bin2Counter(0) = Bin2Counter(0) + 1
                Bin2(0).Caption = CStr(Bin2Counter(0))
                Bin2ContiFail(0) = Bin2ContiFail(0) + 1
                ContiFailCounter(0) = ContiFailCounter(0) + 1
                
                If Bin2ContiFail(0) > 4 And ResetPC.value = 1 Then
                    result = DO_WritePort(card, Channel, PCI7296_PC_RESET)
                    Call MsecDelay(PC_RESET_TIME)
                    result = DO_WritePort(card, Channel, &HFF)
                    Bin2ContiFail(0) = 0
                End If
                
                CurBinning = 2
                
            Case "gponFail", "bin3", "Bin3", "SD_WF", "SD_RF", "CF_WF", "CF_RF"
                TestResult1 = "bin3"
                
                TestResultLbl(0).Caption = "bin3"
                TestResultLbl(0).BackColor = YELLOW_COLOR
                Bin3Counter(0) = Bin3Counter(0) + 1
                ContiFailCounter(0) = ContiFailCounter(0) + 1
                Bin3(0).Caption = CStr(Bin3Counter(0))
                
                CurBinning = 3
                
            Case "XD_WF", "bin4", "Bin4", "XD_RF"
                TestResult1 = "bin4"

                TestResultLbl(0).Caption = "bin4"
                TestResultLbl(0).BackColor = YELLOW_COLOR
                Bin4Counter(0) = Bin4Counter(0) + 1
                ContiFailCounter(0) = ContiFailCounter(0) + 1
                Bin4(0).Caption = CStr(Bin4Counter(0))
                
                CurBinning = 4

            Case "MS_WF", "bin5", "TimeOut", "Bin5", "MS_RF"
                TestResult1 = "bin5"

                TestResultLbl(0).Caption = "bin5"
                TestResultLbl(0).BackColor = YELLOW_COLOR
                Bin5Counter(0) = Bin5Counter(0) + 1
                ContiFailCounter(0) = ContiFailCounter(0) + 1
                Bin5(0).Caption = CStr(Bin5Counter(0))
                
                CurBinning = 5

            Case Else
                TestResult1 = "bin2"
                
                TestResultLbl(0).Caption = "bin2"
                TestResultLbl(0).BackColor = RED_COLOR
                Bin2Counter(0) = Bin2Counter(0) + 1
                Bin2(0).Caption = CStr(Bin2Counter(0))
                ContiFailCounter(0) = ContiFailCounter(0) + 1
                
                CurBinning = 2
                
        End Select
        
        TempNowStr(0) = CurBinning & TempNowStr(0)
        TempNowStr(0) = Left(TempNowStr(0), 20)
        
        CurOutputStr1 = ""
        
        For v = Len(TempNowStr(0)) To 1 Step -1
            If v Mod 10 = 0 Then
                CurOutputStr1 = Mid(TempNowStr(0), v, 1) & vbCrLf & CurOutputStr1
            Else
                CurOutputStr1 = Mid(TempNowStr(0), v, 1) & CurOutputStr1
            End If
        Next
        
        ContFail(0).Caption = CurOutputStr1
        
        tmpFail = (Bin2Counter(0) + Bin3Counter(0) + Bin4Counter(0) + Bin5Counter(0))
        
        TotalFail(0).Caption = CStr(tmpFail)
        tmp = CSng(Bin1Counter(0) / (tmpFail + Bin1Counter(0)))
        tmps = Format$(tmp, "0.00%")
        Yield(0).Caption = tmps
        
    End If

    If Second_site = True Then
        Select Case TestResult2
            Case "PASS"
                TestResult2 = "PASS"
                
                TestResultLbl(1).Caption = "PASS"
                TestResultLbl(1).BackColor = GREEN_COLOR
                Bin1Counter(1) = Bin1Counter(1) + 1
                Bin1(1).Caption = CStr(Bin1Counter(1))
                Bin2ContiFail(1) = 0
                ContiFailCounter(1) = 0
                
                CurBinning = 1
        
            Case "UNKNOW", "bin2", "Bin2"
                TestResult2 = "bin2"
                
                TestResultLbl(1).Caption = "bin2"
                TestResultLbl(1).BackColor = RED_COLOR
                Bin2Counter(1) = Bin2Counter(1) + 1
                Bin2(1).Caption = CStr(Bin2Counter(1))
                Bin2ContiFail(1) = Bin2ContiFail(1) + 1
                ContiFailCounter(1) = ContiFailCounter(1) + 1
                
                If Bin2ContiFail(1) > 4 And ResetPC.value = 1 Then
                    result = DO_WritePort(card, Channel, PCI7296_PC_RESET)
                    Call MsecDelay(PC_RESET_TIME)
                    result = DO_WritePort(card, Channel, &HFF)
                    Bin2ContiFail(1) = 0
                End If
                
                CurBinning = 2
                
            Case "gponFail", "bin3", "Bin3", "SD_WF", "SD_RF", "CF_WF", "CF_RF"
                TestResult2 = "bin3"
                
                TestResultLbl(1).Caption = "bin3"
                TestResultLbl(1).BackColor = YELLOW_COLOR
                Bin3Counter(1) = Bin3Counter(1) + 1
                ContiFailCounter(1) = ContiFailCounter(1) + 1
                Bin3(1).Caption = CStr(Bin3Counter(1))
                
                CurBinning = 3
                
            Case "XD_WF", "bin4", "Bin4", "XD_RF"
                TestResult2 = "bin4"

                TestResultLbl(1).Caption = "bin4"
                TestResultLbl(1).BackColor = YELLOW_COLOR
                Bin4Counter(1) = Bin4Counter(1) + 1
                ContiFailCounter(1) = ContiFailCounter(1) + 1
                Bin4(1).Caption = CStr(Bin4Counter(1))
                
                CurBinning = 4

            Case "MS_WF", "bin5", "TimeOut", "Bin5", "MS_RF"
                TestResult2 = "bin5"

                TestResultLbl(1).Caption = "bin5"
                TestResultLbl(1).BackColor = YELLOW_COLOR
                Bin5Counter(1) = Bin5Counter(1) + 1
                ContiFailCounter(1) = ContiFailCounter(1) + 1
                Bin5(1).Caption = CStr(Bin5Counter(1))
                
                CurBinning = 5

            Case Else
                TestResult2 = "bin2"
                
                TestResultLbl(1).Caption = "bin2"
                TestResultLbl(1).BackColor = RED_COLOR
                Bin2Counter(1) = Bin2Counter(1) + 1
                Bin2(1).Caption = CStr(Bin2Counter(1))
                ContiFailCounter(1) = ContiFailCounter(1) + 1
                
                CurBinning = 2

        End Select
        
        TempNowStr(1) = CurBinning & TempNowStr(1)
        TempNowStr(1) = Left(TempNowStr(1), 20)
        
        CurOutputStr2 = ""
        
        For v = Len(TempNowStr(1)) To 1 Step -1
            If v Mod 10 = 0 Then
                CurOutputStr2 = Mid(TempNowStr(1), v, 1) & vbCrLf & CurOutputStr2
            Else
                CurOutputStr2 = Mid(TempNowStr(1), v, 1) & CurOutputStr2
            End If
        Next
        
        ContFail(1).Caption = CurOutputStr2
        
        tmpFail = (Bin2Counter(1) + Bin3Counter(1) + Bin4Counter(1) + Bin5Counter(1))
        
        TotalFail(1).Caption = CStr(tmpFail)
        tmp = CSng(Bin1Counter(1) / (tmpFail + Bin1Counter(1)))
        tmps = Format$(tmp, "0.00%")
        Yield(1).Caption = tmps
        
    End If
    
    If GreaTekChipName = "AU6368A1" Then  ' For GTK AU6368A1 sorting case
        If TestResult1 = "bin2" Then TestResult1 = "bin4"
        If TestResult1 = "bin3" Then TestResult1 = "bin5"
        If TestResult2 = "bin2" Then TestResult2 = "bin4"
        If TestResult2 = "bin3" Then TestResult2 = "bin5"
    End If

    'Binning
    Select Case TestResult1
        Case "PASS"
            If TestMode = 0 Then
                PassCounter1 = PassCounter1 + 1
            Else
                OffLPassCounter1 = OffLPassCounter1 + 1
            End If
            
            Call PCI7248_bin(Channel_P2B, PCI7248_PASS)
        Case "bin2"
'            If OffLineCheck.value = 0 Then
'                Bin2Counter(0) = Bin2Counter(0) + 1
'            End If

            Call PCI7248_bin(Channel_P2B, PCI7248_BIN2)
        Case "bin3"
'            If OffLineCheck.value = 0 Then
'                Bin3Counter(0) = Bin3Counter(0) + 1
'            End If

            Call PCI7248_bin(Channel_P2B, PCI7248_BIN3)
        Case "bin4"
'            If OffLineCheck.value = 0 Then
'                Bin4Counter(0) = Bin4Counter(0) + 1
'            End If

            Call PCI7248_bin(Channel_P2B, PCI7248_BIN4)
        Case "bin5"
'            If OffLineCheck.value = 0 Then
'                Bin5Counter(0) = Bin5Counter(0) + 1
'            End If

            Call PCI7248_bin(Channel_P2B, PCI7248_BIN5)

        Case Else
'            If OffLineCheck.value = 0 Then
'                Bin2Counter(0) = Bin2Counter(0) + 1
'            End If

            Call PCI7248_bin(Channel_P2B, PCI7248_BIN2)

    End Select

    Select Case TestResult2
        Case "PASS"
            If TestMode = 0 Then
                PassCounter2 = PassCounter2 + 1
            Else
                OffLPassCounter2 = OffLPassCounter2 + 1
            End If
            
            Call PCI7248_bin(Channel_P2A, PCI7248_PASS)
            
        Case "bin2"
'            If OffLineCheck.value = 0 Then
'                Bin2Counter(1) = Bin2Counter(1) + 1
'            End If

            Call PCI7248_bin(Channel_P2A, PCI7248_BIN2)
        Case "bin3"
'            If OffLineCheck.value = 0 Then
'               Bin3Counter(1) = Bin3Counter(1) + 1
'            End If

            Call PCI7248_bin(Channel_P2A, PCI7248_BIN3)
        Case "bin4"
'            If OffLineCheck.value = 0 Then
'                Bin4Counter(1) = Bin4Counter(1) + 1
'            End If

            Call PCI7248_bin(Channel_P2A, PCI7248_BIN4)
        Case "bin5"
'            If OffLineCheck.value = 0 Then
'                Bin5Counter(1) = Bin5Counter(1) + 1
'            End If

            Call PCI7248_bin(Channel_P2A, PCI7248_BIN5)

        Case Else
'            If OffLineCheck.value = 0 Then
'                Bin2Counter(1) = Bin2Counter(1) + 1
'            End If

            Call PCI7248_bin(Channel_P2A, PCI7248_BIN2)

    End Select
    
    
    If TestResult1 = "PASS" Then
        continuefail1 = 0
    ElseIf TestResult1 = "bin2" Then
        continuefail1 = continuefail1 + 1
    ElseIf TestResult1 = "bin3" Then
        continuefail1 = continuefail1 + 1
    ElseIf TestResult1 = "bin4" Then
        continuefail1 = continuefail1 + 1
    ElseIf TestResult1 = "bin5" Then
        continuefail1 = continuefail1 + 1
    End If
    
    If TestResult2 = "PASS" Then
        continuefail2 = 0
    ElseIf TestResult2 = "bin2" Then
        continuefail2 = continuefail2 + 1
    ElseIf TestResult2 = "bin3" Then
        continuefail2 = continuefail2 + 1
    ElseIf TestResult2 = "bin4" Then
        continuefail2 = continuefail2 + 1
    ElseIf TestResult2 = "bin5" Then
        continuefail2 = continuefail2 + 1
    End If
    
    If ((continuefail1 >= 5) Or (continuefail2 >= 5)) And (InStr(ChipName, "U69") = 2) And (Len(ChipName) = 14) And (Mid(ChipName, 12, 1) <> "U") Then
        SendMP_Flag = True
        If (RealChipName = "") And (MPChipName = "") Then
            RealChipName = Trim(ChipNameCombo.Text) & Trim(ChipNameCombo2.Text)
            MPChipName = Left(ChipName, 10) & "M" & Right(ChipName, 3)
        End If
        ChipName = MPChipName
    End If
    
    If (continuefail1 = 0) And (continuefail2 = 0) And (InStr(ChipName, "U69") = 2) And (Mid(ChipName, 12, 1) <> "U") Then
        SendMP_Flag = False
        ChipName = Trim(ChipNameCombo.Text) & Trim(ChipNameCombo2.Text)
    End If
    
    Call UpdateDB
    
    If (Len(ChipName) = 14) Then
        If SendMP_Flag = True Then
           ChipName = MPChipName
        End If
    End If
    
    ' for PrintReport use
    Bin1Site1 = Bin1Counter(0)
    Bin2Site1 = Bin2Counter(0)
    Bin3Site1 = Bin3Counter(0)
    Bin4Site1 = Bin4Counter(0)
    Bin5Site1 = Bin5Counter(0)
    
    Bin1Site2 = Bin1Counter(1)
    Bin2Site2 = Bin2Counter(1)
    Bin3Site2 = Bin3Counter(1)
    Bin4Site2 = Bin4Counter(1)
    Bin5Site2 = Bin5Counter(1)
    
End Sub

Sub PCI7248_bin(Channel As Byte, PCI7248bin As Byte)
Dim k As Integer
Dim DO_P As Long

        DO_P = PCI7248bin
        k = DO_WritePort(card, Channel, DO_P)
        Call Timer_1ms(7)
    
        DO_P = PCI7248bin - PCI7248_EOT
        k = DO_WritePort(card, Channel, DO_P)
        Call Timer_1ms(7)
        
        DO_P = PCI7248bin
        k = DO_WritePort(card, Channel, DO_P)
        Call Timer_1ms(7)
        
        DO_P = &HFF
        k = DO_WritePort(card, Channel, DO_P)

End Sub

Sub DualCommonTest()
Dim k As Integer

    ''''''''''''''''''''''''''''''''
    ' wait Start Signal From Handle
    ''''''''''''''''''''''''''''''''
    GetGPIBStatus(0) = False
    GetGPIBStatus(1) = False
    TestStop1 = 0
    TestStop2 = 0
    
    WaitForStart = Timer    'Get Vcc on from Chip
    
    If TestMode = 0 Then    'ON LINE MODE    'Check1.Value = 0 => TestMode = 0  '上線模式
        Print "wait Start"
        Do                  ' wait  (VCC PowerON) & (handler 5ms start) signal
            DoEvents
            WaitStartTime = Timer - WaitForStart
            k = DO_ReadPort(card, Channel_P2CH, DI_P)
        Loop Until DI_P = 14 Or DI_P = 13 Or DI_P = 12 Or AllenStop = 1
    
        TotalRealTestTime = Timer - OldTotalRealTestTime
        OldTotalRealTestTime = Timer
        OldRealTestTime = Timer
    
    Else                    ' Check1.Value = 1 => TestMode = 1   '離線模式
    
        Call MsecDelay(0.2)
        WaitStartTime = 0.2
        DI_P = 14
        TotalRealTestTime = Timer - OldTotalRealTestTime
        OldTotalRealTestTime = Timer
        OldRealTestTime = Timer
    
    End If
    
    buf1 = MSComm2.Input
    WaitStartCounter = WaitStartCounter + 1
    
    If WaitStartTime > WAIT_START_TIME_OUT Then
        WaitStartTimeOut = 1
        WaitStartTimeOutCounter = WaitStartTimeOutCounter + 1
    End If
    
    
    '''''''''''''''''''''''''''''
    ' 5 fail reset
    '''''''''''''''''''''''''''''
    If ResetPC.value = 1 Then
        If continuefail1 >= AlarmLimit Or continuefail2 >= AlarmLimit Then
            
            continuefail1 = 0
            continuefail2 = 0
    
            If Left(ChipName, 10) = "AU6254XLS4" Then
                k = DO_WritePort(card, Channel_P1CH, &H0)
                Call MsecDelay(6)
        
                k = DO_WritePort(card, Channel_P1CH, &HF)
                Call MsecDelay(2)
            End If
    
            k = DO_WritePort(card, Channel_P1B, &HF)    ' send 0000,1111 => RESET PC " Channel_P1B= 15
            Call MsecDelay(2)
            k = DO_WritePort(card, Channel_P1B, &HFF)   ' send 1111,1111 => RESET PC " Channel_P1B= 255
        End If
    End If
    
    ''''''''''''''''''''''''''''''''
    '   Open Power
    ''''''''''''''''''''''''''''''''
    If DI_P < 12 And DI_P >= 15 Then                    'Allen 20050607 , change DI_P > 15, to DI_P >= 15
        Print "no start"
        Exit Sub
    Else
        
        Host.Cls
        Print "get start signal!"
        
        Call MsecDelay(CAPACTOR_CHARGE)
        Call MsecDelay(UNLOAD_DRIVER)
        
        k = DO_WritePort(card, Channel_P2A, &H7F)       ' send 0111,1111 => Set power" Channel_P2A = 127
        k = DO_WritePort(card, Channel_P2B, &H7F)       ' send 0111,1111 => Set power" Channel_P2b = 127
        NewPowerOnTime = POWER_ON_TIME - 0.4
        
        If NewPowerOnTime > 0 Then
            Call MsecDelay(NewPowerOnTime)
        End If
    
    End If
    
    '''''''''''''''''''''''''''''''''''''
    '   Wait Tester Signal
    '''''''''''''''''''''''''''''''''''''
    
LoopTest_Start:
    
    MSComm2.InBufferCount = 0
    MSComm3.InBufferCount = 0
    buf1 = ""
    buf2 = ""
    TesterStatus1 = ""
    TesterStatus2 = ""
    TesterReady1 = 0
    TesterReady2 = 0
    ResetCounter1 = 0
    ResetCounter2 = 0
    TesterDownCount1 = 0
    TesterDownCount2 = 0
    WaitForReady = Timer
    
    Do
        DoEvents
        
        '========================
        ' wait for tester1 ready
        '========================
        
        If SiteCheck(0).value = 1 Then
            If TesterReady1 = 0 Then
                If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                    MSComm2.Output = "~"
                    If HubEnaOn = 0 Then
                        k = DO_WritePort(card, Channel_P1CH, &H0) ' set HUB module Ena on
                        HubEnaOn = 1
                    End If
                End If
    
                buf1 = MSComm2.Input
                TesterStatus1 = TesterStatus1 & buf1
                
                If (InStr(1, TesterStatus1, "Ready") <> 0) Then
                    TesterReady1 = 1
                End If
    
            End If
        Else
            TesterReady1 = 1
        End If
    
        '========================
        ' wait for tester2 ready
        '========================
        
        If SiteCheck(1).value = 1 Then
            If TesterReady2 = 0 Then
                If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                    MSComm3.Output = "~"
            
                    If HubEnaOn = 0 Then
                        k = DO_WritePort(card, Channel_P1CH, &H0) ' set HUB module Ena on
                        HubEnaOn = 1
                    End If
        
                End If
        
                buf2 = MSComm3.Input
                TesterStatus2 = TesterStatus2 & buf2
                
                If (InStr(1, TesterStatus2, "Ready") <> 0) Then
                    TesterReady2 = 1
                End If
            
            End If
        Else
            TesterReady2 = 1
        End If
        
        '===================================
        ' Reset rountine : consider Reset fail
        '===================================
        
        If Timer - WaitForReady > 1 Then
            If ResetCounter1 > 2 Or ResetCounter2 > 2 Then  ' Alarm for reset fail
                Call PrintReport  'print routine
                Print "Alarm : Reset PC fail"
                MsgBox "Alarm : Reset PC fail "
                ResetCounter1 = 0
                ResetCounter2 = 0
                Exit Sub
            Else
                ' Reset  Rountine
                If TesterReady1 = 0 And TesterDownCount1 = 0 And FirstRun = 1 Then ' reset tester1
                
                    ' close module power
                    ResetCounter1 = ResetCounter1 + 1
                    
                    k = DO_WritePort(card, Channel_P2A, &HFF)   ' send 1111,1111 => Set power off" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &HFF)   ' send 1111,1111 => Set power off" Channel_P2A = 255
                    
                    ' Reset PC
                    TesterDownCount1 = 1
            
                    If Left(ChipName, 10) = "AU6254XLS4" Then
                        k = DO_WritePort(card, Channel_P1CH, &H0)
                        Call MsecDelay(6)
                    
                        k = DO_WritePort(card, Channel_P1CH, &HF)
                        Call MsecDelay(2)
                    End If
            
                    k = DO_WritePort(card, Channel_P1B, &HF)    ' send 0000,1111 => RESET PC " Channel_P1B= 15
                    Call MsecDelay(2)
                    k = DO_WritePort(card, Channel_P1B, &HFF)   ' send 1111,1111 => RESET PC " Channel_P1B= 255
                    WaitForPowerOn1 = Timer
            
                    ' clear comm buffer
                    MSComm2.InBufferCount = 0
                    TesterStatus1 = ""
                End If
        
                If TesterReady2 = 0 And TesterDownCount2 = 0 And FirstRun = 1 Then ' reset tester2
                
                    ' close module power
                    ResetCounter2 = ResetCounter2 + 1
                    
                    k = DO_WritePort(card, Channel_P2A, &HFF)   ' send 1111,1111 => Set power off" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &HFF)   ' send 1111,1111 => Set power off" Channel_P2A = 255
                    
                    ' Reset PC
                    TesterDownCount2 = 1
                
                    If Left(ChipName, 10) = "AU6254XLS4" Then
                        k = DO_WritePort(card, Channel_P1CH, &H0)
                        Call MsecDelay(6)
                        
                        k = DO_WritePort(card, Channel_P1CH, &HF)
                        Call MsecDelay(2)
                    End If
            
                    k = DO_WritePort(card, Channel_P1B, &HF)    ' send 0000,1111 => RESET PC " Channel_P1B= 15
                    Call MsecDelay(2)
                    k = DO_WritePort(card, Channel_P1B, &HFF)   ' send 1111,1111 => RESET PC " Channel_P1B= 255
                    WaitForPowerOn2 = Timer
                    
                    ' clear comm buffer
                    MSComm3.InBufferCount = 0
                    TesterStatus2 = ""
                End If
            End If
        
            '===============================
            ' screen down count routine
            '==============================
            
            If TesterDownCount1 = 1 Then
            
                TesterDownCountTimer1 = Timer - WaitForPowerOn1
                
                If TesterReady1 = 1 Then
                
                    ' open module power
                    k = DO_WritePort(card, Channel_P2A, &H7F)   ' send 0111,1111 => Set power on" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &H7F)   ' send 0111,1111 => Set power on" Channel_P2A = 255
                    
                    Call MsecDelay(POWER_ON_TIME)
                    
                    ' clear flag
                    TesterDownCount1 = 0
                End If
                
                If TesterDownCountTimer1 > 90 Then  'Reset fail
                    TesterDownCount1 = 0
                End If
            
            End If
            
            If TesterDownCount2 = 1 Then
            
                TesterDownCountTimer2 = Timer - WaitForPowerOn2
                
                If TesterReady2 = 1 Then
                
                    ' open module power
                    k = DO_WritePort(card, Channel_P2A, &H7F)   ' send 0111,1111 => Set power on" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &H7F)   ' send 0111,1111 => Set power on" Channel_P2A = 255
                    Call MsecDelay(POWER_ON_TIME)
                    
                    ' clear flag
                    TesterDownCount2 = 0
                End If
                
                If TesterDownCountTimer2 > 90 Then              ' Reset fail
                    TesterDownCount2 = 0
                End If
            
            End If
        End If
           
    Loop Until (TesterReady1 = 1) And (TesterReady2 = 1) Or AllenStop = 1
    
    FirstRun = 1

    ''''''''''''''''''''''''''
    '   Testing Loop
    ''''''''''''''''''''''''''
    
    If (DI_P >= 12) And (DI_P < 15) Then

        k = DO_WritePort(card, Channel_P2A, &H3F)       ' send 0011,1111 => Set power on"& " SetSITE2 CDN Low" Channel_P2A = 63
        k = DO_WritePort(card, Channel_P2B, &H3F)       ' send 0011,1111 => Set power on"& " SetSITE1 CDN Low" Channel_P2A = 63

        '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '*STEP4=> Waitting for Response from  Tester
        '*
        '*    Wait Test Result from each Tester
        '*
        '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

        NoCardTestResult1 = ""
        NoCardTestResult2 = ""
        
        TesterStatus1 = ""
        TesterStatus2 = ""

        Do
            DoEvents
            If SiteCheck(0).value = 1 Then
                If NoCardTestStop1 = 0 Then
                
                    If MSComm2.InBufferCount >= 4 Then
                        NoCardTestResult1 = MSComm2.Input
                    End If
                                       
                    NoCardTestResult1 = Parser(NoCardTestResult1)
                  
                    NoCardTestCycleTime1 = Timer - NoCardWaitForTest1
                    
                    If (NoCardTestResult1 <> "" Or NoCardTestCycleTime1 > NO_CARD_TEST_TIME) Then
                       NoCardTestStop1 = 1
                    End If
                End If
            
            Else
                NoCardTestStop1 = 1
            End If

            If SiteCheck(1).value = 1 Then
                If NoCardTestStop2 = 0 Then
                     If MSComm3.InBufferCount >= 4 Then
                            NoCardTestResult2 = MSComm3.Input
                     End If
                     
                     NoCardTestCycleTime2 = Timer - NoCardWaitForTest2
                    
                     NoCardTestResult2 = Parser(NoCardTestResult2)
                     
                    If (NoCardTestResult2 <> "" Or NoCardTestCycleTime2 > NO_CARD_TEST_TIME) Then
                        NoCardTestStop2 = 1
                        
                    End If
                End If
            
            Else
                NoCardTestStop2 = 1
            End If
          
        Loop Until (NoCardTestStop1 = 1) And (NoCardTestStop2 = 1) Or AllenStop = 1
            
        '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '*STEP3=>Send command to PC tester
        '*
        '*    Send ChipName to PC tester
        '*
        '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        
        TestResult1 = ""
        TestResult2 = ""

        If SiteCheck(0).value = 1 Then
             NoCardTestResult1 = "PASS"
            
                 If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                     For i = 1 To Len(ChipName)
                         MSComm2.Output = Mid(ChipName, i, 1)
                         Call MsecDelay(0.02)
                     Next
                 Else
                     MSComm2.Output = ChipName   ' trans strat test signal to TEST PC
                 End If
            
             MSComm2.InBufferCount = 0
             MSComm2.InputLen = 0
             WaitForTest1 = Timer
        End If
        
        If SiteCheck(1).value = 1 Then
            NoCardTestResult2 = "PASS"
            
                If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                    For i = 1 To Len(ChipName)
                        MSComm3.Output = Mid(ChipName, i, 1)
                        Call MsecDelay(0.02)
                    Next
                Else
                    MSComm3.Output = ChipName   ' trans strat test signal to TEST PC
                End If
            
            MSComm3.InBufferCount = 0
            MSComm3.InputLen = 0
            WaitForTest2 = Timer
        End If
   
        Do
            DoEvents
            If SiteCheck(0).value = 1 And NoCardTestResult1 = "PASS" Then
                If TestStop1 = 0 Then
                
                    If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                        MSComm2.Output = "~"
                    End If
                    
                    If MSComm2.InBufferCount >= 4 Then
                        TestResult1 = MSComm2.Input
                    End If
                    
                    TestResult1 = Parser(TestResult1)
                    TestCycleTime1 = Timer - WaitForTest1
                    
                    If (TestResult1 <> "" Or TestCycleTime1 > WAIT_TEST_CYCLE_OUT) Then
                        TestStop1 = 1
                    End If
                End If
                
            Else
                TestStop1 = 1
            End If

            If SiteCheck(1).value = 1 And NoCardTestResult2 = "PASS" Then
                If TestStop2 = 0 Then
                    
                    If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                        MSComm3.Output = "~"
                    End If
                    
                    If MSComm3.InBufferCount >= 4 Then
                        TestResult2 = MSComm3.Input
                    End If
                    
                    TestResult2 = Parser(TestResult2)
                    TestCycleTime2 = Timer - WaitForTest2
                    
                    If (TestResult2 <> "" Or TestCycleTime2 > WAIT_TEST_CYCLE_OUT) Then
                        TestStop2 = 1
                    End If
                End If
                
            Else
                TestStop2 = 1
            End If
        
        Loop Until (TestStop1 = 1) And (TestStop2 = 1) Or AllenStop = 1
    
        ' wait Tester response END
        If HubEnaOn = 1 Then
            k = DO_WritePort(card, Channel_P1CH, &HF) ' set HUB module Ena off
            HubEnaOn = 0
        End If
        
        TestCounter = TestCounter + 1 ' Allen Debug
        
        If TestCycleTime1 > WAIT_TEST_CYCLE_OUT And SiteCheck(0).value = 1 Then
            WaitTestTimeOut1 = 1
            WaitTestTimeOutCounter1 = WaitTestTimeOutCounter1 + 1
        End If
                  
        If TestCycleTime2 > WAIT_TEST_CYCLE_OUT And SiteCheck(1).value = 1 Then
            WaitTestTimeOut2 = 1
            WaitTestTimeOutCounter2 = WaitTestTimeOutCounter2 + 1
        End If
       
    End If  ' Test end
End Sub
