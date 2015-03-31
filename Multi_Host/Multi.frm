VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Host 
   Caption         =   "Multi "
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   13620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Report 
      Caption         =   "Report"
      Height          =   375
      Left            =   10800
      TabIndex        =   132
      Top             =   8160
      Width           =   975
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      Height          =   195
      Index           =   7
      Left            =   1320
      TabIndex        =   50
      Top             =   1440
      Width           =   735
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      Height          =   195
      Index           =   6
      Left            =   1320
      TabIndex        =   49
      Top             =   1440
      Width           =   735
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      Height          =   195
      Index           =   5
      Left            =   1320
      TabIndex        =   48
      Top             =   1440
      Width           =   735
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      Height          =   195
      Index           =   4
      Left            =   1320
      TabIndex        =   47
      Top             =   1440
      Width           =   735
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      Height          =   195
      Index           =   3
      Left            =   1320
      TabIndex        =   46
      Top             =   1440
      Width           =   735
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      Height          =   195
      Index           =   2
      Left            =   1320
      TabIndex        =   45
      Top             =   1440
      Width           =   735
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   44
      Top             =   1440
      Width           =   735
   End
   Begin VB.Timer Blank_Timer 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   9960
      Top             =   1200
   End
   Begin VB.OptionButton Dummy 
      Caption         =   "Dummy"
      Height          =   255
      Left            =   8880
      TabIndex        =   33
      Top             =   1200
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Hander_SRM 
      Caption         =   "SRM"
      Height          =   255
      Left            =   9120
      TabIndex        =   32
      Top             =   720
      Width           =   735
   End
   Begin VB.OptionButton Hander_NS 
      Caption         =   "NS"
      Height          =   255
      Left            =   8400
      TabIndex        =   31
      Top             =   720
      Width           =   615
   End
   Begin VB.ComboBox ChipNameCombo2 
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
      Left            =   1560
      TabIndex        =   29
      Text            =   "Ver"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton EndBtn 
      Caption         =   "End"
      Height          =   495
      Left            =   12120
      TabIndex        =   28
      Top             =   600
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
      Height          =   255
      Left            =   8520
      TabIndex        =   27
      Top             =   120
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
      Left            =   11400
      TabIndex        =   26
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton StopBtn 
      Caption         =   "STOP"
      Height          =   495
      Left            =   11040
      TabIndex        =   18
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox ChipNameCombo 
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
      Left            =   120
      TabIndex        =   17
      Text            =   "ChipName"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CheckBox SiteCheck 
      Caption         =   "Site"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton BeginBtn 
      Caption         =   "Begin"
      Height          =   495
      Left            =   9960
      TabIndex        =   14
      Top             =   600
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
      Height          =   255
      Left            =   5280
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
      Left            =   9840
      TabIndex        =   12
      Top             =   120
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
      Left            =   4440
      TabIndex        =   9
      Text            =   "4"
      Top             =   120
      Width           =   615
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   0
      Left            =   10440
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   1
      Left            =   11160
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   2
      Left            =   11880
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   3
      Left            =   12600
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   4
      Left            =   10440
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   5
      Left            =   11160
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   6
      Left            =   11880
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   7
      Left            =   12600
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label UPH 
      Height          =   375
      Left            =   11880
      TabIndex        =   134
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label avgTestTimeLbl 
      Caption         =   "AvgTestTime"
      Height          =   255
      Left            =   4680
      TabIndex        =   133
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label TotalTestTitle 
      Caption         =   "TotalTest"
      Height          =   375
      Left            =   12000
      TabIndex        =   131
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label TotalTest 
      Caption         =   "TotalTest"
      Height          =   495
      Index           =   7
      Left            =   1320
      TabIndex        =   130
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label TotalTest 
      Caption         =   "TotalTest"
      Height          =   495
      Index           =   6
      Left            =   1320
      TabIndex        =   129
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label TotalTest 
      Caption         =   "TotalTest"
      Height          =   495
      Index           =   5
      Left            =   1320
      TabIndex        =   128
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label TotalTest 
      Caption         =   "TotalTest"
      Height          =   495
      Index           =   4
      Left            =   1320
      TabIndex        =   127
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label TotalTest 
      Caption         =   "TotalTest"
      Height          =   495
      Index           =   3
      Left            =   1320
      TabIndex        =   126
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label TotalTest 
      Caption         =   "TotalTest"
      Height          =   495
      Index           =   2
      Left            =   1320
      TabIndex        =   125
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label TotalTest 
      Caption         =   "TotalTest"
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   124
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label TotalTest 
      Caption         =   "TotalTest"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   123
      Top             =   7680
      Width           =   735
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   122
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      Height          =   495
      Index           =   6
      Left            =   1320
      TabIndex        =   121
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      Height          =   495
      Index           =   5
      Left            =   1320
      TabIndex        =   120
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      Height          =   495
      Index           =   4
      Left            =   1320
      TabIndex        =   119
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      Height          =   495
      Index           =   3
      Left            =   1320
      TabIndex        =   118
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      Height          =   495
      Index           =   2
      Left            =   1320
      TabIndex        =   117
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label ContFailTitle 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   9600
      TabIndex        =   116
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   115
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label ContFail 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   114
      Top             =   8280
      Width           =   855
   End
   Begin VB.Line Line2 
      X1              =   960
      X2              =   960
      Y1              =   960
      Y2              =   8160
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   113
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   112
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   111
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   110
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   109
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   108
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   107
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   106
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   105
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   104
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   103
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   102
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   101
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   100
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   99
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   98
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   97
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   96
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   95
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   94
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   93
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   92
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   91
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   90
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   89
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   88
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   87
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   86
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   85
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   84
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   83
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   82
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   81
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   80
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   79
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   78
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   77
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   76
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   75
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   74
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   73
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   72
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   71
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   70
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   69
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   68
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   67
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   66
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   65
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   64
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   63
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   62
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   61
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   60
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   59
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   58
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   57
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   56
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   55
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   54
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   53
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   52
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   51
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Bin4Title 
      Caption         =   "Bin4"
      Height          =   375
      Left            =   12000
      TabIndex        =   43
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Bin2Title 
      Caption         =   "Bin2"
      Height          =   375
      Left            =   12000
      TabIndex        =   42
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   41
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   40
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   39
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   38
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   37
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   36
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   35
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label TestCycleTimeLbl 
      Caption         =   "TestCycleTime"
      Height          =   255
      Left            =   2400
      TabIndex        =   25
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label SiteLbl 
      Caption         =   "Sites"
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
      Left            =   3720
      TabIndex        =   24
      Top             =   120
      Width           =   615
   End
   Begin VB.Label YieldTitle 
      Caption         =   "Yield"
      Height          =   375
      Left            =   12000
      TabIndex        =   23
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label TotalFailTitle 
      Caption         =   "Fail"
      Height          =   375
      Left            =   12000
      TabIndex        =   22
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label TimeOutTitle 
      Caption         =   "TimeOut"
      Height          =   375
      Left            =   12000
      TabIndex        =   21
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Bin5Title 
      Caption         =   "Bin5"
      Height          =   375
      Left            =   12000
      TabIndex        =   20
      Top             =   5160
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   13200
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label TestResultLbl 
      Caption         =   "TestResult"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Bin3Title 
      Caption         =   "Bin3"
      Height          =   375
      Left            =   12000
      TabIndex        =   15
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Bin1Title 
      Caption         =   "Bin1"
      Height          =   375
      Left            =   12000
      TabIndex        =   11
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label CycleTimeLbl 
      Caption         =   "MAX CycleTime"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Yield 
      Caption         =   "Yield"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label TotalFail 
      Caption         =   "TotalFail"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label TimeOut 
      Caption         =   "TimeOut"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Bin5 
      Caption         =   "Bin5"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Bin4 
      Caption         =   "Bin4"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Bin3 
      Caption         =   "Bin3"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Bin2 
      Caption         =   "Bin2"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Bin1 
      Caption         =   "Bin1"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Hander Type :"
      Height          =   255
      Left            =   7080
      TabIndex        =   30
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Hander_Blank 
      Height          =   495
      Left            =   7080
      TabIndex        =   34
      Top             =   600
      Width           =   2775
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
'Const Bin1 = 5
'Const Bin2 = 6
'Const Bin3 = 7
'Const Bin4 = 8
'Const Bin5 = 9
 
Dim CPort(0 To 7) As Byte

Dim FirstTimeFlag As Byte
 
Dim ChipName As String
Dim Channel As Byte
Dim i As Byte
Dim result

Public ALCOR As Boolean
Public ALCORChipName As String

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal brevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hmenu As Long, ByVal nposition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hmenu As Long) As Long

Private Const SC_CLOSE = &HF060
Private Const MF_REMOVE = &H1000&

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

Dim FailSite(0 To 7) As Long
Dim FailTotal As Long
Dim FailPercent As Single

Dim Bin2Total As Long
Dim Bin2Percent As Single
Dim Bin3Total As Long
Dim Bin3Percent As Single
Dim Bin4Total As Long
Dim Bin4Percent As Single
Dim Bin5Total As Long
Dim Bin5Percent As Single

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
    
    Open "D:\SLT Summary\" & OutFileName For Output As #1
    
    Print #1, "#####################################################"
    Print #1, "Name of PC: " & NameofPC
    Print #1, "Program Name: " & ProgramName
    Print #1, "Program Rersion Code: " & ProgramRevisionCode
    Print #1, "Device ID: " & DeviceID
    Print #1, "Run Card NO: " & RunCardNO
    Print #1, "Lot ID: " & LotID
    Print #1, "Process: " & ProcessIDSum
    Print #1, "Start at: " & StartAtMin
    Print #1, "End at: " & EndAtMax
    Print #1, "HandelerID: " & HandlerID
    Print #1, "Operator Name: " & OperatorName
    Print #1,
    Print #1, "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Print #1, Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Site 3 " & Space(3) & "Site 4 " & Space(3) & "Site 5 " & Space(3) & "Site 6 " & Space(3) & "Site 7 " & Space(3) & "Site 8 " & Space(3) & "Total  " & Space(3) & "Total"
    Print #1, Space(13) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percen"
    Print #1, "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

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
    Close #1

    '============================ printer section ===========================
    'Printer.CurrentX = 300

    If ReportDebug = 1 Then
        Exit Sub
    End If
    
    Printer.FontSize = 10
    Printer.Font = "標楷體"
    Printer.Print "#####################################################"
    Printer.Print "Name of PC: " & NameofPC
    Printer.Print "Program Name: " & ProgramName
    Printer.Print "Program Rersion Code: " & ProgramRevisionCode
    Printer.Print "Device ID: " & DeviceID
    Printer.Print "Run Card NO: " & RunCardNO
    Printer.Print "Lot ID: " & LotID
    Printer.Print "Process: " & ProcessIDSum
    Printer.Print "Start at: " & StartAtMin
    Printer.Print "End at: " & EndAtMax
    Printer.Print "HandelerID: " & HandlerID
    Printer.Print "Operator Name: " & OperatorName
    Printer.Print
    Printer.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Site 3 " & Space(3) & "Site 4 " & Space(3) & "Site 5 " & Space(3) & "Site 6 " & Space(3) & "Site 7 " & Space(3) & "Site 8 " & Space(3) & "Total  " & Space(3) & "Total"
    Printer.Print Space(13) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percen"
    Printer.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

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
    
    '=============== file output
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
    sDBPAth = "D:\SLT Summary\MultiSummary.mdb"
    'Debug.Print "1"; Dir(sDBPAth, vbNormal + vbDirectory)
    If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
        MsgBox "MDB no EXIST"
        Exit Sub
    End If

    'sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDBPAth & ";Persist   Security   Info=False;Jet   OLEDB:Database   Password=058f"
    sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\MultiSLT.mdb"
     
    ' ------------------------
    ' Create New ADOX Object
    ' ------------------------
    'Set oDB = New ADOX.Catalog
    'oDB.Create sConStr
    
    Set oCn = New ADODB.Connection
    oCn.ConnectionString = sConStr
    oCn.Open
    
    Set oCM = New ADODB.Command
    oCM.ActiveConnection = oCn
    
    Set RS = New ADODB.Recordset
    
    cmstr = "SELECT Min(Summary.StartAt) as StartATMin,Max(Summary.EndAt) as EndAtMax  " & _
            "from Summary where RunCardNO= '" & RunCardNO & "' and ProcessID= '" & ProcessIDSum & "'"
            
    'cmstr = "SELECT Min(Summary.StartAt) as StartATMin,Max(Summary.EndAt) as EndAt.Max " & _
            "from Summary where  ProcessID= '" & ProcessIDSum & "'"
            
    oCM.CommandText = cmstr
    Debug.Print cmstr
    Set RS = oCM.Execute
    
    'StartAtMin = rs.Fields("StartATMin")
    'Debug.Print StartAtMin
    'cmstr = "SELECT Max(Summary.EndAt) as EndAtMax " & _
            "from Summary where RunCardNO= '" & RunCardNO & "' and ProcessID= '" & ProcessIDSum & "'"
            
    'cmstr = "SELECT Min(Summary.StartAt) as StartATMin, " & _
            "from Summary where  ProcessID= '" & ProcessIDSum & "'"
            
    'oCM.CommandText = cmstr
    'Debug.Print cmstr
    'Set rs = oCM.Execute
    
    StartAtMin = RS.Fields("StartATMin")
    Debug.Print StartAtMin
    
    EndAtMax = RS.Fields("EndAtMax")
    Debug.Print EndAtMax
     
    'cmstr = "SELECT Sum(Summary.Bin1Site1) as Bin1Site1Sum ," & _
    '        "Sum(Summary.Bin2Site1) as Bin2Site1Sum ," & _
    '        "Sum(Summary.Bin3Site1) as Bin3Site1Sum ," & _
    '        "Sum(Summary.Bin4Site1) as Bin4Site1Sum ," & _
    '        "Sum(Summary.Bin5Site1) as Bin5Site1Sum ," & _
    '       "Sum(Summary.Bin1Site2) as Bin1Site2Sum ," & _
    '        "Sum(Summary.Bin2Site2) as Bin2Site2Sum ," & _
    '        "Sum(Summary.Bin3Site2) as Bin3Site2Sum ," & _
    '        "Sum(Summary.Bin4Site2) as Bin4Site2Sum ," & _
    '       "Sum(Summary.Bin5Site2) as Bin5Site2Sum " & _
    '        "from Summary where RunCardNO= '" & RunCardNO & "' and ProcessID= '" & ProcessIDSum & "' "
         
    For i = 0 To 6
        Cmstr11 = Cmstr11 & "Sum(Summary.Bin1_" & CStr(i) & ") as Bin1Sum" & CStr(i) & " ,"
    Next i
    Cmstr12 = "Sum(Summary.Bin1_7) as Bin1Sum7,"
    
    'Debug.Print Cmstr1 & Cmstr2
     
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
    'Debug.Print cmstr
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
    
    RS.Close
    'cmstr = "UPDATE Summary SET " & _
    ' "EndAt= '" & EndAt & "' ," & _
    '"Bin1Site1=" & CStr(Bin1Site1) & "," & _
    '"Bin1Site2=" & CStr(Bin1Site2) & "," & _
    '"Bin2Site1=" & CStr(Bin2Site1) & "," & _
    '"Bin2Site2=" & CStr(Bin2Site2) & "," & _
    '"Bin3Site1=" & CStr(Bin3Site1) & "," & _
    '"Bin3Site2=" & CStr(Bin3Site2) & "," & _
    '"Bin4Site1=" & CStr(Bin4Site1) & "," & _
    '"Bin4Site2=" & CStr(Bin4Site2) & "," & _
    '"Bin5Site1=" & CStr(Bin5Site1) & "," & _
    '"Bin5Site2=" & CStr(Bin5Site1) & _
    '" where StartAT= '" & StartAt & "'"
    'oCM.CommandText = cmstr
    'Debug.Print cmstr
    'oCM.Execute
    
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
    'If err <> 0 Then
    'err.Clear
    'Resume Next
    'End If

End Sub
Sub PrintReport()
On Error Resume Next
Dim i As Byte

' Summary
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
                
                
    '=================================================================
    'Call UpdateDB
    '=======================================================================
    Open "D:\SLT Summary\" & OutFileName For Output As #1
    
    Print #1, "#####################################################"
    Print #1, "Name of PC: " & NameofPC
    Print #1, "Program Name: " & ProgramName
    Print #1, "Program Rersion Code: " & ProgramRevisionCode
    Print #1, "Device ID: " & DeviceID
    Print #1, "Run Card NO: " & RunCardNO
    Print #1, "Lot ID: " & LotID
    Print #1, "Process: " & ProcessID
    Print #1, "Start at: " & StartAt
    Print #1, "End at: " & EndAt
    Print #1, "HandelerID: " & HandlerID
    Print #1, "Operator Name: " & OperatorName
    Print #1,
    Print #1, "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Print #1, Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Site 3 " & Space(3) & "Site 4 " & Space(3) & "Site 5 " & Space(3) & "Site 6 " & Space(3) & "Site 7 " & Space(3) & "Site 8 " & Space(3) & "Total  " & Space(3) & "Total"
    Print #1, Space(13) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percen"
    Print #1, "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
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
    
    
    '=============== file output
    
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
    
    Close #1
    
    '=================================================== printer section ===========================

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
        'DoEvents
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

    EndSecond = Format(Now, "HH:MM:SS")
    EndDay = Format(Now, "YYYY/MM/DD")
   
    
    EndAt = EndDay & Space(1) & EndSecond
    

    '-----------------------------
    ' set Path and connection string
    '---------------------------
    sDBPAth = "D:\SLT Summary\MultiSummary.mdb"
    'Debug.Print "1"; Dir(sDBPAth, vbNormal + vbDirectory)
    If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
        MsgBox "MDB no EXIST"
        Exit Sub
    End If
    
     
    
    'sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDBPAth & ";Persist   Security   Info=False;Jet   OLEDB:Database   Password=058f"
    sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\MultiSLT.mdb"
     
    ' ------------------------
    ' Create New ADOX Object
    ' ------------------------
    'Set oDB = New ADOX.Catalog
    'oDB.Create sConStr
    
    Set oCn = New ADODB.Connection
    oCn.ConnectionString = sConStr
    oCn.Open
    
    Set oCM = New ADODB.Command
    oCM.ActiveConnection = oCn
    
    
     Dim cmstr As String
     Dim Cmstr1 As String
     Dim Cmstr2 As String
     Dim Cmstr3 As String
     Dim Cmstr4 As String
     Dim Cmstr5 As String
     Dim Cmstr0 As String
     Dim i As Byte
    
    'cmstr = "INSERT INTO Summary VALUES(" & _

    'oCM.CommandText = "INSERT INTO Summary VALUES(" & _

    'cmstr = "INSERT INTO Summary VALUES(" & _

    'oCM.CommandText = "UPDATE Summary SET" & _

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
    
    
    ' ------------------------
    ' Error Handling
    ' ------------------------
Err_Handler:
    'If err <> 0 Then
    'err.Clear
    'Resume Next
    'End If
    
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
    
    '-----------------------------
    ' set Path and connection string
    '---------------------------
    sDBPAth = "D:\SLT Summary\MultiSummary.mdb"
    'Debug.Print "1"; Dir(sDBPAth, vbNormal + vbDirectory)
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
    'If err <> 0 Then
    'err.Clear
    'Resume Next
    'End If
End Sub

Public Sub PCIBinSub(i As Byte)

On Error Resume Next
'TestResult(i) = "PASS"
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
            'Call Timer_1ms(2)
            Call PCI7296_bin(Channel, PCI7296_PASS)
            'Print i, "bin"
        Case "UNKNOW", "bin2", "Bin2"
            Call PCI7296_bin(Channel, PCI7296_BIN2)
        Case "gponFail", "bin3", "Bin3", "SD_WF", "SD_RF", "CF_WF", "CF_RF"
            Call PCI7296_bin(Channel, PCI7296_BIN3)
        Case "XD_WF", "bin4", "Bin4", "XD_RF"
            Call PCI7296_bin(Channel, PCI7296_BIN4)
        Case "MS_WF", "bin5", "Bin5", "MS_RF"
            Call PCI7296_bin(Channel, PCI7296_BIN5)
            
        ' mark on 20131129
'        Case "TimeOut"
'            Call PCI7296_bin(Channel, PCI7296_EOT)
            
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
                If AllenDebug = 0 Then
                    Call PCIBinSub(k)
                    State(k) = IdleState
                    MSComm1(k).InBufferCount = 0
                    MSComm1(k).InputLen = 0
                Else
                    State(k) = IdleState
                End If
            End If
        Else
            If SiteCheck(k).value = 1 Then
                If AllenDebug = 0 Then
                    Call PCIBinSub(k)
                    State(k) = IdleState
                    MSComm1(k).InBufferCount = 0
                    MSComm1(k).InputLen = 0
                Else
                    State(k) = IdleState
                End If
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
        '     Call PCI7296_bin(Channel, PCI7296_PASS)
            CurBinning = 1
            
        Case "UNKNOW", "bin2", "Bin2"
             TestResultLbl(i).Caption = "bin2"
              TestResultLbl(i).BackColor = RED_COLOR
            Bin2Counter(i) = Bin2Counter(i) + 1
            Bin2(i).Caption = CStr(Bin2Counter(i))
            Bin2ContiFail(i) = Bin2ContiFail(i) + 1
            ContiFailCounter(i) = ContiFailCounter(i) + 1
           ' Call PCI7296_bin(Channel, PCI7296_BIN2)
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
           
            ' Call PCI7296_bin(Channel, PCI7296_BIN3)
        Case "XD_WF", "bin4", "Bin4", "XD_RF"
             TestResultLbl(i).Caption = "bin4"
              TestResultLbl(i).BackColor = YELLOW_COLOR
            Bin4Counter(i) = Bin4Counter(i) + 1
            ContiFailCounter(i) = ContiFailCounter(i) + 1
            Bin4(i).Caption = CStr(Bin4Counter(i))
         '   Call PCI7296_bin(Channel, PCI7296_BIN4)
            CurBinning = 4
        Case "MS_WF", "bin5", "Bin5", "MS_RF"
           TestResultLbl(i).Caption = "bin5"
            TestResultLbl(i).BackColor = YELLOW_COLOR
            Bin5Counter(i) = Bin5Counter(i) + 1
            ContiFailCounter(i) = ContiFailCounter(i) + 1
            Bin5(i).Caption = CStr(Bin5Counter(i))
         '    Call PCI7296_bin(Channel, PCI7296_BIN5)
            CurBinning = 5
            
'        ' mark in 20131129
'        Case "TimeOut"
'            TestResultLbl(i).Caption = "TimeOut"
'            TestResultLbl(i).BackColor = YELLOW_COLOR
        
        Case Else
            TestResultLbl(i).Caption = "bin2"
              TestResultLbl(i).BackColor = RED_COLOR
            Bin2Counter(i) = Bin2Counter(i) + 1
            Bin2(i).Caption = CStr(Bin2Counter(i))
            ContiFailCounter(i) = ContiFailCounter(i) + 1
         '   Call PCI7296_bin(Channel, PCI7296_BIN2)
            CurBinning = 2
    End Select
 
 
    Call UpdateDB
    'CurBinning = "#"
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
    TotalTest(i).Caption = tmpFail + Bin1Counter(i)
     
    If OffLineCheck.value = 0 Then
    
        BinFlag(i) = 1   ' set Bin flag
    
        OneCycleFlag = 1
        
        For k = 0 To 7
            If BinFlag(k) <> StartFlag(k) Then
            OneCycleFlag = 0
            Exit For
            End If
        Next k
        
        If OneCycleFlag = 1 Then   ' reset all falg
        
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
            TestCycleTimeLbl.Caption = "TestTime:" & CStr(TotalCycleTime)
        End If
    End If

    UPH.Caption = "UPH: " & (3600 / avgTestTime)

End Sub
Public Sub BinSubOld(i)
On Error Resume Next

Dim result
Dim k As Byte
          
Select Case TestResult(i)
        Case "PASS"
             TestResultLbl(i).Caption = "PASS"
             TestResultLbl(i).BackColor = GREEN_COLOR
             Bin1Counter(i) = Bin1Counter(i) + 1
             Bin1(i).Caption = CStr(Bin1Counter(i))
             Bin2ContiFail(i) = 0
             Call PCI7296_bin(Channel, PCI7296_PASS)
           
        Case "UNKNOW", "bin2", "Bin2"
             TestResultLbl(i).Caption = "bin2"
              TestResultLbl(i).BackColor = RED_COLOR
            Bin2Counter(i) = Bin2Counter(i) + 1
            Bin2(i).Caption = CStr(Bin2Counter(i))
            Bin2ContiFail(i) = Bin2ContiFail(i) + 1
            Call PCI7296_bin(Channel, PCI7296_BIN2)
           
            If Bin2ContiFail(i) > 5 And ResetPC.value = 1 Then
               
               result = DO_WritePort(card, Channel, PCI7296_PC_RESET)
               Call MsecDelay(PC_RESET_TIME)
              result = DO_WritePort(card, Channel, &HFF)
            End If
                    
             
             
        Case "gponFail", "bin3", "Bin3", "SD_WF", "SD_RF", "CF_WF", "CF_RF"
        
            TestResultLbl(i).Caption = "bin3"
            TestResultLbl(i).BackColor = YELLOW_COLOR
            Bin3Counter(i) = Bin3Counter(i) + 1
           
            Bin3(i).Caption = CStr(Bin3Counter(i))
            
             Call PCI7296_bin(Channel, PCI7296_BIN3)
        Case "XD_WF", "bin4", "Bin4", "XD_RF"
             TestResultLbl(i).Caption = "bin4"
              TestResultLbl(i).BackColor = YELLOW_COLOR
            Bin4Counter(i) = Bin4Counter(i) + 1
            Bin4(i).Caption = CStr(Bin4Counter(i))
            Call PCI7296_bin(Channel, PCI7296_BIN4)
        Case "MS_WF", "bin5", "TimeOut", "Bin5", "MS_RF"
           TestResultLbl(i).Caption = "bin5"
            TestResultLbl(i).BackColor = YELLOW_COLOR
            Bin5Counter(i) = Bin5Counter(i) + 1
            Bin5(i).Caption = CStr(Bin5Counter(i))
             Call PCI7296_bin(Channel, PCI7296_BIN5)
        Case Else
            TestResultLbl(i).Caption = "bin2"
              TestResultLbl(i).BackColor = RED_COLOR
            Bin2Counter(i) = Bin2Counter(i) + 1
            Bin2(i).Caption = CStr(Bin2Counter(i))
            Call PCI7296_bin(Channel, PCI7296_BIN2)
 End Select

 Call UpdateDB
 
 State(i) = IdleState
 



 
Dim tmp As Single
Dim tmpFail As Long
Dim tmps As String
 
tmpFail = (Bin2Counter(i) + Bin3Counter(i) + Bin4Counter(i) + Bin5Counter(i))

TotalFail(i).Caption = CStr(tmpFail)
tmp = CSng(Bin1Counter(i) / (tmpFail + Bin1Counter(i)))
tmps = Format$(tmp, "0.00%")
Yield(i).Caption = tmps
 
  MSComm1(i).InBufferCount = 0
  MSComm1(i).InputLen = 0
  
  
If OffLineCheck.value = 0 Then

    BinFlag(i) = 1   ' set Bin flag

    OneCycleFlag = 1
    
    For k = 0 To 7
      If BinFlag(k) <> StartFlag(k) Then
      OneCycleFlag = 0
      Exit For
      End If
    Next k
    
    If OneCycleFlag = 1 Then   ' reset all falg
    
       For k = 0 To 7
          BinFlag(i) = 0
          StartFlag(i) = 0
        
       Next k
       StartCounter = 0
    
        TotalCycleTime = Timer - MinGetStartTime
          '  Debug.Print Timer
        TestCycleTimeLbl.Caption = "TestTime:" & CStr(TotalCycleTime)
        
     End If
         
Else
  
    BinCounter = BinCounter + 1
   Debug.Print "BinCounterr="; BinCounter

  If BinCounter = OffLineSiteCounter Then
     OneCycleFlag = 1
 
     TotalCycleTime = Timer - MinGetStartTime
      '  Debug.Print Timer
    TestCycleTimeLbl.Caption = "TestTime:" & CStr(TotalCycleTime)
        
  End If
  
End If
 'Debug.Print State(i)
End Sub
Sub PCI7296_bin(Channel As Byte, PCI7296bin As Byte)

result = DO_WritePort(card, Channel, PCI7296bin)
Call Timer_1ms(12)
    
'========================================
result = DO_WritePort(card, Channel, PCI7296bin - PCI7296_EOT)
Call Timer_1ms(7)
        
'=======================================
result = DO_WritePort(card, Channel, PCI7296bin)
Call Timer_1ms(7)
       
'========================================
result = DO_WritePort(card, Channel, &HFF)

'Const PCI7248_EOT = &H1 'for 7248 card
'Const PCI7248_PASS = &HFD 'for 7248 card  11111101
'Const PCI7248_BIN2 = &HFB 'for 7248 card  11111011
'Const PCI7248_BIN3 = &HF7 'for 7248 card  11110111
'Const PCI7248_BIN4 = &HEF 'for 7248 card  11101111
'Const PCI7248_BIN5 = &HDF 'for 7248 card  11011111
  
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
Public Sub FirtTimeSub()
Dim k As Byte
Dim tmp As Byte

    SelectFailFlag = 0
    If ChipName = "" Then
        MsgBox "please select Chip"
        SelectFailFlag = 1
        Exit Sub
    End If

    If FirstTimeFlag = 0 Then
        BinCounter = 0
        StartCounter = 0
        OffLineSiteCounter = 0
        
        For k = 0 To 7
            StartFlag(k) = 0
            BinFlag(k) = 0
            State(k) = IdleState
        Next k
        
        tmp = PCI7296_CLEAR_START
        result = DO_WritePort(card, Channel_P1B, tmp)
        result = DO_WritePort(card, Channel_P1C, tmp)
        result = DO_WritePort(card, Channel_P2A, tmp)
        result = DO_WritePort(card, Channel_P2B, tmp)
        result = DO_WritePort(card, Channel_P2C, tmp)
        result = DO_WritePort(card, Channel_P3A, tmp)
        result = DO_WritePort(card, Channel_P3B, tmp)
        result = DO_WritePort(card, Channel_P3C, tmp)
        
        If AllenDebug = 0 Then
            Call Timer_1ms(10)
        End If
        
        tmp = &HFF
        result = DO_WritePort(card, Channel_P1B, tmp)
        result = DO_WritePort(card, Channel_P1C, tmp)
        result = DO_WritePort(card, Channel_P2A, tmp)
        result = DO_WritePort(card, Channel_P2B, tmp)
        result = DO_WritePort(card, Channel_P2C, tmp)
        result = DO_WritePort(card, Channel_P3A, tmp)
        result = DO_WritePort(card, Channel_P3B, tmp)
        result = DO_WritePort(card, Channel_P3C, tmp)
        
        If OffLineCheck.value = 0 Then
            For i = 0 To CInt(SiteCombo.Text) - 1  ' auto mode
                Status(i).Caption = IdleState
                SiteCheck(i).value = 1
            Next i
        Else
            For i = 0 To 7  ' off line mode
                If SiteCheck(i).value = 1 Then
                    OffLineSiteCounter = OffLineSiteCounter + 1
                    Status(i).Caption = IdleState
                End If
            Next i
                    
            If OffLineSiteCounter = 0 Then
                MsgBox "please select Site"
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
            TimeOutCounter(i) = 0
            TotalTest(i) = 0
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
Private Sub Card_Initial()
  
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
        Loop Until old_value1 <> old_value2

        Do
            DoEvents
            result = CTR_Read(0, 2, old_value2)
        Loop Until old_value1 = old_value2
    Next
     
End Sub

Private Sub BeginBtn_Click()
Dim i As Byte

    avgTestTime = 0
    totalTestTime = 0
    testTime = 0
    UPH = 0

    If Dummy.value = True Then
        MsgBox ("Please Select Handler Type")
        Blank_Timer.Enabled = True
        Exit Sub
    End If
        
    PrintEnable = 0
    StopFlag = 0
    AllenStop = 0
    
    Call FirtTimeSub
    
    If SelectFailFlag = 1 Then
        Exit Sub
    End If
    
    If AllenDebug = 1 Then
        DebugEntryTime = Timer
    End If

    Call LockOption

    'HubTestEnd = 0
    'SiteCheckCount = 0
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
    '    EnCheck(i) = False
    '    GetBinning(i) = False
        ContiFailCounter(i) = 0
    Next
    
    ReportBegin = 0
    ' report control  begin
    Call ReportActive
    
    If ReportCheck.value = 1 Then
        Do
            DoEvents
        Loop While (ReportBegin = 0) And (StopFlag = 0)
    End If
    
    ReportBegin = 0
    GetFirstStart = False
    
    Do
MultiLoop:
        For i = 0 To 7
            
            DoEvents
            
            If SiteCheck(i).value = 1 Then
            'assign Channel
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
    
            'assign State
        
                Debug.Print "q"; i; State(i)
                Select Case State(i)
                    Case IdleState
                        Debug.Print "1"
                        Call HandlerStartSub(i)         '1. wait start  '2. change to HandlerStartState
                        Status(i).BackColor = GREEN_COLOR
                    Case HandlerStartState
                        Debug.Print "2"
                        Call PCReadySub(i)  '1 . wait for PC reader
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
                        Call TestingSub(i)  ' next state
                        Status(i).BackColor = YELLOW_COLOR
                    Case BinState
                        Debug.Print "6"
                        Call BinSub(i) ' back to idle state
                        Status(i).BackColor = GREEN_COLOR
                    
                        If (Bin2ContiFail(i) >= 3) And (HubNonUPT2Flag = False) Then
                            ResetUPT2_Flag = True
                        End If
                    Case PCIState
                        Call PCIStateSub
                        If PCIFlag = 1 Then
                            If (OneCycleCheck.value = 1) Then
                                AllenStop = 1
                                MsgBox "one cycel test Finish"
                                FirstTimeFlag = 0
                            End If
      
                            If StopFlag = 1 Then
                                AllenStop = 1
                                MsgBox "STOP test Finish"
                                FirstTimeFlag = 0
                            End If
                            PCIFlag = 0
                            
                            If UPT2TestFlag Then
                                Call UPT2RoutineSub
                            End If
                            
                            GoTo MultiLoop
                        End If
                End Select
            End If
        
'            If UPT2TestFlag = True Then
'
'                If (State(i) = PCIState) Then
'                    GetBinning(i) = False
'                End If
'
'                If (GetBinning(i) = False) And (State(i) = BinState) Then
'                    GetBinning(i) = True
'                    RealSiteCount = RealSiteCount + 1
'                End If
'
'                If (RealSiteCount <> 0) And (RealSiteCount = SiteCheckCount) Then
'                    HubTestEnd = 1
'                    GetFirstStart = False
'                End If
'
'            End If
        
        Next
      
        For i = 0 To 7
            Status(i).Caption = State(i)
        Next i
        'Debug.Print OneCycleFlag
      
        If StopFlag = 1 Then
            AllenStop = 1
            FirstTimeFlag = 0
        End If
        
        OneCycleFlag = 0
    
    Loop While AllenStop = 0
    
    Call UnlockOption
    
    If OneCycleCheck.value = 0 Then
        Call PrintReport
        MsgBox "generate Report finsih"
    End If

End Sub
Public Sub UPT2RoutineSub()

    If (UPT2TestFlag = True) Then
        If (Host.Hander_NS = True) Then
            result = DO_WritePort(card, Channel_P3B, &HFF)  'Site1 ~ Site4 (bit1~bit4) Ena Off
            'HubEnaOn = 0
            'HubTestEnd = 0
            'RealSiteCount = 0
            'SiteCheckCount = 0
            'AllReady_flag = False
            'For i = 0 To 7
            '    EnCheck(i) = False
            '    GetBinning(i) = False
            'Next
            If AllenDebug = 1 Then
                DebugForm.textState = DebugForm.textState & " ; " & "Hub Ena OFF"
            End If
        End If
    
        If (Host.Hander_SRM = True) Then
            result = DO_WritePort(card, Channel_P4B, &HFF)  'Site1 ~ Site8 (bit1~bit8) Ena OFF
            'HubEnaOn = 0
            'HubTestEnd = 0
            'RealSiteCount = 0
            'SiteCheckCount = 0
            'AllReady_flag = False
            'For i = 0 To 7
            '    EnCheck(i) = False
            '    GetBinning(i) = False
            'Next
            If AllenDebug = 1 Then
                DebugForm.textState = DebugForm.textState & " ; " & "Hub Ena OFF"
            End If
        End If
    End If

    '20130327 modify UPT2 Reset when test end
    If (ResetUPT2_Flag = True) Then
        If Hander_NS.value = True Then
            result = DO_WritePort(card, Channel_P3A, &HFC)      'UPT2 Reset
            Call MsecDelay(0.2)
            result = DO_WritePort(card, Channel_P3A, &HFF)
            'Call MsecDelay(3#)
            If AllenDebug = 1 Then
                DebugForm.textState = DebugForm.textState & " ; " & "Hub Reset"
            End If
        End If
                
        If Hander_SRM.value = True Then
            result = DO_WritePort(card, Channel_P4A, &HFC)
            Call MsecDelay(0.2)
            result = DO_WritePort(card, Channel_P4A, &HFF)
            'Call MsecDelay(3#)
            If AllenDebug = 1 Then
                DebugForm.textState = DebugForm.textState & " ; " & "Hub Reset"
            End If
        End If
            
        ResetUPT2_Flag = False
    End If
    
    If AllenDebug = 1 Then
        DebugForm.textState = DebugForm.textState & vbCrLf
    End If

End Sub

Public Sub HandlerStartSub(i)

'Step1: wait for Handle Start

Dim result
Dim k As Byte
Dim j As Integer
Dim ReadStartSignal
Dim UPT2DetectStartTime
Dim ii As Byte
Dim TempValue As Byte
 
    TestResult(i) = ""
    
    'atheist debug
    '=======================
    'UPT2TestFlag = True
    '=======================
    
    If OffLineCheck.value = 0 Then
    
'        If (UPT2TestFlag = True) And (SiteCheckCount = 0) Then
'
'            If AllenDebug = 0 Then
'                result = DO_ReadPort(card, Channel_P1A, ReadStartSignal)
'            Else
'                ReadStartSignal = &HF0
'                DebugEntryTime = Timer
'            End If
'
'            If ReadStartSignal <> &HFF Then
'                For j = 0 To 7
'                    OldState(j) = State(j)
'                    GetStart(j) = 0
'                    TestResult(j) = ""
'                    'Change to Next State
'                    If OldState(j) <> State(j) Then
'                        Fire(j) = Timer
'                    End If
'                Next
'
'                UPT2DetectStartTime = Timer
'
'                Do
'                    For j = 0 To 7
'                        If (CAndValue(ReadStartSignal, CPort(j)) = 0) And (EnCheck(j) = False) Then
'
'                            If GetFirstStart = False Then
'                                GetFirstStart = True
'                            End If
'
'                            GetStart(j) = 1
'                            State(j) = HandlerStartState
'                            CycleTestTime(j) = Timer
'                            StartFlag(i) = 1
'                            EnCheck(j) = True
'
'                            If StartCounter = 0 Then
'                                OneCycleFlag = 0
'                            End If
'
'                            SiteCheckCount = SiteCheckCount + 1
'                            StartCounter = StartCounter + 1
'
'                            If StartCounter = 1 Then
'                                OneCycleFlag = 0
'                                MinGetStartTime = Timer
'                            End If
'
'                            If SiteCheckCount = 1 Then
'                                If (Host.Hander_NS = True) Then
'                                    result = DO_WritePort(card, Channel_P3B, &HF0)  'Site1 ~ Site4 (bit1~bit4) Ena ON
'                                    'HubEnaOn = 1
'                                End If
'
'                                If (Host.Hander_SRM = True) Then
'                                    result = DO_WritePort(card, Channel_P4B, &H0)   'Site1 ~ Site8 (bit1~bit8) Ena ON
'                                    'HubEnaOn = 1
'                                End If
'                            End If
'                            'Debug.Print MinGetStartTime
'
'                            PassTime(j) = PassTimeFcn(CByte(j))
'
'                            If PassTime(j) > HandlerStartTimeOut Then
'                                TestResultLbl(j).Caption = "Start Time Out"
'                            End If
'
'                            Debug.Print "StartCounter="; StartCounter
'                            'Clear Start
'                        End If
'                    Next
'                Loop Until (Timer - UPT2DetectStartTime > 0.5) Or (SiteCheckCount = CInt(SiteCombo))
'
'            End If
'
'        Else
            If AllenDebug = 0 Then
                result = DO_ReadPort(card, Channel_P1A, ReadStartSignal)
            Else
                TempValue = 0
                For ii = 0 To 3
                    If DebugForm.SiteOnOff(ii).BackColor = DebugSiteOn Then
                        TempValue = TempValue + 2 ^ (ii)
                    End If
                Next
                
                ReadStartSignal = &HFF - TempValue
            End If
            
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
                    If UPT2TestFlag Then
                        If (Host.Hander_NS = True) Then
                            result = DO_WritePort(card, Channel_P3B, &HF0)  'Site1 ~ Site4 (bit1~bit4) Ena ON
                            'HubEnaOn = 1
                            If AllenDebug = 1 Then
                                DebugForm.textState = DebugForm.textState & "Hub Ena On"
                            End If
                        End If
    
                        If (Host.Hander_SRM = True) Then
                            result = DO_WritePort(card, Channel_P4B, &H0)   'Site1 ~ Site8 (bit1~bit8) Ena ON
                            'HubEnaOn = 1
                            If AllenDebug = 1 Then
                                DebugForm.textState = DebugForm.textState & "Hub Ena On"
                            End If
                        End If
                    End If
                End If
                'Debug.Print MinGetStartTime
          
                Debug.Print "StartCounter="; StartCounter
                'Clear Start
                Exit Sub
            End If
    
            ' Time Out Conditions
            PassTime(i) = PassTimeFcn(CByte(i))
            If PassTime(i) > HandlerStartTimeOut Then
                TestResultLbl(i).Caption = "Start Time Out"
            End If
        'End If
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
            If UPT2TestFlag Then
                If (Host.Hander_NS = True) Then
                    result = DO_WritePort(card, Channel_P3B, &HF0)  'Site1 ~ Site4 (bit1~bit4) Ena ON
                    If AllenDebug = 1 Then
                        DebugForm.textState = DebugForm.textState & "Hub Ena On"
                    End If
                    'HubEnaOn = 1
                End If

                If (Host.Hander_SRM = True) Then
                    result = DO_WritePort(card, Channel_P4B, &H0)   'Site1 ~ Site8 (bit1~bit8) Ena ON
                    If AllenDebug = 1 Then
                        DebugForm.textState = DebugForm.textState & "Hub Ena On"
                    End If
                    'HubEnaOn = 1
                End If
            End If
        End If
        'Debug.Print MinGetStartTime
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
    If AllenDebug = 0 Then
        MSComm1(i).Output = "~"
    End If
End If

If AllenDebug = 0 Then
    TmpBuf = MSComm1(i).Input
Else
    If i < 4 Then
        TmpBuf = "Ready"
    End If
End If
Buf(i) = Buf(i) & TmpBuf

If InStr(1, Buf(i), "Rea") <> 0 Then
    If (UPT2TestFlag = True) And (HubNonUPT2Flag = False) Then
        If AllenDebug = 0 Then
            For j = 1 To Len(ChipName)
                MSComm1(i).Output = Mid(ChipName, j, 1)
                Call MsecDelay(0.02)
            Next
        End If
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
   
    If AllenDebug = 0 Then
        MSComm1(i).InBufferCount = 0
        MSComm1(i).InputLen = 0
    Else
        DebugEntryTime = Timer
    End If
    
    Buf(i) = ""
    State(i) = PCReadyState
    Exit Sub
End If

' Time Out Condition
PassTime(i) = PassTimeFcn(i)

If PassTime(i) > PCReadyTimeOut Then
    TestResultLbl(i).Caption = "PC Ready Time Out"
    'Reset PC
    State(i) = PCResetState ' Reset PC
    result = DO_WritePort(card, Channel, PCI7296_PC_RESET)
    Call MsecDelay(PC_RESET_TIME)
    result = DO_WritePort(card, Channel, &HFF)
End If

End Sub

Public Sub PCResetFailSub(i As Byte)
Dim j As Integer
'Step2: wait for PC Ready

'1
If OldState(i) <> State(i) Then
    Fire(i) = Timer
End If
OldState(i) = State(i)

'2 Action

If (UPT2TestFlag = True) And (HubNonUPT2Flag = False) Then
    If AllenDebug = 0 Then
        MSComm1(i).Output = "~"
    End If
End If

TmpBuf = MSComm1(i).Input
Buf(i) = Buf(i) & TmpBuf
If InStr(1, Buf(i), "Rea") <> 0 Then
    If (UPT2TestFlag = True) And (HubNonUPT2Flag = False) Then
        If AllenDebug = 0 Then
            For j = 1 To Len(ChipName)
                MSComm1(i).Output = Mid(ChipName, j, 1)
                Call MsecDelay(0.02)
            Next
        End If
        State(i) = PCReadyState
        Buf(i) = ""
        Exit Sub
    Else
        MSComm1(i).Output = ChipName
        MSComm1(i).InBufferCount = 0
        MSComm1(i).InputLen = 0
        State(i) = PCReadyState
        Buf(i) = ""
        Exit Sub
    End If
End If

'Time Out Condition
PassTime(i) = PassTimeFcn(i)
'Debug.Print "2"; PassTime(i)
If PassTime(i) > PCResetTimeOut Then
    TestResultLbl(i).Caption = "PC Reset Fail 2"
    result = DO_WritePort(card, Channel, PCI7296_PC_RESET)
    Call MsecDelay(PC_RESET_TIME)
    result = DO_WritePort(card, Channel, &HFF)
    State(i) = PCResetState ' Reset PC
End If
 
End Sub
Public Sub PCResetSub(i As Byte)
Dim j As Integer
'Step2: wait for PC Ready

'1
If OldState(i) <> State(i) Then
   Fire(i) = Timer
End If
OldState(i) = State(i)

'2 Action

If (UPT2TestFlag = True) And (HubNonUPT2Flag = False) Then
    If AllenDebug = 0 Then
        MSComm1(i).Output = "~"
    End If
End If

TmpBuf = MSComm1(i).Input
Buf(i) = Buf(i) & TmpBuf
If InStr(1, Buf(i), "Rea") <> 0 Then
    If (UPT2TestFlag = True) And (HubNonUPT2Flag = False) Then
        If AllenDebug = 0 Then
            For j = 1 To Len(ChipName)
                MSComm1(i).Output = Mid(ChipName, j, 1)
                Call MsecDelay(0.02)
            Next
        End If
        State(i) = PCReadyState
        Buf(i) = ""
        Exit Sub
    Else
        MSComm1(i).Output = ChipName
        MSComm1(i).InBufferCount = 0
        MSComm1(i).InputLen = 0
        State(i) = PCReadyState
        Buf(i) = ""
        Exit Sub
    End If
End If


'Time Out Condition
PassTime(i) = PassTimeFcn(i)
'Debug.Print PassTime(i)

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
    If AllenDebug = 0 Then
        MSComm1(i).Output = "~"
    End If
End If

If AllenDebug = 0 Then
    If MSComm1(i).InBufferCount >= 3 Then
        TmpBuf = MSComm1(i).Input
        TestResult(i) = TestResult(i) & TmpBuf
    End If
Else
    If i = 0 Then
        If Timer > DebugEntryTime + 4 Then
            TestResult(i) = Trim(DebugForm.textS1Bin)
        End If
    ElseIf i = 1 Then
        If Timer > DebugEntryTime + 2 Then
            TestResult(i) = Trim(DebugForm.textS2Bin)
        End If
    ElseIf i = 2 Then
        If Timer > DebugEntryTime + 4 Then
            TestResult(i) = Trim(DebugForm.textS3Bin)
        End If
    ElseIf i = 3 Then
        If Timer > DebugEntryTime + 5 Then
            TestResult(i) = Trim(DebugForm.textS4Bin)
        End If
    End If
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
    TestResult(i) = "TimeOut"
    'Reset PC
    State(i) = BinState ' Reset PC
End If

End Sub

Private Sub Blank_Timer_Timer()
    If Hander_Blank.BackColor = System_Color Then
        Label1.BackColor = Blank_Color
        Hander_Blank.BackColor = Blank_Color
        Hander_NS.BackColor = Blank_Color
        Hander_SRM.BackColor = Blank_Color
    Else
        Label1.BackColor = System_Color
        Hander_Blank.BackColor = System_Color
        Hander_NS.BackColor = System_Color
        Hander_SRM.BackColor = System_Color
    End If
End Sub

Private Sub ChipNameCombo_Change()

    If ALCOR = True Then
         ChipNameCombo.Text = Left(ALCORChipName, 6)
         ChipNameCombo2.Enabled = True
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
    
    If ALCOR = True Then
        ChipNameCombo2.Text = Mid(ALCORChipName, 7, Len(ALCORChipName))
    End If
    
End Sub

Private Sub Command1_Click()
'Debug.Print Timer
 Call Timer_1ms(10)
Debug.Print Timer
Dim ReadStartSignal
Dim i As Byte
Dim result
Dim Channel
Dim x As Byte
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
 'result = DO_WritePort(card, Channel, PCI7296bin)
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
    
    If ALCOR = True Then
        ChipName = ALCORChipName
        ChipNameCombo.Text = ALCORChipName
    End If
    
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
    End If
        
    If DB_Path = "Tester" Then
        TestTimeOut = TesterRS.Fields(4)
        ResetPC.value = TesterRS.Fields(14)
        ReportCheck.value = TesterRS.Fields(15)
    End If
    
    If AllenDebug = 1 Then
        TestTimeOut = 65530
        ReportCheck.value = 0
        Hander_NS.value = True
    End If
    
    If SPIL_Flag Then
        If Dir(App.Path & "\" & ChipName) <> ChipName Then
            MsgBox ("Please Check Program Name !!")
            End
        End If
    End If
    
    CycleTimeLbl.Caption = "MAX Cycle Time :" & CStr(TestTimeOut) & " S"
    BeginBtn.Enabled = True

End Sub

Private Sub EndBtn_Click()
    End
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

    'On Error Resume Next
    'ReportCheck.value = 1  setting by DataBase

    ' for ContFail use

    SiteCombo.AddItem "2"
    SiteCombo.AddItem "4"
    SiteCombo.AddItem "6"
    SiteCombo.AddItem "8"

    'HubEnaOn = 0    'initial HUB ENA pin control

    'ChipNameCombo.AddItem "AU6433BLF25"

    BinLeft = 1200
    BinHeight = 400
    BinTop = 1200
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
        Bin2(i).Width = 1200
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
        
        TotalTest(i).Left = BinLeft + i * (BinWidth + 100)
        TotalTest(i).Height = BinHeight
        TotalTest(i).Width = BinWidth
        TotalTest(i).Top = BinTop + 500 * 10
    
        TotalTestTitle.Left = 100
        TotalTestTitle.Height = BinHeight
        TotalTestTitle.Top = TotalTest(i).Top
        TotalTestTitle.Width = TitleWidth
    
        Yield(i).Left = BinLeft + i * (BinWidth + 100)
        Yield(i).Height = BinHeight
        Yield(i).Width = BinWidth
        Yield(i).Top = BinTop + 500 * 11
    
        YieldTitle.Left = 100
        YieldTitle.Height = BinHeight
        YieldTitle.Top = Yield(i).Top
        YieldTitle.Width = TitleWidth
        
        ContFail(i).Left = BinLeft + i * (BinWidth + 100)
        ContFail(i).Height = BinHeight
        ContFail(i).Width = BinWidth - 450
        ContFail(i).Top = BinTop + 500 * 12

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

    If AllenDebug = 0 Then
        For i = 0 To 7  ' off line mode
            MSComm1(i).CommPort = i + 2
            MSComm1(i).Settings = "9600,N,8,1"
            
            If MSComm1(i).PortOpen = False Then
                MSComm1(i).PortOpen = True
            End If
            
            MSComm1(i).InBufferCount = 0
            MSComm1(i).InputLen = 0
        Next
    End If
    
    '=========================read Program List from DB==============================
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
        
        'Debug.Print ff.Name
            
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
    'Debug.Print MPTesterRS.Fields(1)
    
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
    'Label1.Caption = Label1.Caption & Mid(LastDateCode, 1, 4) & "/" & Mid(LastDateCode, 5, 2) & "/" & Mid(LastDateCode, 7, 2)
   
    PreviousChipName = ""

    '============= FOR PCI_7248 ==============
    If AllenDebug = 0 Then
        card = Register_Card(PCI_7296, 0)
        Call SetTimer_1ms
        'SettingForm.Show 1
        If card < 0 Then
            MsgBox "Register Card Failed"
        End If
        Card_Initial
    End If
    
    Call UseExtChipName
    
    If AllenDebug = 1 Then
        ChipNameCombo = "AU6350"
        Call ChipNameCombo_Click
        ChipNameCombo2 = "AL_4Port"
        Call ChipNameCombo2_Click
        DebugForm.Show
    End If

End Sub

Sub UseExtChipName()
Dim tmp1 As String
Dim tmp2 As String
Dim HPC As Integer

    If Dir(App.Path & "\ALCOR.PC") = "ALCOR.PC" Then
        Me.BackColor = &HC0E0FF
        ALCOR = True
        tmp1 = Dir(App.Path & "\*.HPC")
        tmp2 = Dir
        
        If tmp1 <> "" And tmp2 = "" Then
            ALCORChipName = Left(tmp1, Len(tmp1) - 4)
            ChipNameCombo.Text = Left(ALCORChipName, 6)
            Call ChipNameCombo_Click
            Call ChipNameCombo2_Click
        Else
            MsgBox " *.HPC file error"
            End
        End If
    Else
        Me.BackColor = &H8000000F
    End If

End Sub

Private Sub Hander_NS_Click()
    Blank_Timer.Enabled = False
    Label1.BackColor = System_Color
    Hander_Blank.BackColor = System_Color
    Hander_NS.BackColor = System_Color
    Hander_SRM.BackColor = System_Color
End Sub

Private Sub Hander_SRM_Click()
    Blank_Timer.Enabled = False
    Label1.BackColor = System_Color
    Hander_Blank.BackColor = System_Color
    Hander_NS.BackColor = System_Color
    Hander_SRM.BackColor = System_Color
End Sub

Private Sub Report_Click()
    Dim hwndMenu As Long
    Dim c As Long
    hwndMenu = GetSystemMenu(ReportForm.hwnd, 0)
    Call GetMenuItemCount(hwndMenu)
    DeleteMenu hwndMenu, SC_CLOSE, MF_REMOVE
    ReportForm.Show
End Sub

Private Sub StopBtn_Click()
On Error Resume Next

    If (MsgBox("Stop Test ?", vbYesNo + vbQuestion + vbDefaultButton2, "Comform Stop") = vbNo) Then
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
    Hander_NS.Enabled = False
    Hander_SRM.Enabled = False
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
    Hander_NS.Enabled = True
    Hander_SRM.Enabled = True
End Sub

