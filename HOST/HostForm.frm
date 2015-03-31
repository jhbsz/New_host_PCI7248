VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form HostForm 
   BackColor       =   &H80000016&
   Caption         =   "Alcor Host V1.36 "
   ClientHeight    =   10290
   ClientLeft      =   1380
   ClientTop       =   645
   ClientWidth     =   11490
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   11490
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8160
      TabIndex        =   83
      Text            =   "Combo2"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Start"
      Height          =   615
      Left            =   9240
      TabIndex        =   82
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reset"
      Height          =   615
      Left            =   10080
      TabIndex        =   77
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Old Card測試主控程式"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      TabIndex        =   65
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "New Interface control program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      TabIndex        =   64
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "clean counter"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   61
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      Caption         =   "BinOutTest Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   58
      Top             =   9960
      Width           =   2175
   End
   Begin VB.CommandButton ShowBin 
      Caption         =   "Show Bin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   57
      Top             =   7560
      Width           =   1815
   End
   Begin VB.TextBox Text26 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   52
      Text            =   "Text26"
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox Text25 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   51
      Text            =   "Text25"
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox Text24 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   49
      Text            =   "Text24"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text23 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4920
      TabIndex        =   48
      Text            =   "Text23"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text22 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   47
      Text            =   "Text22"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text21 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   46
      Text            =   "Text21"
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text20 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      TabIndex        =   45
      Text            =   "Text20"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text19 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   44
      Text            =   "Text19"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   43
      Text            =   "Text18"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text17 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   42
      Text            =   "Text17"
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6480
      TabIndex        =   39
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "離線模式"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   35
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   34
      Text            =   "Text16"
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   33
      Text            =   "Text14"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   32
      Text            =   "Text13"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3840
      TabIndex        =   31
      Text            =   "Text12"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   30
      Text            =   "Text11"
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   29
      Text            =   "Text6"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   23
      Text            =   "Text15"
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   22
      Text            =   "Text10"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   20
      Text            =   "Text9"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   19
      Text            =   "Text8"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1680
      TabIndex        =   18
      Text            =   "Text7"
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4560
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   0
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   480
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   480
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command3 
      Caption         =   "關閉程式"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   8
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "停止測試"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   7
      Top             =   8040
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   6720
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "(2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   6360
      TabIndex        =   67
      Top             =   5160
      Width           =   2535
      Begin VB.CheckBox LoopTestCheck 
         Caption         =   "LoopTest "
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   84
         Top             =   4800
         Width           =   1815
      End
      Begin VB.CheckBox ReportCheck 
         Caption         =   "報表"
         Height          =   255
         Left            =   480
         TabIndex        =   81
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CheckBox Check7 
         Caption         =   "5  個fail reset "
         Height          =   375
         Left            =   480
         TabIndex        =   78
         Top             =   4080
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "雙機"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   76
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "1號機"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   75
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "2號機"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   74
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "連續供電"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   73
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         Caption         =   "不RT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   72
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CheckBox Check4 
         Caption         =   "One Cycle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   71
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CheckBox Check5 
         Caption         =   "連續fail警告"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   70
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Check_GPON7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   69
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CheckBox NoCardTest 
         Caption         =   "No Card Test"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   68
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label LoopTest 
         Caption         =   "Label35"
         Height          =   255
         Left            =   720
         TabIndex        =   85
         Top             =   4800
         Width           =   1215
      End
   End
   Begin VB.Label Label36 
      BackColor       =   &H0080FF80&
      Caption         =   "Label36"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   90
      Top             =   9240
      Width           =   4095
   End
   Begin VB.Label Label26 
      BackColor       =   &H0080FF80&
      Caption         =   "Label26"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   89
      Top             =   4440
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label35 
      BackColor       =   &H0080FF80&
      Caption         =   "Label35"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   88
      Top             =   4800
      Width           =   4215
   End
   Begin VB.Label Site2GPIB_Label 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   87
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Site1GPIB_Label 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   86
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label34 
      BackColor       =   &H0080FF80&
      Caption         =   "Label34"
      Height          =   375
      Left            =   1680
      TabIndex        =   80
      Top             =   8760
      Width           =   4095
   End
   Begin VB.Label Label33 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label33"
      Height          =   375
      Left            =   1680
      TabIndex        =   79
      Top             =   8280
      Width           =   4095
   End
   Begin VB.Line Line4 
      X1              =   1560
      X2              =   5880
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Label Label32 
      Caption         =   "Label32"
      Height          =   495
      Left            =   120
      TabIndex        =   66
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label31 
      Caption         =   "Label31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   63
      Top             =   8760
      Width           =   495
   End
   Begin VB.Label Label30 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   62
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Label Label29 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   60
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Label Label28 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   59
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Label Label27 
      BackColor       =   &H0080FF80&
      Caption         =   "Label27"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   56
      Top             =   4440
      Width           =   4215
   End
   Begin VB.Label Label25 
      BackColor       =   &H000080FF&
      Caption         =   "Label25"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   55
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Label Label24 
      BackColor       =   &H000080FF&
      Caption         =   "Label24"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   54
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label Label21 
      BackColor       =   &H000080FF&
      Caption         =   "Label21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   53
      Top             =   3360
      Width           =   4215
   End
   Begin VB.Label Label20 
      Caption         =   "ReTest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   50
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label23 
      Caption         =   "(3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   41
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label22 
      Caption         =   "(1) .選擇 IC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   40
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label19 
      BackColor       =   &H008080FF&
      Caption         =   "Label19"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   38
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label Label18 
      BackColor       =   &H008080FF&
      Caption         =   "Label18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   37
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Label Label17 
      BackColor       =   &H008080FF&
      Caption         =   "Label17"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   36
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   5880
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label16 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3840
      TabIndex        =   28
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label15 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3840
      TabIndex        =   27
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Line Line5 
      X1              =   5880
      X2              =   5880
      Y1              =   10440
      Y2              =   0
   End
   Begin VB.Label Label14 
      Caption         =   "2 號 機"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   26
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "1 號 機"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   25
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Line Line3 
      X1              =   1560
      X2              =   1560
      Y1              =   2040
      Y2              =   10440
   End
   Begin VB.Line Line2 
      X1              =   3720
      X2              =   3720
      Y1              =   2040
      Y2              =   8160
   End
   Begin VB.Line Line1 
      X1              =   -120
      X2              =   5880
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label13 
      Caption         =   "測試愈時數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "等Start 愈時次數"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   21
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Fail  數"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Pass  數"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "等待 Start 次數"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "測試時間"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "等 Start 時間"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label coun 
      Caption         =   "測試數"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "output data tmp(0)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "input data tmp1(1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BackColor       =   &H0000FFFF&
      Caption         =   "DUAL SITE HOST "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6480
      TabIndex        =   1
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "HostForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public AU6254Msg As String

Public GreaTekFlag As Integer
Public SPILFlag As Integer

Public GreaTekChipName As String
Public TimeCounter As Integer
Public TimeCounterBegin As Boolean
Public Present
Dim slot As String
Dim hDevice As Integer
Dim a As Integer
Dim tmp(2)  As Byte
Dim tmp1(2)  As Byte
Dim b As Byte
Public flag As Byte
Dim ret As Integer
Dim ret1 As Integer
Private Declare Function OpenLinkDevice Lib "DevLink.dll" (ByVal hInst As Integer, ByRef pHandle As Integer) As Byte
Private Declare Function WriteIOData Lib "DevLink.dll" (ByVal hDevice As Integer, ByRef pSrc As Byte, ByVal dwSize As Integer, ByRef dwRet As Integer) As Byte
Private Declare Function ReadIOData Lib "DevLink.dll" (ByVal hDevice As Integer, ByRef pSrc As Byte, ByVal dwSize As Integer, ByRef dwRet As Integer) As Byte
Public NBCount As Long
'Dim TestMode As Integer

Public WAIT_START_TIME_OUT  As Single  ' wait start signal time out condition
Public WAIT_TEST_CYCLE_OUT As Single    ' wait test time out cycle time
Public POWER_ON_TIME As Single
Public RT_INTERVAL As Single
Public UNLOAD_DRIVER As Single
Public CAPACTOR_CHARGE As Single
Public NO_CARD_TEST_TIME As Single
Public Need_GPIB As Byte
Public site1 As Byte
Public site2 As Byte
Public PowerStatus As Byte
Public TesterStatus1
Public TesterStatus2
Const AlarmLimit = 5
                                     
Public buf1
Public buf2
Public TestMode As Byte
                                     
Public FirstRun As Byte
Public WaitForReady ' timer
                                     
Public TesterReady1 As Byte
Public TesterReady2 As Byte
                                   
Public ResetCounter1 As Byte
Public ResetCounter2 As Byte
                              
Public TesterDownCount1 As Byte
Public TesterDownCount2 As Byte
                                  
Public WaitForPowerOn1   ' timer
Public WaitForPowerOn2
                                    
Public TesterDownCountTimer1  ' timer
Public TesterDownCountTimer2
                                  
                                  
Public WaitForStart
Public WaitForVcc
                                   
Public TotalRealTestTime
Public OldTotalRealTestTime
                                   
Public RealTestTime
Public OldRealTestTime
                                
                                
Public WaitStartTime
                                   
Public GetStart As Integer
Public TimeOut As Integer
                                  
Public WaitStartCounter As Integer
Public WaitStartTimeOutCounter As Integer
Public WaitStartTimeOut As Integer
                                   
                                   
Public TestCounter As Integer
Public RTTestCounter1 As Integer
Public RTTestCounter2 As Integer

Public OffLTestCounter As Integer
Public OffLRTTestCounter1 As Integer
Public OffLRTTestCounter2 As Integer
                                    
      'Test Result
Public TestResult1 As String
Public TestResult2 As String
Public OldTestResult1 As String
Public OldTestResult2 As String
                                   
Public RTTestResult1
Public RTTestResult2
                                  
Public NoCardTestResult1 As String
Public NoCardTestResult2 As String
                                    
     ' Stop Flag
                                    
Public TestStop1 As Byte
Public TestStop2 As Byte
Public TestStop1_1 As Byte
Public TestStop2_1 As Byte
Public TestStop1_2 As Byte
Public TestStop2_2 As Byte
Public TestStop1_3 As Byte
Public TestStop2_3 As Byte
                                    
Public RTTestStop1 As Byte
Public RTTestStop2 As Byte
                                  
                                  
Public NoCardTestStop1 As Byte
Public NoCardTestStop2 As Byte
                                   
      ' Wait for Test Time
                                   
                                   
Public WaitForTest1
Public WaitForTest2
                                     
                                     
Public RTWaitForTest1
Public RTWaitForTest2
                                    
Public NoCardWaitForTest1
Public NoCardWaitForTest2
                                     
                              
                              
                              
Public WaitTestTimeOutCounter1 As Integer
Public WaitTestTimeOutCounter2 As Integer
                                    
Public RTWaitTestTimeOutCounter1 As Integer
Public RTWaitTestTimeOutCounter2 As Integer
                                    
Public WaitTestTimeOut1 As Integer
Public WaitTestTimeOut2 As Integer
                                  
                                  
Public RTWaitTestTimeOut1 As Integer
Public RTWaitTestTimeOut2 As Integer
                                    
       ' Test Cycle
                                    
Public TestCycleTime1
Public TestCycleTime2
                                  
Public RTTestCycleTime1
Public RTTestCycleTime2
                                    
Public NoCardTestCycleTime1
Public NoCardTestCycleTime2
                                 
                                 
Public gpon1 As String
Public gpon2 As String
                                  
Public Bin1Counter1 As Integer
Public Bin2Counter1 As Integer
Public Bin3Counter1 As Integer
Public Bin4Counter1 As Integer
Public Bin5Counter1 As Integer
Public Bin1Counter2 As Integer
Public Bin2Counter2 As Integer
Public Bin3Counter2 As Integer
Public Bin4Counter2 As Integer
Public Bin5Counter2 As Integer
                                    
'Public PassCounter1 As Integer
'Public FailCounter1 As Integer
'Public RTPassCounter1 As Integer
'Public RTFailCounter1 As Integer
                                    
'Public PassCounter2 As Integer
'Public FailCounter2 As Integer
'Public RTPassCounter2 As Integer
'Public RTFailCounter2 As Integer
                                    
Public continuefail1 As Integer
Public continuefail1_bin2 As Integer
Public continuefail1_bin3 As Integer
Public continuefail1_bin4 As Integer
Public continuefail1_bin5 As Integer
Public continuefail2 As Integer
Public continuefail2_bin2 As Integer
Public continuefail2_bin3 As Integer
Public continuefail2_bin4 As Integer
Public continuefail2_bin5 As Integer
                                    
Public T1NotGetReadyCounter As Integer
Public T1UnknownCounter As Integer
Public T1GponFailCounter As Integer
Public T1SD_WFCounter As Integer
Public T1SD_RFCounter As Integer
Public T1CF_WFCounter As Integer
Public T1CF_RFCounter As Integer
Public T1XD_WFCounter As Integer
Public T1XD_RFCounter As Integer
Public T1SM_WFCounter As Integer
Public T1SM_RFCounter As Integer
Public T2NotGetReadyCounter As Integer
Public T2UnknownCounter As Integer
Public T2GponFailCounter As Integer
Public T2SD_WFCounter As Integer
Public T2SD_RFCounter As Integer
Public T2CF_WFCounter As Integer
Public T2CF_RFCounter As Integer
Public T2XD_WFCounter As Integer
Public T2XD_RFCounter As Integer
Public T2SM_WFCounter As Integer
Public T2SM_RFCounter As Integer
Public i As Integer, k As Integer
Public result As Integer
Public DO_P As Long
Public DI_P As Long
Public DI_S1 As Long
Public DI_S2 As Long
Public NewPowerOnTime As Single
Const PCI7248_EOT = &H1 'for 7248 card
Const PCI7248_PASS = &HFD 'for 7248 card  11111101
Const PCI7248_BIN2 = &HFB 'for 7248 card  11111011
Const PCI7248_BIN3 = &HF7 'for 7248 card  11110111
Const PCI7248_BIN4 = &HEF 'for 7248 card  11101111
Const PCI7248_BIN5 = &HDF 'for 7248 card  11011111
Const PCI_7248 = 9

Const EOT1 = &H20 'for old card
Const EOT2 = &H10 'for old card
Const RESET_LATCH = &H3F
Public ChipName As String

Dim value_a(0 To 1) As Long, value_b(0 To 1) As Long, value_cu(0 To 1) As Long, value_cl(0 To 1) As Long
Dim status_a(0 To 1) As Integer, status_b(0 To 1) As Integer, status_cu(0 To 1) As Long, status_cl(0 To 1) As Integer

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

If (ReportCheck.value = 0) Or (Check1.value = 1) Then
Exit Sub
End If

 
If AllenDebug = 1 Then
Bin1Site1 = 2903
Bin2Site1 = 10
Bin3Site1 = 6
Bin4Site1 = 1
Bin5Site1 = 1
Bin1Site2 = 2886
Bin2Site2 = 20
Bin3Site2 = 9
Bin4Site2 = 3
Bin5Site2 = 2
End If
'2. time control
 
   ' Dim EndDay As String
   ' Dim EndSecond As String
   ' Dim SNow As String
   ' Dim OutFileName As String
 '   EndSecond = Format(Now, "HH:MM:SS")
 '   EndDay = Format(Now, "YYYY/MM/DD")
    
    OutFileName = RunCardNO & "_" & ProcessIDSum & "_" & Left(EndDay, 4) & Mid(EndDay, 6, 2) & Right(EndDay, 2)
    
    OutFileName = OutFileName & Left(EndSecond, 2) & Mid(EndSecond, 4, 2) & Right(EndSecond, 2) & "Sum.txt"
    
  '  EndAt = EndDay & Space(1) & EndSecond
  
'3. Summary


Dim TestedSite1 As Long
Dim TestedSite2 As Long
Dim TestedTotal As Long
Dim TestedPercent As Single

Dim PassSite1 As Long
Dim PassSite2 As Long
Dim PassTotal As Long
Dim PassPercent As Single

Dim FailSite1 As Long
Dim FailSite2 As Long
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

' calculate summary


Call GetReportSummarySub2


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
            
            
'=================================================================

            
            
            
            
'=======================================================================
            
    
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
Print #1, "-------------------------------------------------------"
Print #1, Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Total  " & Space(3) & "Total"
Print #1, Space(13) & "COUNT  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percen"
Print #1, "-------------------------------------------------------"


'================ file output

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


'=============== file output

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

Close #1

'=================================================== printer section ===========================
    
'Printer.CurrentX = 300

If AllenDebug = 1 Then
    Exit Sub
End If
 Printer.FontSize = 14
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
 Printer.Print "-------------------------------------------------------"
 Printer.Print Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Total  " & Space(3) & "Total"
 Printer.Print Space(13) & "COUNT  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percen"
 Printer.Print "-------------------------------------------------------"



 
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


    
End Sub

Sub GetReportSummarySub()

Dim oDB As ADOX.Catalog
Dim sDBPAth As String
Dim sConStr As String
Dim oCn As ADODB.Connection
Dim oCM As ADODB.Command
Dim RS As ADODB.Recordset

If AllenDebug = 1 Then
 
RunCardNO = "2008CLOP0040"
End If
'-----------------------------
' set Path and connection string
'---------------------------
sDBPAth = "D:\SLT Summary\Summary.mdb"
'Debug.Print "1"; Dir(sDBPAth, vbNormal + vbDirectory)
If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
    MsgBox "MDB no EXIST"
    Exit Sub
End If

 

'sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDBPAth & ";Persist   Security   Info=False;Jet   OLEDB:Database   Password=058f"
sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\SLT.mdb"
 
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

Dim cmstr As String

cmstr = "SELECT Min(Summary.StartAT) as StartATMin " & _
        "from Summary where RunCardNO= '" & RunCardNO & "'"
oCM.CommandText = cmstr
Debug.Print cmstr
Set RS = oCM.Execute

StartAtMin = RS.Fields("StartATMin")
Debug.Print StartAtMin



 
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
        "from Summary where RunCardNO= '" & RunCardNO & "'"
oCM.CommandText = cmstr
Debug.Print cmstr
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


Debug.Print Bin1Site1Sum
Debug.Print Bin2Site1Sum
Debug.Print Bin3Site1Sum
Debug.Print Bin4Site1Sum
Debug.Print Bin5Site1Sum
Debug.Print Bin1Site2Sum
Debug.Print Bin2Site2Sum
Debug.Print Bin3Site2Sum
Debug.Print Bin4Site2Sum
Debug.Print Bin5Site2Sum

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
Sub GetReportSummarySub2()

Dim oDB As ADOX.Catalog
Dim sDBPAth As String
Dim sConStr As String
Dim oCn As ADODB.Connection
Dim oCM As ADODB.Command
Dim RS As ADODB.Recordset

If AllenDebug = 1 Then
 
RunCardNO = "2008CLOP0040"
End If
'-----------------------------
' set Path and connection string
'---------------------------
sDBPAth = "D:\SLT Summary\Summary.mdb"
'Debug.Print "1"; Dir(sDBPAth, vbNormal + vbDirectory)
If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
    MsgBox "MDB no EXIST"
    Exit Sub
End If

 

'sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDBPAth & ";Persist   Security   Info=False;Jet   OLEDB:Database   Password=058f"
sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\SLT.mdb"
 
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

Dim cmstr As String

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
Debug.Print cmstr
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


Debug.Print Bin1Site1Sum
Debug.Print Bin2Site1Sum
Debug.Print Bin3Site1Sum
Debug.Print Bin4Site1Sum
Debug.Print Bin5Site1Sum
Debug.Print Bin1Site2Sum
Debug.Print Bin2Site2Sum
Debug.Print Bin3Site2Sum
Debug.Print Bin4Site2Sum
Debug.Print Bin5Site2Sum

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
Sub GetProcessIDSub()

Dim oDB As ADOX.Catalog
Dim sDBPAth As String
Dim sConStr As String
Dim oCn As ADODB.Connection
Dim oCM As ADODB.Command
Dim RS As ADODB.Recordset

If AllenDebug = 1 Then
 
RunCardNO = "2008CLOP0040"
End If
'-----------------------------
' set Path and connection string
'---------------------------
sDBPAth = "D:\SLT Summary\Summary.mdb"
'Debug.Print "1"; Dir(sDBPAth, vbNormal + vbDirectory)
If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
    MsgBox "MDB no EXIST"
    Exit Sub
End If

 

'sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDBPAth & ";Persist   Security   Info=False;Jet   OLEDB:Database   Password=058f"
sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\SLT.mdb"
 
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

Dim cmstr As String

cmstr = "SELECT DISTINCT(Summary.ProcessID)   " & _
        "from Summary where RunCardNO= '" & RunCardNO & "'"
oCM.CommandText = cmstr
Debug.Print cmstr
Set RS = oCM.Execute

RS.MoveFirst

'While Not rs.EOF

  'ProcessIDSum = rs("ProcessID")
 ' Debug.Print ProcessIDSum
 ' Call PrintReportSummary2
 ' rs.MoveNext
'Wend


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
Sub PrintReportSummary()
If (ReportCheck.value = 0) Or (Check1.value = 1) Then
Exit Sub
End If

 
If AllenDebug = 1 Then
Bin1Site1 = 2903
Bin2Site1 = 10
Bin3Site1 = 6
Bin4Site1 = 1
Bin5Site1 = 1
Bin1Site2 = 2886
Bin2Site2 = 20
Bin3Site2 = 9
Bin4Site2 = 3
Bin5Site2 = 2
End If
'2. time control
 
   ' Dim EndDay As String
   ' Dim EndSecond As String
   ' Dim SNow As String
   ' Dim OutFileName As String
 '   EndSecond = Format(Now, "HH:MM:SS")
 '   EndDay = Format(Now, "YYYY/MM/DD")
    
    OutFileName = RunCardNO & "_" & ProcessID & "_" & Left(EndDay, 4) & Mid(EndDay, 6, 2) & Right(EndDay, 2)
    
    OutFileName = OutFileName & Left(EndSecond, 2) & Mid(EndSecond, 4, 2) & Right(EndSecond, 2) & "Sum.txt"
    
  '  EndAt = EndDay & Space(1) & EndSecond
  
'3. Summary


Dim TestedSite1 As Long
Dim TestedSite2 As Long
Dim TestedTotal As Long
Dim TestedPercent As Single

Dim PassSite1 As Long
Dim PassSite2 As Long
Dim PassTotal As Long
Dim PassPercent As Single

Dim FailSite1 As Long
Dim FailSite2 As Long
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

' calculate summary


Call GetReportSummarySub


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
            
            
'=================================================================

            
            
            
            
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
Print #1, "Start at: " & StartAtMin
Print #1, "End at: " & EndAt
Print #1, "HandelerID: " & HandlerID
Print #1, "Operator Name: " & OperatorName
Print #1,
Print #1, "-------------------------------------------------------"
Print #1, Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Total  " & Space(3) & "Total"
Print #1, Space(13) & "COUNT  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percen"
Print #1, "-------------------------------------------------------"


'================ file output

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


'=============== file output

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

Close #1

'=================================================== printer section ===========================
    
'Printer.CurrentX = 300

If AllenDebug = 1 Then
      Exit Sub
End If
 Printer.FontSize = 14
 Printer.Font = "標楷體"
 Printer.Print "#####################################################"
 Printer.Print "Name of PC: " & NameofPC
 Printer.Print "Program Name: " & ProgramName
 Printer.Print "Program Rersion Code: " & ProgramRevisionCode
 Printer.Print "Device ID: " & DeviceID
 Printer.Print "Run Card NO: " & RunCardNO
 Printer.Print "Lot ID: " & LotID
 Printer.Print "Process: " & ProcessID
 Printer.Print "Start at: " & StartAtMin
 Printer.Print "End at: " & EndAt
 Printer.Print "HandelerID: " & HandlerID
 Printer.Print "Operator Name: " & OperatorName
 Printer.Print
 Printer.Print "-------------------------------------------------------"
 Printer.Print Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Total  " & Space(3) & "Total"
 Printer.Print Space(13) & "COUNT  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percen"
 Printer.Print "-------------------------------------------------------"



 
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
sDBPAth = "D:\SLT Summary\Summary.mdb"
'Debug.Print "1"; Dir(sDBPAth, vbNormal + vbDirectory)
If Dir(sDBPAth, vbNormal + vbDirectory) = " " Then
    MsgBox "MDB no EXIST"
    Exit Sub
End If

 

'sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDBPAth & ";Persist   Security   Info=False;Jet   OLEDB:Database   Password=058f"
sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\SLT Summary" & "\SLT.mdb"
 
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

'cmstr = "INSERT INTO Summary VALUES(" & _

'oCM.CommandText = "INSERT INTO Summary VALUES(" & _

'cmstr = "INSERT INTO Summary VALUES(" & _

'oCM.CommandText = "UPDATE Summary SET" & _

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


' ------------------------
' Error Handling
' ------------------------
Err_Handler:
'If err <> 0 Then
'err.Clear
'Resume Next
'End If
End Sub


Sub ReportActive()

Dim winHwnd As Long
If ReportCheck = 1 Then

ReportForm.Show

winHwnd = FindWindow(vbNullString, "報表設定")
                    
                  
SetWindowPos winHwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags



End If

  

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
            
       'Check downstream port1-->(connect hub64 module)
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
Dim ii As Integer


 
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

Sub SPIL()
Dim tmp1 As String
Dim tmp2 As String
Dim HPC As Integer



If Dir("C:\Documents and Settings\User\桌面\SPIL.PC") = "SPIL.PC" Then

Me.BackColor = &HC0E0FF
SPILFlag = 1
'Label32.Caption = "GREATEK"
'Label32.BackColor = &HFF00
   
 

Else
     Me.BackColor = &H8000000F
     

End If



End Sub


Sub Greatek()
Dim tmp1 As String
Dim tmp2 As String
Dim HPC As Integer



'If Dir("C:\Documents and Settings\User\桌面\GREATEK.PC") = "GREATEK.PC" Then
If Dir(App.Path & "\GREATEK.PC") = "GREATEK.PC" Then
Me.BackColor = &HC0E0FF
GreaTekFlag = 1
'Label32.Caption = "GREATEK"
'Label32.BackColor = &HFF00
   
'tmp1 = Dir("C:\Documents and Settings\User\桌面\*.HPC")
tmp1 = Dir(App.Path & "\*.HPC")
tmp2 = Dir
 

  If tmp1 <> "" And tmp2 = "" Then
  
      GreaTekChipName = Left(tmp1, Len(tmp1) - 4)
      
        If InStr(GreaTekChipName, "AU87100") Then
            Combo1.Text = Left(GreaTekChipName, 7)
        Else
            Combo1.Text = Left(GreaTekChipName, 6)
        End If
      'Combo2.Text = Mid(GreaTekChipName, 7, Len(GreaTekChipName))
      Call Combo1_Click
      Call Combo2_Click
   Else
      MsgBox "  *.HPC File Error"
      End
   End If
   
   

Else
     Me.BackColor = &H8000000F
     

End If



End Sub

Private Sub AU6375AS_BT_Click()
Dim PowerStatus As Byte
Dim TesterStatus1
Dim TesterStatus2
Const AlarmLimit = 5   'Allen 20050607

Dim buf1
Dim buf2
Dim buf11
Dim buf21
Dim TestMode As Byte

Dim FirstRun As Byte
Dim WaitForReady ' timer

Dim TesterReady1 As Byte
Dim TesterReady2 As Byte

Dim ResetCounter1 As Byte
Dim ResetCounter2 As Byte

Dim TesterDownCount1 As Byte
Dim TesterDownCount2 As Byte

Dim WaitForPowerOn1   ' timer
Dim WaitForPowerOn2
 
Dim TesterDownCountTimer1  ' timer
Dim TesterDownCountTimer2


Dim WaitForStart
Dim WaitForVcc

Dim TotalRealTestTime
Dim OldTotalRealTestTime

Dim RealTestTime
Dim OldRealTestTime


Dim WaitStartTime

Dim GetStart As Integer
Dim TimeOut As Integer
 
Dim WaitStartCounter As Integer
Dim WaitStartTimeOutCounter As Integer
Dim WaitStartTimeOut As Integer


Dim TestCounter As Integer
Dim RTTestCounter1 As Integer
Dim RTTestCounter2 As Integer

'Test Result
Dim TestResult1 As String
Dim TestResult2 As String

Dim RTTestResult1
Dim RTTestResult2

Dim NoCardTestResult1 As String
Dim NoCardTestResult2 As String

' Stop Flag

Dim TestStop1 As Byte
Dim TestStop2 As Byte

Dim RTTestStop1 As Byte
Dim RTTestStop2 As Byte


Dim NoCardTestStop1 As Byte
Dim NoCardTestStop2 As Byte

' Wait for Test Time


Dim WaitForTest1
Dim WaitForTest2


Dim RTWaitForTest1
Dim RTWaitForTest2

Dim NoCardWaitForTest1
Dim NoCardWaitForTest2

Dim WaitTestTimeOutCounter1 As Integer
Dim WaitTestTimeOutCounter2 As Integer

Dim RTWaitTestTimeOutCounter1 As Integer
Dim RTWaitTestTimeOutCounter2 As Integer

Dim WaitTestTimeOut1 As Integer
Dim WaitTestTimeOut2 As Integer


Dim RTWaitTestTimeOut1 As Integer
Dim RTWaitTestTimeOut2 As Integer

' Test Cycle

Dim TestCycleTime1
Dim TestCycleTime2

Dim RTTestCycleTime1
Dim RTTestCycleTime2

Dim NoCardTestCycleTime1
Dim NoCardTestCycleTime2

Dim gpon1 As String
Dim gpon2 As String



Dim PassCounter1 As Integer
Dim FailCounter1 As Integer
Dim RTPassCounter1 As Integer
Dim RTFailCounter1 As Integer

Dim PassCounter2 As Integer
Dim FailCounter2 As Integer
Dim RTPassCounter2 As Integer
Dim RTFailCounter2 As Integer

Dim continuefail1 As Integer
Dim continuefail1_bin2 As Integer
Dim continuefail1_bin3 As Integer
Dim continuefail1_bin4 As Integer
Dim continuefail1_bin5 As Integer
Dim continuefail2 As Integer
Dim continuefail2_bin2 As Integer
Dim continuefail2_bin3 As Integer
Dim continuefail2_bin4 As Integer
Dim continuefail2_bin5 As Integer

Dim T1NotGetReadyCounter As Integer
Dim T1UnknownCounter As Integer
Dim T1GponFailCounter As Integer
Dim T1SD_WFCounter As Integer
Dim T1SD_RFCounter As Integer
Dim T1CF_WFCounter As Integer
Dim T1CF_RFCounter As Integer
Dim T1XD_WFCounter As Integer
Dim T1XD_RFCounter As Integer
Dim T1SM_WFCounter As Integer
Dim T1SM_RFCounter As Integer
Dim T2NotGetReadyCounter As Integer
Dim T2UnknownCounter As Integer
Dim T2GponFailCounter As Integer
Dim T2SD_WFCounter As Integer
Dim T2SD_RFCounter As Integer
Dim T2CF_WFCounter As Integer
Dim T2CF_RFCounter As Integer
Dim T2XD_WFCounter As Integer
Dim T2XD_RFCounter As Integer
Dim T2SM_WFCounter As Integer
Dim T2SM_RFCounter As Integer
Dim i As Integer, k As Integer
Dim result As Integer
Dim DO_P As Long
Dim DI_P As Long
Dim NewPowerOnTime As Single
Dim AU6375LoopCounter As Integer



Dim debug1 As String
'====================================='設定IO PORT輸入輸出

result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
result = DIO_PortConfig(card, Channel_P1CH, INPUT_PORT)
result = DIO_PortConfig(card, Channel_P1CL, INPUT_PORT)

result = DIO_PortConfig(card, Channel_P2A, OUTPUT_PORT)
result = DIO_PortConfig(card, Channel_P2B, OUTPUT_PORT)
result = DIO_PortConfig(card, Channel_P2CH, OUTPUT_PORT)
result = DIO_PortConfig(card, Channel_P2CL, OUTPUT_PORT)


'=====================================

continuefail1 = 0
continuefail1_bin2 = 0
continuefail1_bin3 = 0
continuefail1_bin4 = 0
continuefail1_bin5 = 0
continuefail2 = 0
continuefail2_bin2 = 0
continuefail2_bin3 = 0
continuefail2_bin4 = 0
continuefail2_bin5 = 0

flag = 1
    
Command2.SetFocus



Cls
'=============================================' begin state
Print "begin state"



'///////////////////////////////////////////////////////////////
'
'                     MAIN LOOP
'
'///////////////////////////////////////////////////////////////
Do
Cls
    If ChipName = "" Then
        Label12.BackColor = &H8000000F
        Label14.BackColor = &H8000000F
        Cls
        MsgBox "Select Chip"
        Exit Sub
    End If
    
    
    If Option1.value = True Then '雙機
        site1 = 1
        site2 = 1
        Label12.BackColor = &H8080FF
        Label14.BackColor = &H8080FF
    End If
    
    
    If Option2.value = True Then '1號機
        site1 = 1
        site2 = 0
        Label12.BackColor = &H8080FF
       Label14.BackColor = &H8000000F
    End If
    
    
    If Option3.value = True Then '2號機
        site1 = 0
        site2 = 1
        Label12.BackColor = &H8000000F
        Label14.BackColor = &H8080FF
    End If

   
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Display status
   
    TestResult1 = ""
    TestResult2 = ""
    
    RTTestResult1 = ""
    RTTestResult2 = ""
    
    
    NoCardTestResult1 = ""
    NoCardTestResult2 = ""
    
    'Text1.Text = ""
    'Text2.Text = ""
    Text3.Text = WaitStartCounter
    Text4.Text = WaitStartTime
    Text5.Text = WaitStartTimeOutCounter
    Text6.Text = TestCounter
    Text25.Text = RTTestCounter1
    Text26.Text = RTTestCounter2
    
    If Check1.value = 1 Then
        TestMode = 1  '離線模式
        Text3.Text = "TestMode"
        Text4.Text = "TestMode"
        Text5.Text = "TestMode"
        Text6.Text = "TestMode"
    
    Else
        TestMode = 0  '上線模式
        Text3.Text = WaitStartCounter
        Text4.Text = WaitStartTime
        Text5.Text = WaitStartTimeOutCounter
        Text6.Text = TestCounter
    
    End If

     
    Text7.Text = TestCycleTime1
    Text8.Text = PassCounter1
    Text9.Text = FailCounter1
    Text10.Text = WaitTestTimeOutCounter1
    
    Text17.Text = RTTestCycleTime1
    Text18.Text = RTPassCounter1
    Text19.Text = RTFailCounter1
    Text20.Text = RTWaitTestTimeOutCounter1
    
    Text11.Text = TestCycleTime2
    Text12.Text = PassCounter2
    Text13.Text = FailCounter2
    Text14.Text = WaitTestTimeOutCounter2
   
    Text21.Text = RTTestCycleTime2
    Text22.Text = RTPassCounter2
    Text23.Text = RTFailCounter2
    Text24.Text = RTWaitTestTimeOutCounter2
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Initial Counter and Control variable
    
    
    TestCycleTime1 = 0
    TestCycleTime2 = 0
    
    
    RTTestCycleTime1 = 0
    RTTestCycleTime2 = 0
    
    NoCardTestCycleTime1 = 0
    NoCardTestCycleTime2 = 0
    ' inital time out flag
    GetStart = 0
    WaitStartTime = 0
    WaitStartTimeOut = 0
    
    WaitTestTimeOut1 = 0
    TestStop1 = 0
    NoCardTestStop1 = 0
    WaitTestTimeOut2 = 0
    TestStop2 = 0
    NoCardTestStop2 = 0
    RTTestCycleTime1 = 0
    RTTestCycleTime2 = 0
    
    RTWaitTestTimeOut1 = 0
    RTTestStop1 = 0
    
    RTWaitTestTimeOut2 = 0
    RTTestStop2 = 0
    
    DoEvents
    
        
'*step1=>\\\\\\\\\\\\\\\\\\\\\ Get Start Signal From Handle
'*
'*wait Start Signal From Handle
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'  If ChipName = "AU6375AS" Then   'AU6375As has GPIB to do pwr control
  
'   Call PowerSet(3)
' End If

    WaitForStart = Timer   ' Get Vcc on from Chip
     
    If TestMode = 0 Then  'ON LINE MODE
    
          Print "wait Start"
            Do     'wait  (VCC PowerON) & (hander 5ms start) signal
                   DoEvents
                  WaitStartTime = Timer - WaitForStart
                   k = DO_ReadPort(card, Channel_P1CH, DI_P)
                   'Call MsecDelay(0.1)
        '   Loop Until DI_P = 14 Or DI_P = 13 Or DI_P = 12 Or WaitStartTime > WAIT_START_TIME_OUT  'Allen
           Loop Until DI_P = 14 Or DI_P = 13 Or DI_P = 12
           Label31.Caption = DI_P
            Print "DI_P=", DI_P
            TotalRealTestTime = Timer - OldTotalRealTestTime
            OldTotalRealTestTime = Timer
            OldRealTestTime = Timer
            
     Else
     
            Call MsecDelay(0.8)
            WaitStartTime = 0.8
            DI_P = 14
            TotalRealTestTime = Timer - OldTotalRealTestTime
            OldTotalRealTestTime = Timer
            OldRealTestTime = Timer
            
            'k = DO_WritePort(card, Channel_P1A, 255)
            'k = DO_WritePort(card, Channel_P1B, 255)
            
     End If
     
     
     
    buf1 = MSComm1.Input
    Label26.Caption = "實際總測試時間(含 load / unload) :" & TotalRealTestTime & "s"
    WaitStartCounter = WaitStartCounter + 1
              
    If WaitStartTime > WAIT_START_TIME_OUT Then
         WaitStartTimeOut = 1
         WaitStartTimeOutCounter = WaitStartTimeOutCounter + 1
    End If
      
    
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\  Debug

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'
'    SHOW Alarm
'
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    If Check5.value = 1 Then 'arch change 940529
        If site1 = 1 And continuefail1 >= AlarmLimit Then
        
             If continuefail1_bin2 >= 3 Then
             Call MsecDelay(3)
             End If
             
             
             
        
            If continuefail1_bin2 >= AlarmLimit Then
                Alarm.Show
                Alarm.Label1 = "site1 countiue fail please check Chip Contact and Tester Driver!"
              '  MsgBox "site1 countiue fail please check Chip Contact and Tester Driver!"
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 1
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                    Do
                      DoEvents
                      If AlarmCtrl = 0 Then
                         Exit Do
                      End If
                    Loop While (1)
                
                Print "Alarm Clear"
               
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 1
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                continuefail1_bin2 = 0
                continuefail1 = 0
                
            ElseIf continuefail1_bin3 >= AlarmLimit Then
                Alarm.Show
                Alarm.Label1 = "site1 countiue fail please check  Flash & CF & SD CARD!"
            
               ' MsgBox "site1 countiue fail please check  Flash & CF & SD CARD!"
                
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 2
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                Do
                  DoEvents
                  If AlarmCtrl = 0 Then
                     Exit Do
                  End If
                Loop While (1)
                Print "Alarm Clear"
               
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 2
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                continuefail1_bin3 = 0
                continuefail1 = 0
                
            ElseIf continuefail1_bin4 >= AlarmLimit Then
            
                Alarm.Show
                Alarm.Label1 = "site1 countiue fail please check XD CARD!"
                'MsgBox "site1 countiue fail please check XD CARD!"
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 3
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                Do
                  DoEvents
                  If AlarmCtrl = 0 Then
                     Exit Do
                  End If
                Loop While (1)
                Print "Alarm Clear"
               
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 3
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                continuefail1_bin4 = 0
                continuefail1 = 0
                
            ElseIf continuefail1_bin5 >= AlarmLimit Then
                Alarm.Show
                Alarm.Label1 = "site1 countiue fail please check MS CARD!"
               ' MsgBox "site1 countiue fail please check MS CARD!"
               
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 4
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                Do
                  DoEvents
                  If AlarmCtrl = 0 Then
                     Exit Do
                  End If
                Loop While (1)
                Print "Alarm Clear"
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 4
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                continuefail1_bin5 = 0
                continuefail1 = 0
                
            Else
            
                Print "Site1 check continuefail start !"
                
            End If
        End If
        
        If site2 = 1 And continuefail2 >= AlarmLimit Then
        
        
          If continuefail2_bin2 >= 3 Then
             Call MsecDelay(3)
          End If
             
            If continuefail2_bin2 >= AlarmLimit Then
            
                Alarm.Show
                Alarm.Label1 = "site2 countiue fail please check Chip Contact and Tester Driver!"
              '  MsgBox "site2 countiue fail please check Chip Contact and Tester Driver!"
              
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 5
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                Do
                
                  DoEvents
                  
                  If AlarmCtrl = 0 Then
                  
                     Exit Do
                  
                  End If
                  
                Loop While (1)
  
  
               Print "Alarm Clear"
               
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 5
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
              
              
              
              
              
                continuefail2_bin2 = 0
                continuefail2 = 0
            ElseIf continuefail2_bin3 >= AlarmLimit Then
              Alarm.Show
                Alarm.Label1 = "site2 countiue fail please check  Flash & CF & SD CARD!"
                'MsgBox "site2 countiue fail please check  Flash & CF & SD CARD!"
                
                
                   '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 6
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                Do
                
                  DoEvents
                  
                  If AlarmCtrl = 0 Then
                  
                     Exit Do
                  
                  End If
                  
                Loop While (1)
  
  
               Print "Alarm Clear"
               
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 6
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
              
                
                
                
                continuefail2_bin3 = 0
                continuefail2 = 0
            ElseIf continuefail2_bin4 >= AlarmLimit Then
                Alarm.Show
                Alarm.Label1 = "site2 countiue fail please check XD CARD!"
            
             '   MsgBox "site2 countiue fail please check XD CARD!"
             
                     '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 6
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                Do
                
                  DoEvents
                  
                  If AlarmCtrl = 0 Then
                  
                     Exit Do
                  
                  End If
                  
                Loop While (1)
  
  
               Print "Alarm Clear"
               
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 6
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
              
             
             
                continuefail2_bin4 = 0
                continuefail2 = 0
            ElseIf continuefail2_bin5 >= AlarmLimit Then
            
                Alarm.Show
                Alarm.Label1 = "site2 countiue fail please check MS CARD!"
                
                
                           '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 6
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                Do
                
                  DoEvents
                  
                  If AlarmCtrl = 0 Then
                  
                     Exit Do
                  
                  End If
                  
                Loop While (1)
  
  
               Print "Alarm Clear"
               
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 6
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
              
                
              '  MsgBox "site2 countiue fail please check MS CARD!"
                continuefail2_bin5 = 0
                continuefail2 = 0
            Else
                Print "Site2  check continuefail start !!"
            End If
        End If
        
    Else
        Print "on standard test step!"
    End If
   
 '*******************************************************
 '*
 '*   OPEN power
 '*
 '**********************************************************
   
   
     If DI_P < 12 And DI_P >= 15 Then   'Allen 20050607 , change DI_P > 15, to DI_P >= 15
        Print "no start"
        GoTo err
    Else
       Print "get start signal!"
       Call MsecDelay(CAPACTOR_CHARGE)
       Call MsecDelay(UNLOAD_DRIVER)
      k = DO_WritePort(card, Channel_P1A, &H7F) ' send 0111,1111 => Set power" Channel_P1A = 127
      k = DO_WritePort(card, Channel_P1B, &H7F) ' send 0111,1111 => Set power" Channel_P1b = 127
      
      NewPowerOnTime = POWER_ON_TIME - 0.4
      If NewPowerOnTime > 0 Then
        Call MsecDelay(NewPowerOnTime)
      End If
    End If
   
   
   
'*STEP2=> wait tester send ready signal\\\\\\\\\\\\\\
'*
'*  Check Tester Ready Signal
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\







   MSComm2.InBufferCount = 0
   MSComm1.InBufferCount = 0
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
                               
                               If site1 = 1 Then
                                            If TesterReady1 = 0 Then
                                            
                                                  buf1 = MSComm1.Input
                                                  
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
                                
                               If site2 = 1 Then
                                            If TesterReady2 = 0 Then
                                            
                                                  buf2 = MSComm2.Input
                                                  
                                                  TesterStatus2 = TesterStatus2 & buf2
                                                  
                                                  If (InStr(1, TesterStatus2, "Ready") <> 0) Then
                                                            TesterReady2 = 1
                                                  End If
                                            End If
                                            
                               Else
                                  TesterReady2 = 1
                                     
                               End If
                             
                              '===================================
                              ' Reset rountine : condsider Reset fail
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
                                
                                   '=========== Reset  Rountine
                                    
                                    If TesterReady1 = 0 And TesterDownCount1 = 0 And FirstRun = 1 Then ' reset tester1
                                          '============== close module power
                                           ResetCounter1 = ResetCounter1 + 1
                                           k = DO_WritePort(card, Channel_P1A, &HFF) ' send 0111,1111 => Set power" Channel_P1A = 127
                                           k = DO_WritePort(card, Channel_P1B, &HFF) ' send 0111,1111 => Set power" Channel_P1A = 127
                                          '============= Reset PC
                                          
                                          TesterDownCount1 = 1
                                          k = DO_WritePort(card, Channel_P2CH, 0)
                                                  Call MsecDelay(2)
                                                  k = DO_WritePort(card, Channel_P2CH, 15)
                                                  WaitForPowerOn1 = Timer
                                                  
                                                 '============== clear comm buffer
                                                   MSComm1.InBufferCount = 0
                                                   TesterStatus1 = ""
               
                                    End If
                                
                                
                                   If TesterReady2 = 0 And TesterDownCount2 = 0 And FirstRun = 1 Then ' reset tester2
                                         '============== close module power
                                          ResetCounter2 = ResetCounter2 + 1
                                          k = DO_WritePort(card, Channel_P1A, &HFF) ' send 0111,1111 => Set power" Channel_P1A = 127
                                           k = DO_WritePort(card, Channel_P1B, &HFF) ' send 0111,1111 => Set power" Channel_P1A = 127
                                         '============== Reset PC
                                            TesterDownCount2 = 1
                                             k = DO_WritePort(card, Channel_P2CH, 0)
                                            Call MsecDelay(2)
                                            k = DO_WritePort(card, Channel_P2CH, 15)
                                           WaitForPowerOn2 = Timer
                                         '============== clear comm buffer
                                           MSComm2.InBufferCount = 0
                                            TesterStatus2 = ""
                                                  
                                   End If
               
                               End If
                                        
                               '===============================
                               ' screen down count routine
                               '==============================
                               
                               If TesterDownCount1 = 1 Then
                                        Call PowerSet(5)  '==== Power on
  
                                       Call MsecDelay(0.2)
                                 
                               
                                     TesterDownCountTimer1 = Timer - WaitForPowerOn1
                                     Label28.Caption = CInt(TesterDownCountTimer1)
                                     
                                    If TesterReady1 = 1 Then
                                        '====== open module power
                                         k = DO_WritePort(card, Channel_P1A, &H7F) ' send 0111,1111 => Set power" Channel_P1b = 127
                                          k = DO_WritePort(card, Channel_P1B, &H7F)
                                         Call MsecDelay(POWER_ON_TIME)
                                        
                                        '=== clear flag
                                        TesterDownCount1 = 0
                                     End If
                                     
                                     If TesterDownCountTimer1 > 90 Then  'Reset fail
                                         TesterDownCount1 = 0
                                     End If
                                     
                                  
                               End If
                               
                               
                                If TesterDownCount2 = 1 Then
                                
                                        Call PowerSet(5)  '==== Power on
  
                                       Call MsecDelay(0.2)
                                 
                                 
                                     TesterDownCountTimer2 = Timer - WaitForPowerOn2
                                     Label29.Caption = CInt(TesterDownCountTimer2)
                                     
                                      If TesterReady2 = 1 Then
                                         '====== open module power
                                         k = DO_WritePort(card, Channel_P1B, &H7F) ' send 0111,1111 => Set power" Channel_P1b = 127
                                          k = DO_WritePort(card, Channel_P1A, &H7F)
                                         Call MsecDelay(POWER_ON_TIME)
                                        '=== clear flag
                                        TesterDownCount2 = 0
                                      End If
                                     
                                     
                                        If TesterDownCountTimer2 > 90 Then    ' Reset fail
                                         TesterDownCount2 = 0
                                     End If
                               End If
                                 
                           End If
                               
                      Loop Until (TesterReady1 = 1) And (TesterReady2 = 1)
                      FirstRun = 1
         
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'*
'*    Testing Loop
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
If (DI_P >= 12) And (DI_P < 15) Then
  
    
        ' init falg
         GetStart = 1
    Label3.BackColor = RGB(255, 255, 255)
    Label3 = ""
    Label16.BackColor = RGB(255, 255, 255)
    Label16 = ""
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '
    '                Site1 and Site2  begin
    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
             
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '        Testing LED function    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Print "==========================="
          
  
                ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  Allen 0526 begin 1 : for no card test,pull high Card detect
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                   If NoCardTest.value = 1 Then
                        
                   k = DO_WritePort(card, Channel_P2CL, &HF)   'pull High
                   Else
                   
                   k = DO_WritePort(card, Channel_P2CL, &H0)
                   End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  Allen 0526 End  1 : for no card test,pull high Card detect
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
   
    '*STEP4=> Waitting for Response from  Tester\\\\\\\\\\\\\\\\\\\\\
    '*
    '*    Wait Test Result from each Tester
    '*
    '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
   '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      '
      '  Allen 0601 Remark : no card on board test card detect and card change signal
      '
      '
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
       NoCardTestResult1 = ""
       NoCardTestResult2 = ""
        If site1 = 1 And NoCardTest.value = 1 Then
    
        MSComm1.Output = ChipName   ' trans strat test signal to TEST PC
        MSComm1.InBufferCount = 0
        MSComm1.InputLen = 0
        NoCardWaitForTest1 = Timer ' wait for timer  and test result
        End If
    
        
        If site2 = 1 And NoCardTest.value = 1 Then
        
        MSComm2.Output = ChipName   ' trans strat test signal to TEST PC
        MSComm2.InBufferCount = 0
        MSComm2.InputLen = 0
        NoCardWaitForTest2 = Timer
        End If
    
    
         Print "send begin test signal to test"
         TesterStatus1 = ""
         TesterStatus2 = ""
      
      
             Do
                DoEvents
                
                If site1 = 1 And NoCardTest.value = 1 Then
                    If NoCardTestStop1 = 0 Then
                    
                        If MSComm1.InBufferCount >= 4 Then
                            NoCardTestResult1 = MSComm1.Input
                          
                            
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
                '========================
                
                If site2 = 1 And NoCardTest.value = 1 Then
                    If NoCardTestStop2 = 0 Then
                         If MSComm2.InBufferCount >= 4 Then
                                NoCardTestResult2 = MSComm2.Input
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
            
              
            Loop Until (NoCardTestStop1 = 1) And (NoCardTestStop2 = 1)
            
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      '
      '  Allen 0526 Remark : no card on board test card detect and card change signal
      '
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
         '*STEP3=>Send command to PC teser\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '*
    '*    Send ChipName to PC teser
    '*
    '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
AU6375LoopCounter = 0
 
'==========================================
Do                   '== AU6375 sorting loop

Call PowerSet(4)  '==== Power on
Call MsecDelay(0.5)
Cls
Print "Soritng Counter="; AU6375LoopCounter

If AU6375LoopCounter = 0 Then
    NoCardTestResult1 = "PASS"
    NoCardTestResult2 = "PASS"
    TestResult1 = ""
    TestResult2 = ""
   
    TestStop1 = 0
    TestStop2 = 0
Else
    If TestResult1 <> "PASS" Then
    TestStop1 = 1
    TestResult1 = NoCardTestResult1
    Else
    TestResult1 = ""     'pass
     
    TestStop1 = 0
     
    End If
    
    If TestResult2 <> "PASS" Then
    TestStop2 = 1
    TestResult2 = NoCardTestResult2
    Else
   
    TestResult2 = ""   'pass
   
    
    TestStop2 = 0
    End If
    
    
End If



 TestCycleTime1 = 0
    TestCycleTime2 = 0
    
    
    
    
    If NoCardTest.value = 1 Then
    
             k = DO_WritePort(card, Channel_P2CL, 0)   'pull down
             
'             Call MsecDelay(0.1)
            If site1 = 1 Then  '****** Continue condition lock at PC tester
            
                MSComm1.Output = NoCardTestResult1   ' only pass can continue at PC Tester
                MSComm1.InBufferCount = 0
                MSComm1.InputLen = 0
                WaitForTest1 = Timer ' wait for timer  and test result
             
                If NoCardTestResult1 <> "PASS" Then
                    TestResult1 = NoCardTestResult1
                End If
            End If
          
            
            If site2 = 1 Then
                MSComm2.Output = NoCardTestResult2   ' only pass can continue at PC Tester
                MSComm2.InBufferCount = 0
                MSComm2.InputLen = 0
                WaitForTest2 = Timer
                If NoCardTestResult2 <> "PASS" Then
                    TestResult2 = NoCardTestResult2
                End If
            End If
    
    
    
    Else
    
            If site1 = 1 And TestStop1 = 0 Then
               ' NoCardTestResult1 = "PASS"  ' For AU6375 many times Sorint
                MSComm1.Output = ChipName   ' trans strat test signal to TEST PC
                MSComm1.InBufferCount = 0
                MSComm1.InputLen = 0
                WaitForTest1 = Timer ' wait for timer  and test result
            End If
            
            
            If site2 = 1 And TestStop2 = 0 Then
              '  NoCardTestResult2 = "PASS"   ' For AU6375 many times soritng
                MSComm2.Output = ChipName   ' trans strat test signal to TEST PC
                MSComm2.InBufferCount = 0
                MSComm2.InputLen = 0
                WaitForTest2 = Timer
            End If
    End If
    
   
                    Do
                               DoEvents
                               
                               If site1 = 1 And NoCardTestResult1 = "PASS" Then
                                            If TestStop1 = 0 Then
                                            
                                                  If MSComm1.InBufferCount >= 4 Then
                                                    TestResult1 = MSComm1.Input
                                                  
                                                  
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
                               '========================
                               
                               If site2 = 1 And NoCardTestResult2 = "PASS" Then
                                        If TestStop2 = 0 Then
                                        
                                             If MSComm2.InBufferCount >= 4 Then
                                              TestResult2 = MSComm2.Input
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
                             
                               
                      Loop Until (TestStop1 = 1) And (TestStop2 = 1)
       
        
        
    '\\\\\\\\\\\\\\\\\\\\\\\\\\}wait Tester response END
                           
            If site1 = 1 Then
                Print "TestResult1= "; TestResult1
               
            End If
            
            If site2 = 1 Then
                Print "TestResult2= "; TestResult2
                               
            End If
                             
                             
           
          '  TestCounter = TestCounter + 1 ' Allen Debug
        
            If TestCycleTime1 > WAIT_TEST_CYCLE_OUT And site1 = 1 Then
            
              If ChipName = "AU6375AS" Then  ' only for AU6375AS, because AU6375AS is pass by tester timeout
                TestResult1 = "PASS"
              End If
              
           '   WaitTestTimeOut1 = 1
            '  WaitTestTimeOutCounter1 = WaitTestTimeOutCounter1 + 1
            End If
                     
            If TestCycleTime2 > WAIT_TEST_CYCLE_OUT And site2 = 1 Then
            
              If ChipName = "AU6375AS" Then  ' only for AU6375AS, because AU6375AS is pass by tester timeout
              TestResult2 = "PASS"
              End If
             ' WaitTestTimeOut2 = 1
             ' WaitTestTimeOutCounter2 = WaitTestTimeOutCounter2 + 1
            End If
         
'===========================================================================
'
'    AU6375 sorting control loop
'
'===========================================================================
  
  NoCardTestResult1 = TestResult1
  NoCardTestResult2 = TestResult2
  
  AU6375LoopCounter = AU6375LoopCounter + 1
  If site1 = 1 Then
     buf11 = MSComm1.Input
     MSComm1.InBufferCount = 0
  End If
  
    If site2 = 1 Then
    buf21 = MSComm2.Input  ' to clear mscomm buffer
      MSComm2.InBufferCount = 0
    End If
      '======= GPIB 'power off'
      
   
 
      
Call PowerSet(5)  '==== Power on
  
  
Call MsecDelay(0.5)
 Loop While AU6375LoopCounter < 3 And ((TestResult1 = "PASS") Or (TestResult2 = "PASS"))
  
         
                 
       '/////////////////////////////////////////////////////////////////////////
       '
       '   RT Condition
       '
       '//////////////////////////////////////////////////////////////////////////
                 
                 
        If Check3.value = 1 Then   ' 不RT=> not low yield sorting
            GoTo err
        End If
        
        
        If site1 = 1 And site2 = 0 Then
            If TestResult1 = "PASS" Then
                GoTo err
            End If
            
        End If
        
        
        If site1 = 0 And site2 = 1 Then
            If TestResult2 = "PASS" Then
                GoTo err
            End If
            
        End If
        
        If site1 = 1 And site2 = 1 Then
            If TestResult1 = "PASS" And TestResult2 = "PASS" Then
                GoTo err
            End If
            
        End If
        '////////////////////////////// initial condition
           
         Print "RT begin"
                 
         '1.close power
         '2.delay 10 s
         '3.send power
         '4.RT core
         
         Print "close power"
          
            k = DO_WritePort(card, Channel_P1A, &HFF)
            k = DO_WritePort(card, Channel_P1B, &HFF)
        
         
         Call MsecDelay(RT_INTERVAL)  ' to let system unload driver
         
         Print "Send power"
         k = DO_WritePort(card, Channel_P1A, &H7F)
         k = DO_WritePort(card, Channel_P1B, &H7F)
                  
         Call MsecDelay(POWER_ON_TIME)
         
                 
                 
                 
            If site1 = 1 And TestResult1 <> "PASS" Then
                
                MSComm1.Output = ChipName ' trans strat test signal to TEST PC
                MSComm1.InBufferCount = 0
                MSComm1.InputLen = 0
                RTWaitForTest1 = Timer ' wait for timer  and test result
             End If
                
             If site2 = 1 And TestResult2 <> "PASS" Then
                 
                MSComm2.Output = ChipName    ' trans strat test signal to TEST PC
                MSComm2.InBufferCount = 0
                MSComm2.InputLen = 0
                RTWaitForTest2 = Timer
             End If
             
            
             Print "send begin test signal to test"
             
             '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Wait for Response from PC Tester
                
            Do
                DoEvents
                
                If site1 = 1 And TestResult1 <> "PASS" Then
                    If RTTestStop1 = 0 Then
                        RTTestResult1 = MSComm1.Input
                        
                        RTTestCycleTime1 = Timer - RTWaitForTest1
                        
                        If (RTTestResult1 <> "" Or RTTestCycleTime1 > WAIT_TEST_CYCLE_OUT) Then
                        
                            RTTestCounter1 = RTTestCounter1 + 1 ' Allen Debug
                            RTTestStop1 = 1
                        End If
                    End If
                
                Else
                    RTTestStop1 = 1
                
                End If
                '========================
                
                If site2 = 1 And TestResult2 <> "PASS" Then
                    If RTTestStop2 = 0 Then
                        RTTestResult2 = MSComm2.Input
                        
                        RTTestCycleTime2 = Timer - RTWaitForTest2
                        
                        If (RTTestResult2 <> "" Or RTTestCycleTime2 > WAIT_TEST_CYCLE_OUT) Then
                            RTTestStop2 = 1
                            RTTestCounter2 = RTTestCounter2 + 1
                        End If
                    End If
                
                Else
                    RTTestStop2 = 1
                
                
                End If
            
            
            Loop Until (RTTestStop1 = 1) And (RTTestStop2 = 1)
                           
                If site1 = 1 Then
                    Print "RTTestResult1= "; RTTestResult1
                End If
                
                If site2 = 1 Then
                    Print "RTTestResult2= "; RTTestResult2
                End If
                
               
                
                If RTTestCycleTime1 > WAIT_TEST_CYCLE_OUT And site1 = 1 Then
                
                    RTWaitTestTimeOut1 = 1
                    RTWaitTestTimeOutCounter1 = RTWaitTestTimeOutCounter1 + 1
                End If
                
                If RTTestCycleTime2 > WAIT_TEST_CYCLE_OUT And site2 = 1 Then
                
                    RTWaitTestTimeOut2 = 1
                    RTWaitTestTimeOutCounter2 = RTWaitTestTimeOutCounter2 + 1
                End If
                 
                 
                 
                     
                 
End If  '////////////////////// Test end
              '  Testing Loop end
                 
   
   
                
'////////////////////////////////////// Tester 1 PASS Bin

err:
    ' default value
       gpon1 = "PASS"
       gpon2 = "PASS"
      
      
    If Check6.value = 1 Then     'check GPON7_LED & Power_LED
    
       gpon1 = ""
       gpon2 = ""
    
        Dim DI_Power As Long
        k = DO_ReadPort(card, Channel_P1CL, DI_Power)
        Print "DI_Power="; DI_Power
        
        If TestResult1 = "PASS" Then
            If site1 = 1 And (DI_Power Mod 4 = 0) Then
                gpon1 = "PASS"
                Print "gpon1="; gpon1
            Else
                gpon1 = "FAIL"
                Print "gpon1="; gpon1
                 TestResult1 = "gponFail"
            End If
        End If
        
        If TestResult2 = "PASS" Then
            If site2 = 1 And (DI_Power <= 3) Then
                gpon2 = "PASS"
                Print "gpon2="; gpon2
            Else
                gpon2 = "FAIL"
                Print "gpon2="; gpon2
                 TestResult2 = "gponFail"
            End If
        End If
        
        
     
    
        
    End If
    

    

Label3 = TestResult1
Label16 = TestResult2

Print "close power"

If Check2.value = 0 Then  'continuous supply power
    k = DO_WritePort(card, Channel_P1A, &HFF)
    k = DO_WritePort(card, Channel_P1B, &HFF)
End If

If site1 = 1 Then
 


 Select Case TestResult1
        Case "PASS"
            TestResult1 = "PASS"
        Case "UNKNOW", "bin2", "Bin2"
            TestResult1 = "bin2"
            T1UnknownCounter = T1UnknownCounter + 1
            BinForm.Text1.Text = T1UnknownCounter
        Case "gponFail", "bin3", "Bin3"
            TestResult1 = "bin3"
            T1GponFailCounter = T1GponFailCounter + 1
            BinForm.Text29.Text = T1GponFailCounter
        Case "SD_WF"
            TestResult1 = "bin3"
            T1SD_WFCounter = T1SD_WFCounter + 1
            BinForm.Text2.Text = T1SD_WFCounter
        Case "SD_RF"
            TestResult1 = "bin3"
            T1SD_RFCounter = T1SD_RFCounter + 1
            BinForm.Text3.Text = T1SD_RFCounter
        Case "CF_WF"
            TestResult1 = "bin3"
            T1CF_WFCounter = T1CF_WFCounter + 1
            BinForm.Text4.Text = T1CF_WFCounter
        Case "CF_RF"
            TestResult1 = "bin3"
            T1CF_RFCounter = T1CF_RFCounter + 1
            BinForm.Text5.Text = T1CF_RFCounter
        Case "XD_WF", "bin4", "Bin4"
            TestResult1 = "bin4"
            T1XD_WFCounter = T1XD_WFCounter + 1
            BinForm.Text6.Text = T1XD_WFCounter
        Case "XD_RF"
            TestResult1 = "bin4"
            T1XD_RFCounter = T1XD_RFCounter + 1
            BinForm.Text7.Text = T1XD_RFCounter
        Case "MS_WF", "bin5", "TimeOut", "Bin5"
            TestResult1 = "bin5"
            T1SM_WFCounter = T1SM_WFCounter + 1
             BinForm.Text8.Text = T1SM_WFCounter
        Case "MS_RF"
            TestResult1 = "bin5"
            T1SM_RFCounter = T1SM_RFCounter + 1
            BinForm.Text9.Text = T1SM_RFCounter
        Case Else
            TestResult1 = "bin2"
            T1UnknownCounter = T1UnknownCounter + 1
            BinForm.Text1.Text = T1UnknownCounter
        End Select

End If
If site2 = 1 Then


Select Case TestResult2
        Case "PASS"
            TestResult2 = "PASS"
        Case "UNKNOW", "bin2", "Bin2"
            TestResult2 = "bin2"
            T2UnknownCounter = T2UnknownCounter + 1
            BinForm.Text15.Text = T2UnknownCounter
        Case "gponFail", "bin3", "Bin3"
            TestResult2 = "bin3"
            T2GponFailCounter = T2GponFailCounter + 1
            BinForm.Text30.Text = T2GponFailCounter
        Case "SD_WF"
            TestResult2 = "bin3"
            T2SD_WFCounter = T2SD_WFCounter + 1
            BinForm.Text16.Text = T2SD_WFCounter
        Case "SD_RF"
            TestResult2 = "bin3"
            T2SD_RFCounter = T2SD_RFCounter + 1
            BinForm.Text17.Text = T2SD_RFCounter
        Case "CF_WF"
            TestResult2 = "bin3"
            T2CF_WFCounter = T2CF_WFCounter + 1
            BinForm.Text18.Text = T2CF_WFCounter
        Case "CF_RF"
            TestResult2 = "bin3"
            T2CF_RFCounter = T2CF_RFCounter + 1
            BinForm.Text19.Text = T1CF_RFCounter
        Case "XD_WF", "bin4", "Bin4"
            TestResult2 = "bin4"
            T2XD_WFCounter = T2XD_WFCounter + 1
            BinForm.Text20.Text = T2XD_WFCounter
        Case "XD_RF"
            TestResult2 = "bin4"
            T2XD_RFCounter = T2XD_RFCounter + 1
            BinForm.Text21.Text = T2XD_RFCounter
        Case "MS_WF", "bin5", "TimeOut", "Bin5"
            TestResult2 = "bin5"
            T2SM_WFCounter = T2SM_WFCounter + 1
             BinForm.Text22.Text = T2SM_WFCounter
        Case "MS_RF"
            TestResult2 = "bin5"
            T2SM_RFCounter = T2SM_RFCounter + 1
            BinForm.Text23.Text = T2SM_RFCounter
        Case Else
            TestResult2 = "bin2"
            T2UnknownCounter = T2UnknownCounter + 1
            BinForm.Text15.Text = T2UnknownCounter
        End Select

End If


 RealTestTime = Timer - OldRealTestTime
 If SPILFlag Then
    Label27.Caption = "Real Test Time :" & RealTestTime & " s"
 Else
    Label27.Caption = "實際測試時間(不含 load / unload)  :" & RealTestTime & " s"
 End If
 
If (TestResult1 = "PASS" Or RTTestResult1 = "PASS") And GetStart = 1 And site1 = 1 And gpon1 = "PASS" Then

        Print "\\\\\\\\\\site1 = "; TestResult1
              
        If TestResult1 = "PASS" Then
            PassCounter1 = PassCounter1 + 1
            
            Label2 = "PASS1"
            Label3.BackColor = RGB(0, 255, 0)
        Else
            RTPassCounter1 = RTPassCounter1 + 1
            Print "-------------- RTPASS 1"
            Label2 = "RTPASS1"
            Label3.BackColor = RGB(0, 0, 255)
        End If
        
        Bin1Counter1 = Bin1Counter1 + 1
        BinForm.Text10.Text = Bin1Counter1
        
        
        Call PCI7248_bin(Channel_P1B, PCI7248_PASS, Check2.value)
        
       
       
End If
            
'////////////////////////////////////// Tester 1 FAIL Bin
If Check3.value = 0 Then
    If (((TestResult1 <> "PASS" And RTTestResult1 <> "PASS") And GetStart = 1) Or WaitStartTimeOut = 1) And site1 = 1 Then
    
    
         Print "\\\\\\\\\\site1 = "; TestResult1
        
        If WaitStartTimeOut = 0 Then
            FailCounter1 = FailCounter1 + 1
        End If
        
         
        
        'Bin fail
        Select Case TestResult1
        
            Case "bin2" '(site1_bin2 = &H2)+(PowerStatus=&H80)  '10000010
                Bin2Counter1 = Bin2Counter1 + 1
                BinForm.Text11.Text = Bin2Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN2, Check2.value)
            Case "bin3"
                Bin3Counter1 = Bin3Counter1 + 1
                BinForm.Text12.Text = Bin3Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN3, Check2.value)
            Case "bin4"
                Bin4Counter1 = Bin4Counter1 + 1
                BinForm.Text13.Text = Bin4Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN4, Check2.value)
            Case "bin5"
                Bin5Counter1 = Bin5Counter1 + 1
                BinForm.Text14.Text = Bin5Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN5, Check2.value)
        End Select
        
            
            Label2 = "FAIL1"
            Label3.BackColor = RGB(255, 0, 0)
    End If

End If

 '=======================================Tester 1 FAIL Bin 'Check3.Value = 1=>不RT
        
If Check3.value = 1 And TestResult1 <> "PASS" Then
    If (((TestResult1 <> "PASS") And GetStart = 1) Or WaitStartTimeOut = 1) And site1 = 1 Then
         Print "\\\\\\\\\\site1 = "; TestResult1
            If WaitStartTimeOut = 0 Then
                FailCounter1 = FailCounter1 + 1
            End If
        
        Select Case TestResult1
            Case "bin2"
                Bin2Counter1 = Bin2Counter1 + 1
                BinForm.Text11.Text = Bin2Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN2, Check2.value)
            Case "bin3"
                Bin3Counter1 = Bin3Counter1 + 1
                BinForm.Text12.Text = Bin3Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN3, Check2.value)
            Case "bin4"
                Bin4Counter1 = Bin4Counter1 + 1
                BinForm.Text13.Text = Bin4Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN4, Check2.value)
            Case "bin5"
                Bin5Counter1 = Bin5Counter1 + 1
                BinForm.Text14.Text = Bin5Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN5, Check2.value)
                
            Case Else
                Bin2Counter1 = Bin2Counter1 + 1
                BinForm.Text11.Text = Bin2Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN2, Check2.value)
        End Select
            
        
        Label2 = "FAIL1"
        Label3.BackColor = RGB(255, 0, 0)
    End If

End If

   Call MsecDelay(0.2)
            

 '////////////////////////////////////// Tester 2 PASS Bin
        
If (TestResult2 = "PASS" Or RTTestResult2 = "PASS") And GetStart = 1 And site2 = 1 And gpon2 = "PASS" Then

        Print "\\\\\\\\\\site2 = "; TestResult2
        If TestResult2 = "PASS" Then
            PassCounter2 = PassCounter2 + 1
            Print "\\\\\\\\\\\\\\ PASS 2"
            Label15 = "PASS2"
            Label16.BackColor = RGB(0, 255, 0)
        Else
            RTPassCounter2 = RTPassCounter2 + 1
            Print "-------------- RTPASS 2"
            Label15 = "RTPASS2"
            Label16.BackColor = RGB(0, 0, 255)
        End If
        
     
        Bin1Counter2 = Bin1Counter2 + 1
        BinForm.Text24.Text = Bin1Counter2
        Call PCI7248_bin(Channel_P1A, PCI7248_PASS, Check2.value)
        
    
End If
        
        
'======================================= Tester 2 FAIL Bin 'Check3.Value = 0=>要RT
If Check3.value = 0 Then
     If (((TestResult2 <> "PASS" And RTTestResult2 <> "PASS") And GetStart = 1) Or WaitStartTimeOut = 1) And site2 = 1 Then
           
             Print "\\\\\\\\\\site2 = "; TestResult2
             
             If WaitStartTimeOut = 0 Then
                  FailCounter2 = FailCounter2 + 1
              End If
            
            Select Case TestResult2
            
                Case "bin2"
                    Bin2Counter2 = Bin2Counter2 + 1
                    BinForm.Text25.Text = Bin2Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN2, Check2.value)
                Case "bin3"
                    Bin3Counter2 = Bin3Counter2 + 1
                    BinForm.Text26.Text = Bin3Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN3, Check2.value)
                Case "bin4"
                    Bin4Counter2 = Bin4Counter2 + 1
                    BinForm.Text27.Text = Bin4Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN4, Check2.value)
                Case "bin5"
                    Bin5Counter2 = Bin5Counter2 + 1
                    BinForm.Text28.Text = Bin5Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN5, Check2.value)
            End Select
            
              Label15 = "FAIL!"
              Label16.BackColor = RGB(255, 0, 0)
             
                         
    End If
End If
        
'=========================================Tester 2 FAIL Bin 'Check3.Value = 1=>不RT
If Check3.value = 1 And TestResult2 <> "PASS" Then
    If (((TestResult2 <> "PASS") And GetStart = 1) Or WaitStartTimeOut = 1) And site2 = 1 Then
           
             Print "\\\\\\\\\\site2 = "; TestResult2
             
             If WaitStartTimeOut = 0 Then
                  FailCounter2 = FailCounter2 + 1
              End If
            
              
            
            Select Case TestResult2
            
                Case "bin2"
                    Bin2Counter2 = Bin2Counter2 + 1
                    BinForm.Text25.Text = Bin2Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN2, Check2.value)
                Case "bin3"
                    Bin3Counter2 = Bin3Counter2 + 1
                    BinForm.Text26.Text = Bin3Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN3, Check2.value)
                Case "bin4"
                    Bin4Counter2 = Bin4Counter2 + 1
                    BinForm.Text27.Text = Bin4Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN4, Check2.value)
                Case "bin5"
                    Bin5Counter2 = Bin5Counter2 + 1
                    BinForm.Text28.Text = Bin5Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN5, Check2.value)
                Case Else
                    Bin2Counter2 = Bin2Counter2 + 1
                    BinForm.Text25.Text = Bin2Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN2, Check2.value)
                 
            End Select
           
              Label15 = "FAIL!"
              Label16.BackColor = RGB(255, 0, 0)
             
                         
    End If
End If
'===========================================================
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Reset  Latch
                  
        
            
             Print "end state"
                   
              
               
       ' for check5 =============================Arch add 940529
        If (TestResult1 = "PASS" Or RTTestResult1 = "PASS") Then
            continuefail1 = 0
            continuefail1_bin2 = 0
            continuefail1_bin3 = 0
            continuefail1_bin4 = 0
            continuefail1_bin5 = 0
        ElseIf TestResult1 = "bin2" Or RTTestResult1 = "bin2" Then
            continuefail1 = continuefail1 + 1
            continuefail1_bin2 = continuefail1_bin2 + 1
        ElseIf TestResult1 = "bin3" Or RTTestResult1 = "bin3" Then
            continuefail1 = continuefail1 + 1
            continuefail1_bin3 = continuefail1_bin3 + 1
        ElseIf TestResult1 = "bin4" Or RTTestResult1 = "bin4" Then
            continuefail1 = continuefail1 + 1
            continuefail1_bin4 = continuefail1_bin4 + 1
        ElseIf TestResult1 = "bin5" Or RTTestResult1 = "bin5" Then
            continuefail1 = continuefail1 + 1
            continuefail1_bin5 = continuefail1_bin5 + 1
        End If
        
        If (TestResult2 = "PASS" Or RTTestResult2 = "PASS") Then
            continuefail2 = 0
            continuefail2_bin2 = 0
            continuefail2_bin3 = 0
            continuefail2_bin4 = 0
            continuefail2_bin5 = 0
        ElseIf TestResult2 = "bin2" Or RTTestResult2 = "bin2" Then
            continuefail2 = continuefail2 + 1
            continuefail2_bin2 = continuefail2_bin2 + 1
        ElseIf TestResult2 = "bin3" Or RTTestResult2 = "bin3" Then
            continuefail2 = continuefail2 + 1
            continuefail2_bin3 = continuefail2_bin3 + 1
        ElseIf TestResult2 = "bin4" Or RTTestResult2 = "bin4" Then
            continuefail2 = continuefail2 + 1
            continuefail2_bin4 = continuefail2_bin4 + 1
        ElseIf TestResult2 = "bin5" Or RTTestResult2 = "bin5" Then
            continuefail2 = continuefail2 + 1
            continuefail2_bin5 = continuefail2_bin5 + 1
        End If
       
        
'=======================================send testend to test for start loop?
   
    
   ' If Site1 = 1 Then
    
   ' MSComm1.Output = "testend"   ' trans end signal to TEST PC
   ' MSComm1.InBufferCount = 0
   ' MSComm1.InputLen = 0
   ' WaitForTest1 = Timer ' wait for timer  and test result
   ' End If
   ' buf1 = MSComm1.Input
    
   ' If Site2 = 1 Then
    
   ' MSComm2.Output = "testend"   ' trans end signal to TEST PC
   ' MSComm2.InBufferCount = 0
   ' MSComm2.InputLen = 0
   ' WaitForTest2 = Timer
   ' End If
   ' buf2 = MSComm2.Input
    
   ' Print "send end test signal to test"
        
TestEnd:
                   
               If flag = 0 Then      ' stop or one cycle
                    'ChipName = ""
                    'Label17.Caption = ""
                    'Label18.Caption = ""
                    'Label19.Caption = ""
                    'Command1.Enabled = False
                    'Command6.Enabled = False
                    'Combo1.Clear
                    'Combo2.Clear
                    'Label12.BackColor = &H8000000F
                    'Label14.BackColor = &H8000000F
                    'Label3.BackColor = &HFFFFFF
                    'Label16.BackColor = &HFFFFFF
                    Cls
                    Print "STOP TEST!!!"
                   Exit Sub
               End If
               
               
               If Check4.value = 1 Then     ' stop or one cycle
                    
                    'Label12.BackColor = &H8000000F
                    'Label14.BackColor = &H8000000F
                    'Label3.BackColor = &HFFFFFF
                    'Label16.BackColor = &HFFFFFF
                    Exit Sub
                   Print "STOP TEST!!!"
                 '   MsgBox "end state"
                   
               End If
               
Loop  'do1

   ' MSComm1.PortOpen = False '關閉序列埠
 
End Sub


Private Sub Combo1_Change()

If GreaTekFlag = 1 Then
    If InStr(GreaTekChipName, "AU87100") Then
        Combo1.Text = Left(GreaTekChipName, 7)
    Else
        Combo1.Text = Left(GreaTekChipName, 6)
    End If
    
     Combo2.Enabled = True
End If


End Sub

Private Sub Combo1_Click()
 
 'Const Ver = " V2.00"

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

IC_Name = Trim(Combo1)
Combo2.Clear
Combo2.Enabled = False

 
NO_CARD_TEST_TIME = 0

    '=============================================
        'connection to MPTester.mdb
    '=============================================
    MpTesterDB_Path = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\PGM_ListDB\MPTester_" & LastMPTesterDateCode & ".mdb"
    ConnMPTesterDB.Open MpTesterDB_Path
    MPTesterRS.CursorLocation = adUseClient
    MPTesterRS.Open "MPTester", ConnMPTesterDB, adOpenKeyset, adLockPessimistic
    MPTesterRS.MoveFirst
    Set MPTesterRS = ConnMPTesterDB.Execute("Select * From [MPTester] Where [PGM_Name] LIKE '" & IC_Name & "%' Order By [PGM_Name]")
    
    Do Until MPTesterRS.EOF = True
        If MPTesterRS.Fields(0) = 1 Then
            If Len(MPTesterRS.Fields(1)) = 15 Then
                Combo2.AddItem Mid(MPTesterRS.Fields(1), 8, Len(MPTesterRS.Fields(1)))
            Else
                Combo2.AddItem Mid(MPTesterRS.Fields(1), 7, Len(MPTesterRS.Fields(1)))
            End If
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
            If Len(TesterRS.Fields(1)) = 15 Then
                Combo2.AddItem Mid(TesterRS.Fields(1), 8, Len(TesterRS.Fields(1)))
            Else
                Combo2.AddItem Mid(TesterRS.Fields(1), 7, Len(TesterRS.Fields(1)))
            End If
        End If
        TesterRS.MoveNext
    Loop
    
    
'    NameTmp = ""
'    TesterRS.MoveFirst
    WAIT_START_TIME_OUT = 0
    WAIT_TEST_CYCLE_OUT = 0
    POWER_ON_TIME = 0
    RT_INTERVAL = 0
    UNLOAD_DRIVER = 0
    Check3.value = 0
    CAPACTOR_CHARGE = 0
    Check6.value = 0
    NoCardTest.value = 0
    NO_CARD_TEST_TIME = 0
    Check5.value = 0
    
    If Combo2.ListCount = 1 And Combo2.List(0) = "" Then
        Call Combo2_Click
        Exit Sub
    End If
    
    Combo2.Enabled = True
        
    If GreaTekFlag = 1 Then
        If InStr(GreaTekChipName, "AU87100") Then
            Combo2.Text = Mid(GreaTekChipName, 8, Len(GreaTekChipName))
        Else
            Combo2.Text = Mid(GreaTekChipName, 7, Len(GreaTekChipName))
        End If
    End If
    
End Sub

Private Sub Combo2_Click()
 Const Ver = " V2.00"

Dim DB_Path As String
Dim MPTesterRS As New ADODB.Recordset
Dim TesterRS As New ADODB.Recordset
Dim ConnMPTesterDB As New ADODB.Connection
Dim ConnTesterDB As New ADODB.Connection
Dim MpTesterDB_Path As String
Dim TesterDB_Path As String
Dim NameTmp As String
Dim DbPath As String

'If Command1.Enabled = True Then
' MsgBox "STOP Testing"
'Exit Sub
'End If
 
NoCardTest.value = 0

If Trim(Combo2.Text = "") Then
    ChipName = Trim(Combo1.Text)
Else
    ChipName = Trim(Combo1.Text) & Trim(Combo2.Text)  '20060329 for GREATEK
End If


If GreaTekFlag = 1 Then
    ChipName = GreaTekChipName
    Combo1.Text = GreaTekChipName
End If
 
NO_CARD_TEST_TIME = 0

    '=============================================
        'connection to MPTester.mdb
    '=============================================
    MpTesterDB_Path = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\PGM_ListDB\MPTester_" & LastMPTesterDateCode & ".mdb"
    ConnMPTesterDB.Open MpTesterDB_Path
    MPTesterRS.CursorLocation = adUseClient
    MPTesterRS.Open "MPTester", ConnMPTesterDB, adOpenKeyset, adLockPessimistic

    NameTmp = ""
    MPTesterRS.MoveFirst
    
    '=============================================
        'connection to Tester.mdb
    '=============================================
    
    TesterDB_Path = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\PGM_ListDB\Tester_" & LastTesterDateCode & ".mdb"
    ConnTesterDB.Open TesterDB_Path
    TesterRS.CursorLocation = adUseClient
    TesterRS.Open "Tester", ConnTesterDB, adOpenKeyset, adLockPessimistic
    
    NameTmp = ""
    TesterRS.MoveFirst
    WAIT_START_TIME_OUT = 0
    WAIT_TEST_CYCLE_OUT = 0
    POWER_ON_TIME = 0
    RT_INTERVAL = 0
    UNLOAD_DRIVER = 0
    Check3.value = 0
    CAPACTOR_CHARGE = 0
    Check6.value = 0
    NoCardTest.value = 0
    NO_CARD_TEST_TIME = 0
    Check5.value = 0
    
    

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
    
        WAIT_START_TIME_OUT = MPTesterRS.Fields(3)
        WAIT_TEST_CYCLE_OUT = MPTesterRS.Fields(4)
        POWER_ON_TIME = MPTesterRS.Fields(5)
        RT_INTERVAL = MPTesterRS.Fields(6)
        UNLOAD_DRIVER = MPTesterRS.Fields(7)
        Check3.value = MPTesterRS.Fields(8)
        CAPACTOR_CHARGE = MPTesterRS.Fields(9)
        Check6.value = MPTesterRS.Fields(10)
        NoCardTest.value = MPTesterRS.Fields(11)
        NO_CARD_TEST_TIME = MPTesterRS.Fields(12)
        Check5.value = MPTesterRS.Fields(13)
        Check7.value = MPTesterRS.Fields(14)
        ReportCheck.value = MPTesterRS.Fields(15)
        LoopTestCycle = MPTesterRS.Fields(16)
        Need_GPIB = MPTesterRS.Fields(17)
    
    End If
    
    If DB_Path = "Tester" Then
    
        WAIT_START_TIME_OUT = TesterRS.Fields(3)
        WAIT_TEST_CYCLE_OUT = TesterRS.Fields(4)
        POWER_ON_TIME = TesterRS.Fields(5)
        RT_INTERVAL = TesterRS.Fields(6)
        UNLOAD_DRIVER = TesterRS.Fields(7)
        Check3.value = TesterRS.Fields(8)
        CAPACTOR_CHARGE = TesterRS.Fields(9)
        Check6.value = TesterRS.Fields(10)
        NoCardTest.value = TesterRS.Fields(11)
        NO_CARD_TEST_TIME = TesterRS.Fields(12)
        Check5.value = TesterRS.Fields(13)
        Check7.value = TesterRS.Fields(14)
        ReportCheck.value = TesterRS.Fields(15)
        LoopTestCycle = TesterRS.Fields(16)
        Need_GPIB = TesterRS.Fields(17)
    
    End If
    
        'Debug.Print WAIT_START_TIME_OUT
        'Debug.Print WAIT_TEST_CYCLE_OUT
        'Debug.Print POWER_ON_TIME
        'Debug.Print RT_INTERVAL
        'Debug.Print UNLOAD_DRIVER
        'Debug.Print Check3.value
        'Debug.Print CAPACTOR_CHARGE
        'Debug.Print Check6.value
        'Debug.Print NoCardTest.value
        'Debug.Print NO_CARD_TEST_TIME
        'Debug.Print Check5.value
    
    
'Select Case ChipName
'           MsgBox "*.HPC file error 2"
'           End
'End Select

If LoopTestCycle = 0 Then
    'LoopTestCheck.Enabled = False
    LoopTestCheck.value = 0
    LoopTestCheck.Caption = "LoopTest"
    Check3.value = 1
Else
    'LoopTestCheck.Enabled = True
    LoopTestCheck.value = 1
    LoopTestCheck.Caption = "LoopTest: " & LoopTestCycle
    
End If

If (ChipName = "AU9540CSF21") Or _
   (ChipName = "AU9540CSF20") Or _
   (ChipName = "AU9562CSF20") Or _
   (ChipName = "AU9525GLF20") Or _
   (Right(ChipName, 4) = "Port") Or _
   (ChipName = "AU6259ILS40") Or _
   (ChipName = "AU6259ILS41") Then
   
    VB6_Flag = False
Else
    VB6_Flag = True
End If

If SPILFlag Then
    If Dir(App.Path & "\" & ChipName) <> ChipName Then
        MsgBox ("Please Check Program Name !!")
        End
    End If
End If

If SPILFlag Then
    Label17.Caption = "Wait START TimeOut :" & WAIT_START_TIME_OUT & " s"
    Label18.Caption = "Test Timeout :" & WAIT_TEST_CYCLE_OUT & " s"
    Label19.Caption = "Power on Time :" & POWER_ON_TIME & " s"
    Label21.Caption = "Re-Test Interval :" & RT_INTERVAL & " s"
    Label24.Caption = "Unload Driver Time :" & UNLOAD_DRIVER & " s"
    Label25.Caption = "Capactor Charge Time :" & CAPACTOR_CHARGE & " s"
    'Label26.Caption = "實際總測試時間(含 load / unload) :"
    Label27.Caption = "Real Test Time :"
    Label35.Caption = "Avarage Test Time :"
    Label36.Caption = "UPH: "
    Label1.Caption = ChipName & " HOST Controller" & Ver
    Label30.Caption = "No Card Test Time :" & NO_CARD_TEST_TIME & " s"
Else
    Label17.Caption = "等 START 逾時 :" & WAIT_START_TIME_OUT & " s"
    Label18.Caption = "最大測試時間 :" & WAIT_TEST_CYCLE_OUT & " s"
    Label19.Caption = "Power on 時間 :" & POWER_ON_TIME & " s"
    Label21.Caption = "RT 間隔時間 :" & RT_INTERVAL & " s"
    Label24.Caption = "Unload Driver 時間 :" & UNLOAD_DRIVER & " s"
    Label25.Caption = "電容放電 時間 :" & CAPACTOR_CHARGE & " s"
    'Label26.Caption = "實際總測試時間(含 load / unload) :"
    Label27.Caption = "實際測試時間(不含 load / unload)  :"
    Label35.Caption = "Avarage Test Time :"
    Label36.Caption = "UPH: "
    Label1.Caption = ChipName & " HOST Controller" & Ver
    Label30.Caption = "不插卡測試時間 :" & NO_CARD_TEST_TIME & " s"
End If
Command1.Enabled = True
Command6.Enabled = True

End Sub

Private Sub Command1_Click()
On Error Resume Next
ReportBegin = 0
' report control  begin
Call ReportActive
If ReportCheck.value = 1 Then
Do
  DoEvents
Loop While ReportBegin = 0
End If

    avgTestTime = 0
    totalTestTime = 0
    testTime = 0
    Label35.Caption = "Avarage Test Time : "

ReportBegin = 0
' report control end

Dim HUBEnaOn As Byte
Dim TmpStr As String
Dim PowerStatus As Byte
Dim TesterStatus1
Dim TesterStatus2
Const AlarmLimit = 5   'Allen 20050607

Dim buf1
Dim buf2
Dim TestMode As Byte

Dim FirstRun As Byte
Dim WaitForReady ' timer

Dim TesterReady1 As Byte
Dim TesterReady2 As Byte

Dim ResetCounter1 As Byte
Dim ResetCounter2 As Byte

Dim TesterDownCount1 As Byte
Dim TesterDownCount2 As Byte

Dim WaitForPowerOn1   ' timer
Dim WaitForPowerOn2
 
Dim TesterDownCountTimer1  ' timer
Dim TesterDownCountTimer2


Dim WaitForStart
Dim WaitForVcc

Dim TotalRealTestTime
Dim OldTotalRealTestTime

Dim RealTestTime
Dim OldRealTestTime


Dim WaitStartTime

Dim GetStart As Integer
Dim TimeOut As Integer
 
Dim WaitStartCounter As Integer
Dim WaitStartTimeOutCounter As Integer
Dim WaitStartTimeOut As Integer


Dim TestCounter As Integer
Dim RTTestCounter1 As Integer
Dim RTTestCounter2 As Integer

'Test Result
Dim TestResult1 As String
Dim TestResult2 As String

Dim RTTestResult1
Dim RTTestResult2

Dim NoCardTestResult1 As String
Dim NoCardTestResult2 As String

' Stop Flag

Dim TestStop1 As Byte
Dim TestStop2 As Byte

Dim RTTestStop1 As Byte
Dim RTTestStop2 As Byte


Dim NoCardTestStop1 As Byte
Dim NoCardTestStop2 As Byte

' Wait for Test Time


Dim WaitForTest1
Dim WaitForTest2


Dim RTWaitForTest1
Dim RTWaitForTest2

Dim NoCardWaitForTest1
Dim NoCardWaitForTest2

Dim WaitTestTimeOutCounter1 As Integer
Dim WaitTestTimeOutCounter2 As Integer

Dim RTWaitTestTimeOutCounter1 As Integer
Dim RTWaitTestTimeOutCounter2 As Integer

Dim WaitTestTimeOut1 As Integer
Dim WaitTestTimeOut2 As Integer


Dim RTWaitTestTimeOut1 As Integer
Dim RTWaitTestTimeOut2 As Integer

' Test Cycle

Dim TestCycleTime1
Dim TestCycleTime2

Dim RTTestCycleTime1
Dim RTTestCycleTime2

Dim NoCardTestCycleTime1
Dim NoCardTestCycleTime2

Dim gpon1 As String
Dim gpon2 As String

Dim Bin1Counter1 As Long
Dim Bin2Counter1 As Long
Dim Bin3Counter1 As Long
Dim Bin4Counter1 As Long
Dim Bin5Counter1 As Long
Dim Bin1Counter2 As Long
Dim Bin2Counter2 As Long
Dim Bin3Counter2 As Long
Dim Bin4Counter2 As Long
Dim Bin5Counter2 As Long

Dim PassCounter1 As Long
Dim FailCounter1 As Long
Dim RTPassCounter1 As Long
Dim RTFailCounter1 As Long

Dim PassCounter2 As Long
Dim FailCounter2 As Long
Dim RTPassCounter2 As Long
Dim RTFailCounter2 As Long

Dim OffLPassCounter1 As Long
Dim OffLFailCounter1 As Long
Dim OffLRTPassCounter1 As Long
Dim OffLRTFailCounter1 As Long

Dim OffLPassCounter2 As Long
Dim OffLFailCounter2 As Long
Dim OffLRTPassCounter2 As Long
Dim OffLRTFailCounter2 As Long

Dim continuefail1 As Long
Dim continuefail1_bin2 As Long
Dim continuefail1_bin3 As Long
Dim continuefail1_bin4 As Long
Dim continuefail1_bin5 As Long
Dim continuefail2 As Long
Dim continuefail2_bin2 As Long
Dim continuefail2_bin3 As Long
Dim continuefail2_bin4 As Long
Dim continuefail2_bin5 As Long

Dim T1NotGetReadyCounter As Long
Dim T1UnknownCounter As Long
Dim T1GponFailCounter As Long
Dim T1SD_WFCounter As Long
Dim T1SD_RFCounter As Long
Dim T1CF_WFCounter As Long
Dim T1CF_RFCounter As Long
Dim T1XD_WFCounter As Long
Dim T1XD_RFCounter As Long
Dim T1SM_WFCounter As Long
Dim T1SM_RFCounter As Long
Dim T2NotGetReadyCounter As Long
Dim T2UnknownCounter As Long
Dim T2GponFailCounter As Long
Dim T2SD_WFCounter As Long
Dim T2SD_RFCounter As Long
Dim T2CF_WFCounter As Long
Dim T2CF_RFCounter As Long
Dim T2XD_WFCounter As Long
Dim T2XD_RFCounter As Long
Dim T2SM_WFCounter As Long
Dim T2SM_RFCounter As Long
Dim i As Long, k As Long
Dim result As Long
Dim DO_P As Long
Dim DI_P As Long
Dim NewPowerOnTime As Single

Dim debug1 As String
Dim GPIBInquiryTime

'atheist debug
GetGPIBStatus(0) = False
GetGPIBStatus(1) = False


'====================================='設定IO PORT輸入輸出

result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
result = DIO_PortConfig(card, Channel_P1CH, INPUT_PORT)
result = DIO_PortConfig(card, Channel_P1CL, INPUT_PORT)

result = DIO_PortConfig(card, Channel_P2A, OUTPUT_PORT)
result = DIO_PortConfig(card, Channel_P2B, OUTPUT_PORT)
result = DIO_PortConfig(card, Channel_P2CH, OUTPUT_PORT)
result = DIO_PortConfig(card, Channel_P2CL, OUTPUT_PORT)


'=====================================

continuefail1 = 0
continuefail1_bin2 = 0
continuefail1_bin3 = 0
continuefail1_bin4 = 0
continuefail1_bin5 = 0
continuefail2 = 0
continuefail2_bin2 = 0
continuefail2_bin3 = 0
continuefail2_bin4 = 0
continuefail2_bin5 = 0
SendMP_Flag = False
MPChipName = ""
RealChipName = ""
flag = 1


Command2.SetFocus

Cls
'=============================================' begin state
Print "begin state"


Call LockOption


'///////////////////////////////////////////////////////////////
'
'                     MAIN LOOP
'
'///////////////////////////////////////////////////////////////
Do
Cls
    If ChipName = "" Then
        Label12.BackColor = &H8000000F
        Label14.BackColor = &H8000000F
        Cls
        MsgBox "Select Chip"
        Exit Sub
    End If
    
    
    If Option1.value = True Then '雙機
        site1 = 1
        site2 = 1
        Label12.BackColor = &H8080FF
        Label14.BackColor = &H8080FF
    End If
    
    
    If Option2.value = True Then '1號機
        site1 = 1
        site2 = 0
        Label12.BackColor = &H8080FF
       Label14.BackColor = &H8000000F
    End If
    
    
    If Option3.value = True Then '2號機
        site1 = 0
        site2 = 1
        Label12.BackColor = &H8000000F
        Label14.BackColor = &H8080FF
    End If

   
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Display status
   
    TestResult1 = ""
    TestResult2 = ""
    
    RTTestResult1 = ""
    RTTestResult2 = ""
    
    
    NoCardTestResult1 = ""
    NoCardTestResult2 = ""
    
    'Text1.Text = ""
    'Text2.Text = ""
    Text3.Text = WaitStartCounter
    Text4.Text = WaitStartTime
    Text5.Text = WaitStartTimeOutCounter
    Text6.Text = TestCounter
    Text25.Text = RTTestCounter1
    Text26.Text = RTTestCounter2
    
    If Check1.value = 1 Then
        TestMode = 1  '離線模式
        Text3.Text = "TestMode"
        Text4.Text = "TestMode"
        Text5.Text = "TestMode"
        Text6.Text = "TestMode"
        'Text6.Text = OffLTestCounter
    
    Else
        TestMode = 0  '上線模式
        Text3.Text = WaitStartCounter
        Text4.Text = WaitStartTime
        Text5.Text = WaitStartTimeOutCounter
        Text6.Text = TestCounter
    
    End If

     
    Text7.Text = TestCycleTime1
    
    If TestMode = 0 Then
        Text8.Text = PassCounter1
        Text9.Text = FailCounter1
    Else
        Text8.Text = OffLPassCounter1
        Text9.Text = OffLFailCounter1
    End If
    
    Text10.Text = WaitTestTimeOutCounter1
    
    Text17.Text = RTTestCycleTime1
    
    If TestMode = 0 Then
        Text18.Text = RTPassCounter1
        Text19.Text = RTFailCounter1
    Else
        Text18.Text = OffLRTPassCounter1
        Text19.Text = OffLRTFailCounter1
    End If
    
    Text20.Text = RTWaitTestTimeOutCounter1
    
    Text11.Text = TestCycleTime2
    
    If TestMode = 0 Then
        Text12.Text = PassCounter2
        Text13.Text = FailCounter2
    Else
        Text12.Text = OffLPassCounter2
        Text13.Text = OffLFailCounter2
    End If
    
    Text14.Text = WaitTestTimeOutCounter2
   
    Text21.Text = RTTestCycleTime2
    
    If TestMode = 0 Then
        Text22.Text = RTPassCounter2
        Text23.Text = RTFailCounter2
    Else
        Text22.Text = OffLRTPassCounter2
        Text23.Text = OffLRTFailCounter2
    End If
    
    Text24.Text = RTWaitTestTimeOutCounter2
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Initial Counter and Control variable
    
    
    TestCycleTime1 = 0
    TestCycleTime2 = 0
    
    
    RTTestCycleTime1 = 0
    RTTestCycleTime2 = 0
    
    NoCardTestCycleTime1 = 0
    NoCardTestCycleTime2 = 0
    ' inital time out flag
    GetStart = 0
    WaitStartTime = 0
    WaitStartTimeOut = 0
    
    WaitTestTimeOut1 = 0
    TestStop1 = 0
    NoCardTestStop1 = 0
    WaitTestTimeOut2 = 0
    TestStop2 = 0
    NoCardTestStop2 = 0
    RTTestCycleTime1 = 0
    RTTestCycleTime2 = 0
    
    RTWaitTestTimeOut1 = 0
    RTTestStop1 = 0
    
    RTWaitTestTimeOut2 = 0
    RTTestStop2 = 0
    
    DoEvents
    
        
'*step1=>\\\\\\\\\\\\\\\\\\\\\ Get Start Signal From Handle
'*
'*wait Start Signal From Handle
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'  If ChipName = "AU6375AS" Then   'AU6375As has GPIB to do pwr control
  
'   Call GPIBSet
' End If

    WaitForStart = Timer   ' Get Vcc on from Chip
      
      
    If TestMode = 0 Then  'ON LINE MODE
    
          Print "wait Start"
            Do     'wait  (VCC PowerON) & (hander 5ms start) signal
                   DoEvents
                  WaitStartTime = Timer - WaitForStart
                   k = DO_ReadPort(card, Channel_P1CH, DI_P)
                   'Call MsecDelay(0.1)
        '   Loop Until DI_P = 14 Or DI_P = 13 Or DI_P = 12 Or WaitStartTime > WAIT_START_TIME_OUT  'Allen
           Loop Until DI_P = 14 Or DI_P = 13 Or DI_P = 12
           Label31.Caption = DI_P
            Print "DI_P=", DI_P
            TotalRealTestTime = Timer - OldTotalRealTestTime
            OldTotalRealTestTime = Timer
            OldRealTestTime = Timer
            
     Else
     
            Call MsecDelay(0.2)
            WaitStartTime = 0.2
            DI_P = 14
            TotalRealTestTime = Timer - OldTotalRealTestTime
            OldTotalRealTestTime = Timer
            OldRealTestTime = Timer
            
            'k = DO_WritePort(card, Channel_P1A, 255)
            'k = DO_WritePort(card, Channel_P1B, 255)
            
     End If
     
    buf1 = MSComm1.Input
    Label26.Caption = "實際總測試時間(含 load / unload) :" & TotalRealTestTime & "s"
    WaitStartCounter = WaitStartCounter + 1
              
    If WaitStartTime > WAIT_START_TIME_OUT Then
         WaitStartTimeOut = 1
         WaitStartTimeOutCounter = WaitStartTimeOutCounter + 1
    End If
      
       
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\  Debug

'========== continue 5 fail

     
     If Check7.value = 1 Then
    
     If continuefail1 >= AlarmLimit Or continuefail2 >= AlarmLimit Then
       
          continuefail1_bin2 = 0
          continuefail1_bin3 = 0
          continuefail1_bin4 = 0
          continuefail1_bin5 = 0
       
          continuefail2_bin2 = 0
          continuefail2_bin3 = 0
          continuefail2_bin4 = 0
          continuefail2_bin5 = 0
          
          continuefail1 = 0
          
          continuefail2 = 0
          
            k = DO_WritePort(card, Channel_P2CH, 0)
           Call MsecDelay(2)
           k = DO_WritePort(card, Channel_P2CH, 15)

       '   MsgBox "BIn2 fail"

    End If
    
    
    End If


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'
'    SHOW Alarm
'
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    If Check5.value = 1 Then 'arch change 940529
        If site1 = 1 And continuefail1 >= AlarmLimit Then
        
             If continuefail1_bin2 >= 3 Then
             Call MsecDelay(3)
             End If
             
             
             
        
            If continuefail1_bin2 >= AlarmLimit Then
                Alarm.Show
                Alarm.Label1 = "site1 countiue fail please check Chip Contact and Tester Driver!"
              '  MsgBox "site1 countiue fail please check Chip Contact and Tester Driver!"
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 1
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                    Do
                      DoEvents
                      If AlarmCtrl = 0 Then
                         Exit Do
                      End If
                    Loop While (1)
                
                Print "Alarm Clear"
               
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 1
                '
                '\\\\\\\\\3333\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                continuefail1_bin2 = 0
                continuefail1 = 0
                
            ElseIf continuefail1_bin3 >= AlarmLimit Then
                Alarm.Show
                Alarm.Label1 = "site1 countiue fail please check  Flash & CF & SD CARD!"
            
               ' MsgBox "site1 countiue fail please check  Flash & CF & SD CARD!"
                
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 2
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                Do
                  DoEvents
                  If AlarmCtrl = 0 Then
                     Exit Do
                  End If
                Loop While (1)
                Print "Alarm Clear"
               
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 2
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                continuefail1_bin3 = 0
                continuefail1 = 0
                
            ElseIf continuefail1_bin4 >= AlarmLimit Then
            
                Alarm.Show
                Alarm.Label1 = "site1 countiue fail please check XD CARD!"
                'MsgBox "site1 countiue fail please check XD CARD!"
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 3
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                Do
                  DoEvents
                  If AlarmCtrl = 0 Then
                     Exit Do
                  End If
                Loop While (1)
                Print "Alarm Clear"
               
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 3
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                continuefail1_bin4 = 0
                continuefail1 = 0
                
            ElseIf continuefail1_bin5 >= AlarmLimit Then
                Alarm.Show
                Alarm.Label1 = "site1 countiue fail please check MS CARD!"
               ' MsgBox "site1 countiue fail please check MS CARD!"
               
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 4
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                Do
                  DoEvents
                  If AlarmCtrl = 0 Then
                     Exit Do
                  End If
                Loop While (1)
                Print "Alarm Clear"
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 4
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                continuefail1_bin5 = 0
                continuefail1 = 0
                
            Else
            
                Print "Site1 check continuefail start !"
                
            End If
        End If
        
        If site2 = 1 And continuefail2 >= AlarmLimit Then
        
        
          If continuefail2_bin2 >= 3 Then
             Call MsecDelay(3)
          End If
             
            If continuefail2_bin2 >= AlarmLimit Then
            
                Alarm.Show
                Alarm.Label1 = "site2 countiue fail please check Chip Contact and Tester Driver!"
              '  MsgBox "site2 countiue fail please check Chip Contact and Tester Driver!"
              
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 5
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                Do
                
                  DoEvents
                  
                  If AlarmCtrl = 0 Then
                  
                     Exit Do
                  
                  End If
                  
                Loop While (1)
  
  
               Print "Alarm Clear"
               
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 5
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
              
              
              
              
              
                continuefail2_bin2 = 0
                continuefail2 = 0
            ElseIf continuefail2_bin3 >= AlarmLimit Then
              Alarm.Show
                Alarm.Label1 = "site2 countiue fail please check  Flash & CF & SD CARD!"
                'MsgBox "site2 countiue fail please check  Flash & CF & SD CARD!"
                
                
                   '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 6
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                Do
                
                  DoEvents
                  
                  If AlarmCtrl = 0 Then
                  
                     Exit Do
                  
                  End If
                  
                Loop While (1)
  
  
               Print "Alarm Clear"
               
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 6
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
              
                
                
                
                continuefail2_bin3 = 0
                continuefail2 = 0
            ElseIf continuefail2_bin4 >= AlarmLimit Then
                Alarm.Show
                Alarm.Label1 = "site2 countiue fail please check XD CARD!"
            
             '   MsgBox "site2 countiue fail please check XD CARD!"
             
                     '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 6
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                Do
                
                  DoEvents
                  
                  If AlarmCtrl = 0 Then
                  
                     Exit Do
                  
                  End If
                  
                Loop While (1)
  
  
               Print "Alarm Clear"
               
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 6
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
              
             
             
                continuefail2_bin4 = 0
                continuefail2 = 0
            ElseIf continuefail2_bin5 >= AlarmLimit Then
            
                Alarm.Show
                Alarm.Label1 = "site2 countiue fail please check MS CARD!"
                
                
                           '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 begin 6
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AlarmCtrl = 1
                 Cls
                Print "Alarm!!!"
                Do
                
                  DoEvents
                  
                  If AlarmCtrl = 0 Then
                  
                     Exit Do
                  
                  End If
                  
                Loop While (1)
  
  
               Print "Alarm Clear"
               
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                ' Allen 20050606 end 6
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
              
                
              '  MsgBox "site2 countiue fail please check MS CARD!"
                continuefail2_bin5 = 0
                continuefail2 = 0
            Else
                Print "Site2  check continuefail start !!"
            End If
        End If
        
    Else
        Print "on standard test step!"
    End If
   
 '*******************************************************
 '*
 '*   OPEN power
 '*
 '**********************************************************
   
   
     If DI_P < 12 And DI_P >= 15 Then   'Allen 20050607 , change DI_P > 15, to DI_P >= 15
        Print "no start"
        GoTo err
    Else
       Print "get start signal!"
       Call MsecDelay(CAPACTOR_CHARGE)
       Call MsecDelay(UNLOAD_DRIVER)
      k = DO_WritePort(card, Channel_P1A, &H7F) ' send 0111,1111 => Set power" Channel_P1A = 127
      k = DO_WritePort(card, Channel_P1B, &H7F) ' send 0111,1111 => Set power" Channel_P1b = 127
      
      NewPowerOnTime = POWER_ON_TIME - 0.4
      If NewPowerOnTime > 0 Then
        Call MsecDelay(NewPowerOnTime)
      End If
    End If
    
    
    
   
   
'*STEP2=> wait tester send ready signal\\\\\\\\\\\\\\
'*
'*  Check Tester Ready Signal
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


   LoopTest1_Flag = False
   LoopTest2_Flag = False
   LoopTestCounter1 = 0
   LoopTestCounter2 = 0
   
LoopTest_Label:
   
   If LoopTest1_Flag = True Or LoopTest2_Flag = True Then
        GetStart = 0
        TestStop1 = 0
        TestStop2 = 0
   End If
   

   MSComm2.InBufferCount = 0
   MSComm1.InBufferCount = 0
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
                               
                               If site1 = 1 Then
                             '  TesterReady1 = 1 '-----------Allen 20061207
                                    If TesterReady1 = 0 Then
                                            
                                        If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                                            MSComm1.Output = "~"
                                                    
                                            If HUBEnaOn = 0 Then
                                                k = DO_WritePort(card, Channel_P1CH, &H0) ' set HUB module Ena on
                                                HUBEnaOn = 1
                                            End If
                                                
                                        End If
                                                
                                        buf1 = MSComm1.Input
                                        TesterStatus1 = TesterStatus1 & buf1
                                        If (InStr(1, TesterStatus1, "Ready") <> 0) Or (InStr(1, TesterStatus1, "70") <> 0) Then
                                            TesterReady1 = 1
                                        End If
                                            
                                        'atheist 2011/4/27
                                        If (GetGPIBStatus(0) = False) And (VB6_Flag = True) Then
                                            MSComm1.Output = "AUGPIBGPIBGPIB"   ' trans strat test signal to TEST PC
                                            MSComm1.InBufferCount = 0
                                            MSComm1.InputLen = 0
                                            GPIBInquiryTime = Timer
                                            'Call MsecDelay(0.02)
                                            
                                            Do
                                                buf1 = MSComm1.Input
                                                TesterStatus1 = TesterStatus1 & buf1
                                                Call MsecDelay(0.05)
                                            Loop Until (InStr(1, TesterStatus1, "GPIBReady") <> 0) _
                                                        Or (InStr(1, TesterStatus1, "GPIBUNReady") <> 0) _
                                                        Or (Timer - GPIBInquiryTime > 1)
                                            
                                            If (Timer - GPIBInquiryTime > 1) Then
                                                GoTo LoopTest_Label
                                            End If
                                            
                                            MSComm1.Output = "AUGPIBACK"
                                            GetGPIBStatus(0) = True
                                            
                                            If InStr(1, TesterStatus1, "GPIBReady") <> 0 Then
                                                GPIBReady(0) = True
                                                Site1GPIB_Label.BackColor = &H80FFFF
                                                Site1GPIB_Label.ForeColor = &HFF0000
                                                Site1GPIB_Label.Caption = " GPIB Rdy"
                                            Else
                                                GPIBReady(0) = False
                                                Site1GPIB_Label.BackColor = &H80FFFF
                                                Site1GPIB_Label.ForeColor = &H80FF&
                                                Site1GPIB_Label.Caption = " No GPIB"
                                            End If
                                            Call MsecDelay(0.05)
                                            GoTo LoopTest_Label
                                        
                                        End If
                                    
                                    End If
                                    
                                    Else
                                        TesterReady1 = 1
                                     
                               End If
                               '========================
                               ' wait for tester2 ready
                               '========================
                                
                               If site2 = 1 Then
                                    If TesterReady2 = 0 Then
                                            
                                        If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                                            MSComm2.Output = "~"
                                                    
                                            If HUBEnaOn = 0 Then
                                                k = DO_WritePort(card, Channel_P1CH, &H0) ' set HUB module Ena on
                                                HUBEnaOn = 1
                                            End If
                                                  
                                        End If
                                                
                                        buf2 = MSComm2.Input
                                        TesterStatus2 = TesterStatus2 & buf2
                                                  
                                        If (InStr(1, TesterStatus2, "Ready") <> 0) Or (InStr(1, TesterStatus2, "0x70") <> 0) Then
                                            TesterReady2 = 1
                                        End If
                                    
                                        If (GetGPIBStatus(1) = False) And (VB6_Flag = True) Then
                                            MSComm2.Output = "AUGPIBGPIBGPIB"   ' trans strat test signal to TEST PC
                                            MSComm2.InBufferCount = 0
                                            MSComm2.InputLen = 0
                                            GPIBInquiryTime = Timer
                                            'Call MsecDelay(0.02)
                                            
                                            Do
                                                buf2 = MSComm2.Input
                                                TesterStatus2 = TesterStatus2 & buf2
                                                Call MsecDelay(0.1)
                                            Loop Until (InStr(1, TesterStatus2, "GPIBReady") <> 0) _
                                                        Or (InStr(1, TesterStatus2, "GPIBUNReady") <> 0) _
                                                        Or (Timer - GPIBInquiryTime > 1)
                                                        
                                            If (Timer - GPIBInquiryTime > 1) Then
                                                GoTo LoopTest_Label
                                            End If
                                            
                                            MSComm2.Output = "AUGPIBACK"
                                            GetGPIBStatus(1) = True
                                            
                                            If InStr(1, TesterStatus2, "GPIBReady") <> 0 Then
                                                GPIBReady(1) = True
                                                Site2GPIB_Label.BackColor = &H80FFFF
                                                Site2GPIB_Label.ForeColor = &HFF0000
                                                Site2GPIB_Label.Caption = " GPIB Rdy"
                                            Else
                                                GPIBReady(1) = False
                                                Site2GPIB_Label.BackColor = &H80FFFF
                                                Site2GPIB_Label.ForeColor = &H80FF&
                                                Site2GPIB_Label.Caption = " No GPIB"
                                            End If
                                            Call MsecDelay(0.05)
                                            GoTo LoopTest_Label
                                        
                                        End If
                                    
                                    End If
                                            
                               Else
                                  TesterReady2 = 1
                                     
                               End If
                                
                                
                              '===================================
                              ' Reset rountine : condsider Reset fail
                              '===================================
                           
                              If Timer - WaitForReady > 2 Then
                         
                                If ResetCounter1 > 2 Or ResetCounter2 > 2 Then  ' Alarm for reset fail
                                  Call PrintReport  'print routine
                                  Print "Alarm : Reset PC fail"
                                  MsgBox "Alarm : Reset PC fail "
                                  ResetCounter1 = 0
                                  ResetCounter2 = 0
                                 Exit Sub
                                Else
                                
                                   '=========== Reset  Rountine
                                    
                                    If TesterReady1 = 0 And TesterDownCount1 = 0 And FirstRun = 1 Then ' reset tester1
                                          '============== close module power
                                           ResetCounter1 = ResetCounter1 + 1
                                           k = DO_WritePort(card, Channel_P1A, &HFF) ' send 0111,1111 => Set power" Channel_P1A = 127
                                           k = DO_WritePort(card, Channel_P1B, &HFF) ' send 0111,1111 => Set power" Channel_P1A = 127
                                          '============= Reset PC
                                          
                                          TesterDownCount1 = 1
                                          k = DO_WritePort(card, Channel_P2CH, 0)
                                          Call MsecDelay(2)
                                          k = DO_WritePort(card, Channel_P2CH, 15)
                                          WaitForPowerOn1 = Timer
                                                  
                                          '============== clear comm buffer
                                          MSComm1.InBufferCount = 0
                                          TesterStatus1 = ""
               
                                    End If
                                
                                
                                   If TesterReady2 = 0 And TesterDownCount2 = 0 And FirstRun = 1 Then ' reset tester2
                                         '============== close module power
                                          ResetCounter2 = ResetCounter2 + 1
                                          k = DO_WritePort(card, Channel_P1A, &HFF) ' send 0111,1111 => Set power" Channel_P1A = 127
                                           k = DO_WritePort(card, Channel_P1B, &HFF) ' send 0111,1111 => Set power" Channel_P1A = 127
                                         '============== Reset PC
                                            TesterDownCount2 = 1
                                             k = DO_WritePort(card, Channel_P2CH, 0)
                                            Call MsecDelay(2)
                                            k = DO_WritePort(card, Channel_P2CH, 15)
                                           WaitForPowerOn2 = Timer
                                         '============== clear comm buffer
                                           MSComm2.InBufferCount = 0
                                            TesterStatus2 = ""
                                                  
                                   End If
               
                               End If
                                        
                               '===============================
                               ' screen down count routine
                               '==============================
                               
                               If TesterDownCount1 = 1 Then
                               
                                     TesterDownCountTimer1 = Timer - WaitForPowerOn1
                                     Label28.Caption = CInt(TesterDownCountTimer1)
                                     
                                    If TesterReady1 = 1 Then
                                        '====== open module power
                                         k = DO_WritePort(card, Channel_P1A, &H7F) ' send 0111,1111 => Set power" Channel_P1b = 127
                                          k = DO_WritePort(card, Channel_P1B, &H7F)
                                         Call MsecDelay(POWER_ON_TIME)
                                        
                                        '=== clear flag
                                        TesterDownCount1 = 0
                                     End If
                                     
                                     If TesterDownCountTimer1 > 90 Then  'Reset fail
                                         TesterDownCount1 = 0
                                     End If
                                     
                                  
                               End If
                               
                               
                                If TesterDownCount2 = 1 Then
                                
                                     TesterDownCountTimer2 = Timer - WaitForPowerOn2
                                     Label29.Caption = CInt(TesterDownCountTimer2)
                                     
                                      If TesterReady2 = 1 Then
                                         '====== open module power
                                         k = DO_WritePort(card, Channel_P1B, &H7F) ' send 0111,1111 => Set power" Channel_P1b = 127
                                          k = DO_WritePort(card, Channel_P1A, &H7F)
                                         Call MsecDelay(POWER_ON_TIME)
                                        '=== clear flag
                                        TesterDownCount2 = 0
                                      End If
                                     
                                     
                                        If TesterDownCountTimer2 > 90 Then    ' Reset fail
                                         TesterDownCount2 = 0
                                     End If
                               End If
                                 
                           End If
                           
                           
                            If (Need_GPIB = 1) And (VB6_Flag = True) Then
                                If (TesterReady1 = 1) And (TesterReady2 = 1) Then
                                    If (site1 = 1) And (site2 = 1) Then 'dual site
                                        If (GPIBReady(0) = False) And (GPIBReady(1) = False) Then
                                            MsgBox ("Check GPIB CARD Connection Error !!")
                                            End
                                        End If
                                    ElseIf (site1 = 1) And (site2 = 0) Then
                                        If (GPIBReady(0) = False) Then
                                            MsgBox ("Check Site1 GPIB CARD Connection Error !!")
                                            End
                                        End If
                                    ElseIf (site1 = 0) And (site2 = 1) Then
                                        If (GPIBReady(1) = False) Then
                                            MsgBox ("Check Site2 GPIB CARD Connection Error !!")
                                            End
                                        End If
                                    End If
                                End If
                            End If
                      
                      Loop Until (TesterReady1 = 1) And (TesterReady2 = 1)
                      FirstRun = 1
         
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'*
'*    Testing Loop
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

        
        
If (DI_P >= 12) And (DI_P < 15) Then
  
    
        ' init falg
         GetStart = 1
    Label3.BackColor = RGB(255, 255, 255)
    Label3 = ""
    Label16.BackColor = RGB(255, 255, 255)
    Label16 = ""
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '
    '                Site1 and Site2  begin
    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
             
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '        Testing LED function    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Print "==========================="
          
  
                ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  Allen 0526 begin 1 : for no card test,pull high Card detect
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                   If NoCardTest.value = 1 Then
                        
                   k = DO_WritePort(card, Channel_P2CL, &HF)   'pull High
                   Else
                   
                   k = DO_WritePort(card, Channel_P2CL, &H0)
                   End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  Allen 0526 End  1 : for no card test,pull high Card detect
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
   
    '*STEP4=> Waitting for Response from  Tester\\\\\\\\\\\\\\\\\\\\\
    '*
    '*    Wait Test Result from each Tester
    '*
    '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
   '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      '
      '  Allen 0601 Remark : no card on board test card detect and card change signal
      '
      '
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
       NoCardTestResult1 = ""
       NoCardTestResult2 = ""
        
        
        If site1 = 1 And NoCardTest.value = 1 Then
    
        MSComm1.Output = ChipName   ' trans strat test signal to TEST PC
        MSComm1.InBufferCount = 0
        MSComm1.InputLen = 0
        NoCardWaitForTest1 = Timer ' wait for timer  and test result
        End If
    
        
        If site2 = 1 And NoCardTest.value = 1 Then
        
        MSComm2.Output = ChipName   ' trans strat test signal to TEST PC
        MSComm2.InBufferCount = 0
        MSComm2.InputLen = 0
        NoCardWaitForTest2 = Timer
        End If
    
    
         Print "send begin test signal to test"
         TesterStatus1 = ""
         TesterStatus2 = ""
      
      
             Do
                DoEvents
                
                If site1 = 1 And NoCardTest.value = 1 Then
                    If NoCardTestStop1 = 0 Then
                    
                        If MSComm1.InBufferCount >= 4 Then
                            NoCardTestResult1 = MSComm1.Input
                          
                            
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
                '========================
                
                If site2 = 1 And NoCardTest.value = 1 Then
                    If NoCardTestStop2 = 0 Then
                         If MSComm2.InBufferCount >= 4 Then
                                NoCardTestResult2 = MSComm2.Input
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
            
              
            Loop Until (NoCardTestStop1 = 1) And (NoCardTestStop2 = 1)
            
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      '
      '  Allen 0526 Remark : no card on board test card detect and card change signal
      '
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
         '*STEP3=>Send command to PC teser\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '*
    '*    Send ChipName to PC teser
    '*
    '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'If ((Option1 = True) And (LoopTest1_Flag = False) And (LoopTest2_Flag = False)) Then
        TestResult1 = ""
        TestResult2 = ""
    'ElseIf ((Option2 = True) And (LoopTest1_Flag = True)) Then
    '    TestResult1 = ""
    'ElseIf ((Option3 = True) And (LoopTest2_Flag = True)) Then
    '    TestResult2 = ""
    'End If
    
    
    
    If NoCardTest.value = 1 Then
    
             k = DO_WritePort(card, Channel_P2CL, 0)   'pull down
             
'             Call MsecDelay(0.1)
            If site1 = 1 Then  '****** Continue condition lock at PC tester
                
                MSComm1.Output = NoCardTestResult1   ' only pass can continue at PC Tester
                MSComm1.InBufferCount = 0
                MSComm1.InputLen = 0
                WaitForTest1 = Timer ' wait for timer  and test result
             
                If NoCardTestResult1 <> "PASS" Then
                    TestResult1 = NoCardTestResult1
                End If
            End If
          
            
            If site2 = 1 Then
                MSComm2.Output = NoCardTestResult2   ' only pass can continue at PC Tester
                MSComm2.InBufferCount = 0
                MSComm2.InputLen = 0
                WaitForTest2 = Timer
                If NoCardTestResult2 <> "PASS" Then
                    TestResult2 = NoCardTestResult2
                End If
            End If
    
    
    
    Else
    
            If site1 = 1 Then
                NoCardTestResult1 = "PASS"
                If ChipName = "AU6254NJ" Then
              
                   TmpStr = Chr(128)
                   MSComm1.Output = TmpStr
                ElseIf (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                    For i = 1 To Len(ChipName)
                        MSComm1.Output = Mid(ChipName, i, 1)
                        Call MsecDelay(0.02)
                        'Debug.Print Mid(ChipName, i, 1)
                    Next
                Else
                   MSComm1.Output = ChipName   ' trans strat test signal to TEST PC
                End If
                
                MSComm1.InBufferCount = 0
                MSComm1.InputLen = 0
                WaitForTest1 = Timer ' wait for timer  and test result
            End If
            
            
            If site2 = 1 Then
                NoCardTestResult2 = "PASS"
                If ChipName = "AU6254NJ" Then
               
                   TmpStr = Chr(128)
                   MSComm2.Output = TmpStr
                ElseIf (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                    For i = 1 To Len(ChipName)
                        MSComm2.Output = Mid(ChipName, i, 1)
                        Call MsecDelay(0.02)
                        'Debug.Print Mid(ChipName, i, 1)
                    Next
                Else
                
                   MSComm2.Output = ChipName   ' trans strat test signal to TEST PC
                End If
                MSComm2.InBufferCount = 0
                MSComm2.InputLen = 0
                WaitForTest2 = Timer
            End If
    End If
    
   
                    Do
                               DoEvents
                               
                               If site1 = 1 And NoCardTestResult1 = "PASS" Then
                                            
                                            If TestStop1 = 0 Then
                                            
                                                  If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                                                    MSComm1.Output = "~"
                                                  End If
                                                
                                                  If MSComm1.InBufferCount >= 3 Then
                                                    TestResult1 = TestResult1 & MSComm1.Input
                                                  
                                                  
                                                  End If
                                                  
                                                  
                                                     
                                                  TestResult1 = Parser(TestResult1)
                                                  If ChipName = "AU6254NJ" Then
                                                     Label33.Caption = AU6254Msg
                                                  End If
                                                  
                                                  TestCycleTime1 = Timer - WaitForTest1
                                                  
                                                  If (TestResult1 <> "" Or TestCycleTime1 > WAIT_TEST_CYCLE_OUT) Then
                                                  TestStop1 = 1
                                                  End If
                                            End If
                                            
                               Else
                                  TestStop1 = 1
                                     
                               End If
                               '========================
                               
                               If site2 = 1 And NoCardTestResult2 = "PASS" Then
                                        If TestStop2 = 0 Then
                                        
                                             If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                                                MSComm2.Output = "~"
                                             End If
                                                
                                             If MSComm2.InBufferCount >= 4 Then
                                              TestResult2 = MSComm2.Input
                                             End If
                                              
                                              TestResult2 = Parser(TestResult2)
                                              
                                                If ChipName = "AU6254NJ" Then
                                                     Label34.Caption = AU6254Msg
                                                  End If
                                                  
                                              
                                              TestCycleTime2 = Timer - WaitForTest2
                                             
                                               If (TestResult2 <> "" Or TestCycleTime2 > WAIT_TEST_CYCLE_OUT) Then
                                              TestStop2 = 1
                                              End If
                                        End If
                                        
                                Else
                                  TestStop2 = 1
                                     
                                     
                               End If
                             
                               
                      Loop Until (TestStop1 = 1) And (TestStop2 = 1)
       
        
        
    '\\\\\\\\\\\\\\\\\\\\\\\\\\}wait Tester response END
                           
         
                             
                             
         '   MsgBox TestResult1
            If TestMode = 0 Then
                TestCounter = TestCounter + 1 ' Allen Debug
            Else
                OffLTestCounter = OffLTestCounter + 1
            End If
            
            If TestCycleTime1 > WAIT_TEST_CYCLE_OUT And site1 = 1 Then
            
              TestResult1 = "TimeOut"
              WaitTestTimeOut1 = 1
              WaitTestTimeOutCounter1 = WaitTestTimeOutCounter1 + 1
            End If
                     
            If TestCycleTime2 > WAIT_TEST_CYCLE_OUT And site2 = 1 Then
            
               TestResult2 = "TimeOut"
              WaitTestTimeOut2 = 1
              WaitTestTimeOutCounter2 = WaitTestTimeOutCounter2 + 1
            End If
         
              If site1 = 1 Then
                Print "TestResult1= "; TestResult1
               
            End If
            
            If site2 = 1 Then
                Print "TestResult2= "; TestResult2
                               
            End If
                 
            If HUBEnaOn = 1 Then
                k = DO_WritePort(card, Channel_P1CH, &HF) ' set HUB module Ena off
                HUBEnaOn = 0
            End If
            
       '/////////////////////////////////////////////////////////////////////////
       '
       '   RT Condition
       '
       '//////////////////////////////////////////////////////////////////////////
                 
                 
        If Check3.value = 1 Then   ' 不RT=> not low yield sorting
            GoTo err
        End If
        
        
        If site1 = 1 And site2 = 0 Then
            If TestResult1 = "PASS" Then
                GoTo err
            End If
            
        End If
        
        
        If site1 = 0 And site2 = 1 Then
            If TestResult2 = "PASS" Then
                GoTo err
            End If
            
        End If
        
        If site1 = 1 And site2 = 1 Then
            If TestResult1 = "PASS" And TestResult2 = "PASS" Then
                GoTo err
            End If
            
        End If
        '////////////////////////////// initial condition
           
         Print "RT begin"
                 
         '1.close power
         '2.delay 10 s
         '3.send power
         '4.RT core
         
         Print "close power"
          
            k = DO_WritePort(card, Channel_P1A, &HFF)
            k = DO_WritePort(card, Channel_P1B, &HFF)
        
         
         Call MsecDelay(RT_INTERVAL)  ' to let system unload driver
         
         Print "Send power"
         k = DO_WritePort(card, Channel_P1A, &H7F)
         k = DO_WritePort(card, Channel_P1B, &H7F)
                  
         Call MsecDelay(POWER_ON_TIME)
         
                 
                 
                 
            If site1 = 1 And TestResult1 <> "PASS" Then
                
                MSComm1.Output = ChipName ' trans strat test signal to TEST PC
                MSComm1.InBufferCount = 0
                MSComm1.InputLen = 0
                RTWaitForTest1 = Timer ' wait for timer  and test result
             End If
                
             If site2 = 1 And TestResult2 <> "PASS" Then
                 
                MSComm2.Output = ChipName    ' trans strat test signal to TEST PC
                MSComm2.InBufferCount = 0
                MSComm2.InputLen = 0
                RTWaitForTest2 = Timer
             End If
             
            
             Print "send begin test signal to test"
             
             '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Wait for Response from PC Tester
                
            Do
                DoEvents
                
                If site1 = 1 And TestResult1 <> "PASS" Then
                    If RTTestStop1 = 0 Then
                        RTTestResult1 = MSComm1.Input
                        
                        RTTestCycleTime1 = Timer - RTWaitForTest1
                        
                        If (RTTestResult1 <> "" Or RTTestCycleTime1 > WAIT_TEST_CYCLE_OUT) Then
                        
                            RTTestCounter1 = RTTestCounter1 + 1 ' Allen Debug
                            RTTestStop1 = 1
                        End If
                    End If
                
                Else
                    RTTestStop1 = 1
                
                End If
                '========================
                
                If site2 = 1 And TestResult2 <> "PASS" Then
                    If RTTestStop2 = 0 Then
                        RTTestResult2 = MSComm2.Input
                        
                        RTTestCycleTime2 = Timer - RTWaitForTest2
                        
                        If (RTTestResult2 <> "" Or RTTestCycleTime2 > WAIT_TEST_CYCLE_OUT) Then
                            RTTestStop2 = 1
                            RTTestCounter2 = RTTestCounter2 + 1
                        End If
                    End If
                
                Else
                    RTTestStop2 = 1
                
                
                End If
            
            
            Loop Until (RTTestStop1 = 1) And (RTTestStop2 = 1)
                           
                If site1 = 1 Then
                    Print "RTTestResult1= "; RTTestResult1
                End If
                
                If site2 = 1 Then
                    Print "RTTestResult2= "; RTTestResult2
                End If
                
               
                
                If RTTestCycleTime1 > WAIT_TEST_CYCLE_OUT And site1 = 1 Then
                
                    RTWaitTestTimeOut1 = 1
                    RTWaitTestTimeOutCounter1 = RTWaitTestTimeOutCounter1 + 1
                End If
                
                If RTTestCycleTime2 > WAIT_TEST_CYCLE_OUT And site2 = 1 Then
                
                    RTWaitTestTimeOut2 = 1
                    RTWaitTestTimeOutCounter2 = RTWaitTestTimeOutCounter2 + 1
                End If
                 
                 
                 
                     
                 
End If  '////////////////////// Test end
              '  Testing Loop end
                 
   
   
                
'////////////////////////////////////// Tester 1 PASS Bin

err:
    ' default value
       gpon1 = "PASS"
       gpon2 = "PASS"
      
      
    If Check6.value = 1 Then     'check GPON7_LED & Power_LED
    
       gpon1 = ""
       gpon2 = ""
    
        Dim DI_Power As Long
        k = DO_ReadPort(card, Channel_P1CL, DI_Power)
        Print "DI_Power="; DI_Power
        
        If TestResult1 = "PASS" Then
            If site1 = 1 And (DI_Power Mod 4 = 0) Then
                gpon1 = "PASS"
                Print "gpon1="; gpon1
            Else
                gpon1 = "FAIL"
                Print "gpon1="; gpon1
                 TestResult1 = "gponFail"
            End If
        End If
        
        If TestResult2 = "PASS" Then
            If site2 = 1 And (DI_Power <= 3) Then
                gpon2 = "PASS"
                Print "gpon2="; gpon2
            Else
                gpon2 = "FAIL"
                Print "gpon2="; gpon2
                 TestResult2 = "gponFail"
            End If
        End If
        
        
     
    
        
    End If
    

    

Label3 = TestResult1
Label16 = TestResult2

Print "close power"

If Check2.value = 0 Then  'continuous supply power
    k = DO_WritePort(card, Channel_P1A, &HFF)
    k = DO_WritePort(card, Channel_P1B, &HFF)
End If

If site1 = 1 Then

 Select Case TestResult1
        Case "PASS"
            TestResult1 = "PASS"
        Case "UNKNOW", "bin2", "Bin2"
            TestResult1 = "bin2"
            T1UnknownCounter = T1UnknownCounter + 1
            BinForm.Text1.Text = T1UnknownCounter
        Case "gponFail", "bin3", "Bin3"
            TestResult1 = "bin3"
            T1GponFailCounter = T1GponFailCounter + 1
            BinForm.Text29.Text = T1GponFailCounter
        Case "SD_WF"
            TestResult1 = "bin3"
            T1SD_WFCounter = T1SD_WFCounter + 1
            BinForm.Text2.Text = T1SD_WFCounter
        Case "SD_RF"
            TestResult1 = "bin3"
            T1SD_RFCounter = T1SD_RFCounter + 1
            BinForm.Text3.Text = T1SD_RFCounter
        Case "CF_WF"
            TestResult1 = "bin3"
            T1CF_WFCounter = T1CF_WFCounter + 1
            BinForm.Text4.Text = T1CF_WFCounter
        Case "CF_RF"
            TestResult1 = "bin3"
            T1CF_RFCounter = T1CF_RFCounter + 1
            BinForm.Text5.Text = T1CF_RFCounter
        Case "XD_WF", "bin4", "Bin4"
            TestResult1 = "bin4"
            T1XD_WFCounter = T1XD_WFCounter + 1
            BinForm.Text6.Text = T1XD_WFCounter
        Case "XD_RF"
            TestResult1 = "bin4"
            T1XD_RFCounter = T1XD_RFCounter + 1
            BinForm.Text7.Text = T1XD_RFCounter
        Case "MS_WF", "bin5", "TimeOut", "Bin5"
            TestResult1 = "bin5"
            T1SM_WFCounter = T1SM_WFCounter + 1
             BinForm.Text8.Text = T1SM_WFCounter
        Case "MS_RF"
            TestResult1 = "bin5"
            T1SM_RFCounter = T1SM_RFCounter + 1
            BinForm.Text9.Text = T1SM_RFCounter
        Case Else
            TestResult1 = "bin2"
            T1UnknownCounter = T1UnknownCounter + 1
            BinForm.Text1.Text = T1UnknownCounter
        End Select

End If
If site2 = 1 Then

Select Case TestResult2
        Case "PASS"
            TestResult2 = "PASS"
        Case "UNKNOW", "bin2", "Bin2"
            TestResult2 = "bin2"
            T2UnknownCounter = T2UnknownCounter + 1
            BinForm.Text15.Text = T2UnknownCounter
        Case "gponFail", "bin3", "Bin3"
            TestResult2 = "bin3"
            T2GponFailCounter = T2GponFailCounter + 1
            BinForm.Text30.Text = T2GponFailCounter
        Case "SD_WF"
            TestResult2 = "bin3"
            T2SD_WFCounter = T2SD_WFCounter + 1
            BinForm.Text16.Text = T2SD_WFCounter
        Case "SD_RF"
            TestResult2 = "bin3"
            T2SD_RFCounter = T2SD_RFCounter + 1
            BinForm.Text17.Text = T2SD_RFCounter
        Case "CF_WF"
            TestResult2 = "bin3"
            T2CF_WFCounter = T2CF_WFCounter + 1
            BinForm.Text18.Text = T2CF_WFCounter
        Case "CF_RF"
            TestResult2 = "bin3"
            T2CF_RFCounter = T2CF_RFCounter + 1
            BinForm.Text19.Text = T1CF_RFCounter
        Case "XD_WF", "bin4", "Bin4"
            TestResult2 = "bin4"
            T2XD_WFCounter = T2XD_WFCounter + 1
            BinForm.Text20.Text = T2XD_WFCounter
        Case "XD_RF"
            TestResult2 = "bin4"
            T2XD_RFCounter = T2XD_RFCounter + 1
            BinForm.Text21.Text = T2XD_RFCounter
        Case "MS_WF", "bin5", "TimeOut", "Bin5"
            TestResult2 = "bin5"
            T2SM_WFCounter = T2SM_WFCounter + 1
             BinForm.Text22.Text = T2SM_WFCounter
        Case "MS_RF"
            TestResult2 = "bin5"
            T2SM_RFCounter = T2SM_RFCounter + 1
            BinForm.Text23.Text = T2SM_RFCounter
        Case Else
            TestResult2 = "bin2"
            T2UnknownCounter = T2UnknownCounter + 1
            BinForm.Text15.Text = T2UnknownCounter
        End Select

End If


If GreaTekChipName = "AU6368A1" Then
    
     If TestResult1 = "bin2" Then
        TestResult1 = "bin4"
     End If
        
     If TestResult1 = "bin3" Then
        TestResult1 = "bin5"
     End If


    If TestResult2 = "bin2" Then
        TestResult2 = "bin4"
     End If
        
     If TestResult2 = "bin3" Then
        TestResult2 = "bin5"
     End If

End If

If testTime = 0 Or SendMP_Flag = True Then
    
    If testTime = 0 Then
        testTime = testTime + 1
    End If
    
    RealTestTime = Timer - OldRealTestTime
    If SPILFlag Then
        Label27.Caption = "Real Test Time :" & RealTestTime & " s"
    Else
        Label27.Caption = "實際測試時間(不含 load / unload)  :" & RealTestTime & " s"
    End If
Else
    RealTestTime = Timer - OldRealTestTime
    totalTestTime = totalTestTime + RealTestTime
    avgTestTime = totalTestTime / testTime
    Debug.Print avgTestTime
    Label35.Caption = "Avarage Test Time : " & avgTestTime & " s, 測試數： " & testTime
    testTime = testTime + 1
 
    If SPILFlag Then
        Label27.Caption = "Real Test Time :" & RealTestTime & " s"
    Else
        Label27.Caption = "實際測試時間(不含 load / unload)  :" & RealTestTime & " s"
    End If
End If
 
 Label36.Caption = "UPH: " & (3600 / avgTestTime)
 
If (TestResult1 = "PASS" Or RTTestResult1 = "PASS") And GetStart = 1 And site1 = 1 And gpon1 = "PASS" Then

        Print "\\\\\\\\\\site1 = "; TestResult1
              
        If TestResult1 = "PASS" Then
            
        'atheist debug
            If LoopTestCycle <> 0 Then
                'HostForm.Cls
                LoopTestCounter1 = LoopTestCounter1 + 1
                HostForm.Print "Site1 LoopTest: " & LoopTestCounter1
                LoopTest1_Flag = True       'need Loop Test
                
                If LoopTestCounter1 = LoopTestCycle Then
                    LoopTestCounter1 = 0
                    LoopTest1_Flag = False
                    
                    If TestMode = 0 Then
                        PassCounter1 = PassCounter1 + 1
                    Else
                        OffLPassCounter1 = OffLPassCounter1 + 1
                    End If
                
                End If
            Else
                LoopTest1_Flag = False
                If TestMode = 0 Then
                    PassCounter1 = PassCounter1 + 1
                Else
                    OffLPassCounter1 = OffLPassCounter1 + 1
                End If
            End If
            
            Label2 = "PASS1"
            Label3.BackColor = RGB(0, 255, 0)
        Else
            If TestMode = 0 Then
                RTPassCounter1 = RTPassCounter1 + 1
            Else
                OffLRTPassCounter1 = OffLRTPassCounter1 + 1
            End If
            Print "-------------- RTPASS 1"
            Label2 = "RTPASS1"
            Label3.BackColor = RGB(0, 0, 255)
        End If
        
        If LoopTest1_Flag = False Then
            
            If TestMode = 0 Then
                Bin1Counter1 = Bin1Counter1 + 1
            End If
            
            BinForm.Text10.Text = Bin1Counter1
        
        
            Call PCI7248_bin(Channel_P1B, PCI7248_PASS, Check2.value)
        
        End If
       
End If
            
'////////////////////////////////////// Tester 1 FAIL Bin
If Check3.value = 0 Then
    If (((TestResult1 <> "PASS" And RTTestResult1 <> "PASS") And GetStart = 1) Or WaitStartTimeOut = 1) And site1 = 1 Then
    
    
         Print "\\\\\\\\\\site1 = "; TestResult1
        
        If WaitStartTimeOut = 0 Then
            If TestMode = 0 Then
                FailCounter1 = FailCounter1 + 1
            Else
                OffLFailCounter1 = OffLFailCounter1 + 1
            End If
        End If
        
         
        
        'Bin fail
        Select Case TestResult1
        
            Case "bin2" '(site1_bin2 = &H2)+(PowerStatus=&H80)  '10000010
                If TestMode = 0 Then
                    Bin2Counter1 = Bin2Counter1 + 1
                End If
                BinForm.Text11.Text = Bin2Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN2, Check2.value)
            Case "bin3"
                If TestMode = 0 Then
                    Bin3Counter1 = Bin3Counter1 + 1
                End If
                BinForm.Text12.Text = Bin3Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN3, Check2.value)
            Case "bin4"
                If TestMode = 0 Then
                    Bin4Counter1 = Bin4Counter1 + 1
                End If
                BinForm.Text13.Text = Bin4Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN4, Check2.value)
            Case "bin5"
                If TestMode = 0 Then
                    Bin5Counter1 = Bin5Counter1 + 1
                End If
                BinForm.Text14.Text = Bin5Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN5, Check2.value)
        End Select
        
            
            Label2 = "FAIL1"
            Label3.BackColor = RGB(255, 0, 0)
    End If

End If

 '=======================================Tester 1 FAIL Bin 'Check3.Value = 1=>不RT
        
If Check3.value = 1 And TestResult1 <> "PASS" Then
    If (((TestResult1 <> "PASS") And GetStart = 1) Or WaitStartTimeOut = 1) And site1 = 1 Then
         Print "\\\\\\\\\\site1 = "; TestResult1
            If WaitStartTimeOut = 0 Then
                If TestMode = 0 Then
                    FailCounter1 = FailCounter1 + 1
                Else
                    OffLFailCounter1 = OffLFailCounter1 + 1
                End If
            End If
        
        Select Case TestResult1
            Case "bin2"
                If TestMode = 0 Then
                    Bin2Counter1 = Bin2Counter1 + 1
                End If
                BinForm.Text11.Text = Bin2Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN2, Check2.value)
            Case "bin3"
                If TestMode = 0 Then
                    Bin3Counter1 = Bin3Counter1 + 1
                End If
                BinForm.Text12.Text = Bin3Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN3, Check2.value)
            Case "bin4"
                If TestMode = 0 Then
                    Bin4Counter1 = Bin4Counter1 + 1
                End If
                BinForm.Text13.Text = Bin4Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN4, Check2.value)
            Case "bin5"
                If TestMode = 0 Then
                    Bin5Counter1 = Bin5Counter1 + 1
                End If
                BinForm.Text14.Text = Bin5Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN5, Check2.value)
                
            Case Else
                If TestMode = 0 Then
                    Bin2Counter1 = Bin2Counter1 + 1
                End If
                BinForm.Text11.Text = Bin2Counter1
                Call PCI7248_bin(Channel_P1B, PCI7248_BIN2, Check2.value)
        End Select
            
        
        Label2 = "FAIL1"
        Label3.BackColor = RGB(255, 0, 0)
    End If

End If

   Call MsecDelay(0.2)
            

 '////////////////////////////////////// Tester 2 PASS Bin
        
If (TestResult2 = "PASS" Or RTTestResult2 = "PASS") And GetStart = 1 And site2 = 1 And gpon2 = "PASS" Then

        Print "\\\\\\\\\\site2 = "; TestResult2
        If TestResult2 = "PASS" Then
            
            If LoopTestCycle <> 0 Then
                LoopTestCounter2 = LoopTestCounter2 + 1
                HostForm.Print "Site2 LoopTest: " & LoopTestCounter2
                LoopTest2_Flag = True       'need Loop Test
                
                If LoopTestCounter2 = LoopTestCycle Then
                    LoopTestCounter2 = 0
                    LoopTest2_Flag = False
                
                    If TestMode = 0 Then
                        PassCounter2 = PassCounter2 + 1
                    Else
                        OffLPassCounter2 = OffLPassCounter2 + 1
                    End If
                
                End If
            Else
                LoopTest2_Flag = False
                If TestMode = 0 Then
                    PassCounter2 = PassCounter2 + 1
                Else
                    OffLPassCounter2 = OffLPassCounter2 + 1
                End If
            End If
            
            
            Print "\\\\\\\\\\\\\\ PASS 2"
            Label15 = "PASS2"
            Label16.BackColor = RGB(0, 255, 0)
        Else
            If TestMode = 0 Then
                RTPassCounter2 = RTPassCounter2 + 1
            Else
                OffLRTPassCounter2 = OffLRTPassCounter2 + 1
            End If
            Print "-------------- RTPASS 2"
            Label15 = "RTPASS2"
            Label16.BackColor = RGB(0, 0, 255)
        End If
        
        If LoopTest2_Flag = False Then
        
            If TestMode = 0 Then
                Bin1Counter2 = Bin1Counter2 + 1
            End If
            BinForm.Text24.Text = Bin1Counter2
            Call PCI7248_bin(Channel_P1A, PCI7248_PASS, Check2.value)
        
        End If
    
End If
        
        
'======================================= Tester 2 FAIL Bin 'Check3.Value = 0=>要RT
If Check3.value = 0 Then
     If (((TestResult2 <> "PASS" And RTTestResult2 <> "PASS") And GetStart = 1) Or WaitStartTimeOut = 1) And site2 = 1 Then
           
             Print "\\\\\\\\\\site2 = "; TestResult2
             
             If WaitStartTimeOut = 0 Then
                If TestMode = 0 Then
                    FailCounter2 = FailCounter2 + 1
                Else
                    OffLFailCounter2 = OffLFailCounter2 + 1
                End If
             End If
            
            Select Case TestResult2
            
                Case "bin2"
                    If TestMode = 0 Then
                        Bin2Counter2 = Bin2Counter2 + 1
                    End If
                    BinForm.Text25.Text = Bin2Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN2, Check2.value)
                Case "bin3"
                    If TestMode = 0 Then
                        Bin3Counter2 = Bin3Counter2 + 1
                    End If
                    BinForm.Text26.Text = Bin3Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN3, Check2.value)
                Case "bin4"
                    If TestMode = 0 Then
                        Bin4Counter2 = Bin4Counter2 + 1
                    End If
                    BinForm.Text27.Text = Bin4Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN4, Check2.value)
                Case "bin5"
                    If TestMode = 0 Then
                        Bin5Counter2 = Bin5Counter2 + 1
                    End If
                    BinForm.Text28.Text = Bin5Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN5, Check2.value)
            End Select
            
              Label15 = "FAIL!"
              Label16.BackColor = RGB(255, 0, 0)
             
                         
    End If
End If
        
'=========================================Tester 2 FAIL Bin 'Check3.Value = 1=>不RT
If Check3.value = 1 And TestResult2 <> "PASS" Then
    If (((TestResult2 <> "PASS") And GetStart = 1) Or WaitStartTimeOut = 1) And site2 = 1 Then
           
             Print "\\\\\\\\\\site2 = "; TestResult2
             
            If WaitStartTimeOut = 0 Then
                If TestMode = 0 Then
                    FailCounter2 = FailCounter2 + 1
                Else
                    OffLFailCounter2 = OffLFailCounter2 + 1
                End If
            End If
            
              
            
            Select Case TestResult2
            
                Case "bin2"
                    If TestMode = 0 Then
                        Bin2Counter2 = Bin2Counter2 + 1
                    End If
                    BinForm.Text25.Text = Bin2Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN2, Check2.value)
                Case "bin3"
                    If TestMode = 0 Then
                        Bin3Counter2 = Bin3Counter2 + 1
                    End If
                    BinForm.Text26.Text = Bin3Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN3, Check2.value)
                Case "bin4"
                    If TestMode = 0 Then
                        Bin4Counter2 = Bin4Counter2 + 1
                    End If
                    BinForm.Text27.Text = Bin4Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN4, Check2.value)
                Case "bin5"
                    If TestMode = 0 Then
                        Bin5Counter2 = Bin5Counter2 + 1
                    End If
                    BinForm.Text28.Text = Bin5Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN5, Check2.value)
                Case Else
                    If TestMode = 0 Then
                        Bin2Counter2 = Bin2Counter2 + 1
                    End If
                    BinForm.Text25.Text = Bin2Counter2
                    Call PCI7248_bin(Channel_P1A, PCI7248_BIN2, Check2.value)
                 
            End Select
           
              Label15 = "FAIL!"
              Label16.BackColor = RGB(255, 0, 0)
             
                         
    End If
End If
'===========================================================
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Reset  Latch
                  
        'atheist 2013/3/27   ContinueFail reset UPT2 power relay (P1CL:1,2) move to TestEnd (Receive Bin value)
        
        If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
            If (continuefail1 >= 3) Or (continuefail2 >= 3) Then
                Print "Reset UPT2 ......"
                k = DO_WritePort(card, Channel_P1CL, &H0)
                Call MsecDelay(0.2)
                k = DO_WritePort(card, Channel_P1CL, &HF)
                'Call MsecDelay(3#)                          'Wait UPT2 Initial
            End If
        End If
            
             Print "end state"
                   
               
       ' for check5 =============================Arch add 940529
        If (TestResult1 = "PASS" Or RTTestResult1 = "PASS") Then
            continuefail1 = 0
            continuefail1_bin2 = 0
            continuefail1_bin3 = 0
            continuefail1_bin4 = 0
            continuefail1_bin5 = 0
        ElseIf TestResult1 = "bin2" Or RTTestResult1 = "bin2" Then
            continuefail1 = continuefail1 + 1
            continuefail1_bin2 = continuefail1_bin2 + 1
        ElseIf TestResult1 = "bin3" Or RTTestResult1 = "bin3" Then
            continuefail1 = continuefail1 + 1
            continuefail1_bin3 = continuefail1_bin3 + 1
        ElseIf TestResult1 = "bin4" Or RTTestResult1 = "bin4" Then
            continuefail1 = continuefail1 + 1
            continuefail1_bin4 = continuefail1_bin4 + 1
        ElseIf TestResult1 = "bin5" Or RTTestResult1 = "bin5" Then
            continuefail1 = continuefail1 + 1
            continuefail1_bin5 = continuefail1_bin5 + 1
        End If
        
        If (TestResult2 = "PASS" Or RTTestResult2 = "PASS") Then
            continuefail2 = 0
            continuefail2_bin2 = 0
            continuefail2_bin3 = 0
            continuefail2_bin4 = 0
            continuefail2_bin5 = 0
        ElseIf TestResult2 = "bin2" Or RTTestResult2 = "bin2" Then
            continuefail2 = continuefail2 + 1
            continuefail2_bin2 = continuefail2_bin2 + 1
        ElseIf TestResult2 = "bin3" Or RTTestResult2 = "bin3" Then
            continuefail2 = continuefail2 + 1
            continuefail2_bin3 = continuefail2_bin3 + 1
        ElseIf TestResult2 = "bin4" Or RTTestResult2 = "bin4" Then
            continuefail2 = continuefail2 + 1
            continuefail2_bin4 = continuefail2_bin4 + 1
        ElseIf TestResult2 = "bin5" Or RTTestResult2 = "bin5" Then
            continuefail2 = continuefail2 + 1
            continuefail2_bin5 = continuefail2_bin5 + 1
        End If
        
        If ((continuefail1 >= 5) Or (continuefail2 >= 5)) And (InStr(ChipName, "U69") = 2) And (Mid(ChipName, 12, 1) <> "U") And (Len(ChipName) = 14) Then
            If (ChipName = "AU6981HLF30") Or (ChipName = "AU6981HLF28") Then
                'do nothing
            Else
                If (RealChipName = "") And (MPChipName = "") Then
                    SendMP_Flag = True
                    RealChipName = Trim(Combo1.Text) & Trim(Combo2.Text)
                    MPChipName = Left(ChipName, 10) & "M" & Right(ChipName, 3)
                End If
                ChipName = MPChipName
            End If
        ElseIf ((continuefail1 >= 5) Or (continuefail2 >= 5)) And (InStr(ChipName, "U87") = 2) And (Len(ChipName) = 15) Then
            If (RealChipName = "") And (MPChipName = "") Then
                    SendMP_Flag = True
                    RealChipName = Trim(Combo1.Text) & Trim(Combo2.Text)
                    MPChipName = Left(ChipName, 11) & "M" & Right(ChipName, 3)
                End If
            ChipName = MPChipName
        End If
        
        If (continuefail1 = 0) And (continuefail2 = 0) And (InStr(ChipName, "U87") = 2) And (Len(ChipName) = 15) Then
            SendMP_Flag = False
            ChipName = Trim(Combo1.Text) & Trim(Combo2.Text)
        End If
        
        If (continuefail1 = 0) And (continuefail2 = 0) And (InStr(ChipName, "U69") = 2) And (Mid(ChipName, 12, 1) <> "U") And (Len(ChipName) = 14) Then
            SendMP_Flag = False
            ChipName = Trim(Combo1.Text) & Trim(Combo2.Text)
        End If
'=======================================send testend to test for start loop?
   
    
   ' If Site1 = 1 Then
    
   ' MSComm1.Output = "testend"   ' trans end signal to TEST PC
   ' MSComm1.InBufferCount = 0
   ' MSComm1.InputLen = 0
   ' WaitForTest1 = Timer ' wait for timer  and test result
   ' End If
   ' buf1 = MSComm1.Input
    
   ' If Site2 = 1 Then
    
   ' MSComm2.Output = "testend"   ' trans end signal to TEST PC
   ' MSComm2.InBufferCount = 0
   ' MSComm2.InputLen = 0
   ' WaitForTest2 = Timer
   ' End If
   ' buf2 = MSComm2.Input
    
   ' Print "send end test signal to test"
     
    
   If LoopTestCycle <> 0 Then
        
        If ((Option1 = True) And ((TestResult1 <> "PASS") Or (TestResult2 <> "PASS"))) Then
            LoopTest1_Flag = False
            LoopTest2_Flag = False
            LoopTestCounter1 = 0
            LoopTestCounter2 = 0
            GoTo TestEnd
        ElseIf ((Option2 = True) And (TestResult1 <> "PASS")) Then
            LoopTest1_Flag = False
            LoopTestCounter1 = 0
            GoTo TestEnd
        ElseIf ((Option3 = True) And (TestResult2 <> "PASS")) Then
            LoopTest2_Flag = False
            LoopTestCounter2 = 0
            GoTo TestEnd
        End If
        
        If (Option1 = True) And (((LoopTest1_Flag = True) And (TestResult1 = "PASS")) Or ((TestResult2 = "PASS") And (LoopTest2_Flag = True))) _
        Or (Option2 = True) And (LoopTest1_Flag = True) And (TestResult1 = "PASS") _
        Or (Option3 = True) And (LoopTest2_Flag = True) And (TestResult2 = "PASS") Then
            GoTo LoopTest_Label
        End If
    
        If (Option1 = True) And ((LoopTest1_Flag = False) And (LoopTest2_Flag = False)) And (TestResult1 = "PASS") And (TestResult2 = "PASS") Then
            HostForm.Cls
            HostForm.Print "Site1 LoopTest: " & LoopTestCycle & " Cycle All PASS"
            HostForm.Print "Site2 LoopTest: " & LoopTestCycle & " Cycle All PASS"
        ElseIf (Option2 = True) And (LoopTest1_Flag = False) And (TestResult1 = "PASS") Then
            HostForm.Cls
            HostForm.Print "Site1 LoopTest: " & LoopTestCycle & " Cycle All PASS"
        ElseIf (Option3 = True) And (LoopTest2_Flag = False) And (TestResult1 = "PASS") Then
            HostForm.Cls
            HostForm.Print "Site2 LoopTest: " & LoopTestCycle & " Cycle All PASS"
        End If
   
   End If
   
TestEnd:
                   
             Bin1Site1 = Bin1Counter1
             Bin2Site1 = Bin2Counter1
             Bin3Site1 = Bin3Counter1
             Bin4Site1 = Bin4Counter1
             Bin5Site1 = Bin5Counter1
             
             
             Bin1Site2 = Bin1Counter2
             Bin2Site2 = Bin2Counter2
             Bin3Site2 = Bin3Counter2
             Bin4Site2 = Bin4Counter2
             Bin5Site2 = Bin5Counter2
             
             If SendMP_Flag = True Then
                ChipName = RealChipName
             End If
             
              Call UpdateDB
                   
             If SendMP_Flag = True Then
                ChipName = MPChipName
             End If
                   
                   
                   
               If flag = 0 Then      ' stop or one cycle
                    'ChipName = ""
                    'Label17.Caption = ""
                    'Label18.Caption = ""
                    'Label19.Caption = ""
                    'Command1.Enabled = False
                    'Command6.Enabled = False
                    'Combo1.Clear
                    'Combo2.Clear
                    'Label12.BackColor = &H8000000F
                    'Label14.BackColor = &H8000000F
                    'Label3.BackColor = &HFFFFFF
                    'Label16.BackColor = &HFFFFFF
                    Cls
              '      Call PrintReport
                    Print "STOP TEST!!!"
                    
                   Exit Sub
               End If
               
               
               If Check4.value = 1 Then     ' stop or one cycle
                    
                    Combo1.Enabled = True
                    Combo2.Enabled = True
                    Call UnlockOption
                    
                    'Label12.BackColor = &H8000000F
                    'Label14.BackColor = &H8000000F
                    'Label3.BackColor = &HFFFFFF
                    'Label16.BackColor = &HFFFFFF
                    Exit Sub
                   Print "STOP TEST!!!"
                 '   MsgBox "end state"
                   
               End If
               
               
               
Loop  'do1

   ' MSComm1.PortOpen = False '關閉序列埠
End Sub


 

Private Sub Command7_Click()
 Call PCI7248_bin(Channel_P1B, PCI7248_PASS, Check2.value)
 Call PCI7248_bin(Channel_P1A, PCI7248_PASS, Check2.value)
End Sub

Private Sub Command8_Click()
Dim PowerStatus As Byte
Dim TesterStatus1
Dim TesterStatus2

Dim TestMode As Byte

Dim WaitForStart
Dim WaitForVcc

Dim TotalRealTestTime
Dim OldTotalRealTestTime

Dim RealTestTime
Dim OldRealTestTime

Dim WaitStartTime

Dim GetStart As Integer
Dim TimeOut As Integer
 
Dim WaitStartCounter As Integer
Dim WaitStartTimeOutCounter As Integer
Dim WaitStartTimeOut As Integer


Dim TestCounter As Integer
Dim RTTestCounter1 As Integer
Dim RTTestCounter2 As Integer

Dim TestResult1
Dim TestResult2


Dim RTTestResult1
Dim RTTestResult2

Dim TestStop1 As Byte
Dim TestStop2 As Byte

Dim RTTestStop1 As Byte
Dim RTTestStop2 As Byte

Dim WaitForTest1
Dim WaitForTest2


Dim RTWaitForTest1
Dim RTWaitForTest2

Dim WaitTestTimeOutCounter1 As Integer
Dim WaitTestTimeOutCounter2 As Integer

Dim RTWaitTestTimeOutCounter1 As Integer
Dim RTWaitTestTimeOutCounter2 As Integer

Dim WaitTestTimeOut1 As Integer
Dim WaitTestTimeOut2 As Integer


Dim RTWaitTestTimeOut1 As Integer
Dim RTWaitTestTimeOut2 As Integer


Dim TestCycleTime1
Dim TestCycleTime2

Dim RTTestCycleTime1
Dim RTTestCycleTime2

Dim Bin1Counter1 As Integer
Dim Bin2Counter1 As Integer
Dim Bin3Counter1 As Integer
Dim Bin4Counter1 As Integer
Dim Bin5Counter1 As Integer
Dim Bin1Counter2 As Integer
Dim Bin2Counter2 As Integer
Dim Bin3Counter2 As Integer
Dim Bin4Counter2 As Integer
Dim Bin5Counter2 As Integer

Dim PassCounter1 As Integer
Dim FailCounter1 As Integer
Dim RTPassCounter1 As Integer
Dim RTFailCounter1 As Integer

Dim PassCounter2 As Integer
Dim FailCounter2 As Integer
Dim RTPassCounter2 As Integer
Dim RTFailCounter2 As Integer
Dim continuefail1 As Integer
Dim continuefail2 As Integer

Dim OffLPassCounter1 As Integer
Dim OffLFailCounter1 As Integer
Dim OffLRTPassCounter1 As Integer
Dim OffLRTFailCounter1 As Integer

Dim OffLPassCounter2 As Integer
Dim OffLFailCounter2 As Integer
Dim OffLRTPassCounter2 As Integer
Dim OffLRTFailCounter2 As Integer

Dim T1UnknownCounter As Integer
Dim T1SD_WFCounter As Integer
Dim T1SD_RFCounter As Integer
Dim T1CF_WFCounter As Integer
Dim T1CF_RFCounter As Integer
Dim T1XD_WFCounter As Integer
Dim T1XD_RFCounter As Integer
Dim T1SM_WFCounter As Integer
Dim T1SM_RFCounter As Integer

Dim T2UnknownCounter As Integer
Dim T2SD_WFCounter As Integer
Dim T2SD_RFCounter As Integer
Dim T2CF_WFCounter As Integer
Dim T2CF_RFCounter As Integer
Dim T2XD_WFCounter As Integer
Dim T2XD_RFCounter As Integer
Dim T2SM_WFCounter As Integer
Dim T2SM_RFCounter As Integer
Dim i As Integer, k As Integer
Dim result As Integer
Dim DO_P As Long
Dim DI_P As Long
Cls
result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
result = DIO_PortConfig(card, Channel_P1CH, INPUT_PORT)
result = DIO_PortConfig(card, Channel_P1CL, INPUT_PORT)

result = DIO_PortConfig(card, Channel_P2A, OUTPUT_PORT)
result = DIO_PortConfig(card, Channel_P2B, OUTPUT_PORT)
result = DIO_PortConfig(card, Channel_P2CH, OUTPUT_PORT)
result = DIO_PortConfig(card, Channel_P2CL, OUTPUT_PORT)

    If Check2.value = 1 Then
        PowerStatus = 0     'power on
    Else
        PowerStatus = 128  'power off
    End If

continuefail1 = 0
continuefail2 = 0
flag = 1
    
Command2.SetFocus

    DO_P = &HFF
    k = DO_WritePort(card, Channel_P1A, DO_P)
    DO_P = &HFF
    k = DO_WritePort(card, Channel_P1B, DO_P)
    i = 1
    
   Do
    Print "waiting start"
            Do     'wait  (VCC PowerON) & (hander 5ms start) signal
                 DoEvents
                 WaitStartTime = Timer - WaitForStart
                 k = DO_ReadPort(card, Channel_P1CH, DI_P)
                DI_P = 14 'temp 5/20
            Loop Until DI_P = 14 Or DI_P = 13 Or DI_P = 12 Or WaitStartTime > WAIT_START_TIME_OUT  'Allen
            
            Print "DI_P=", DI_P
            TotalRealTestTime = Timer - OldTotalRealTestTime
            OldTotalRealTestTime = Timer
            OldRealTestTime = Timer
    Print " get start"
     
     k = DO_WritePort(card, Channel_P2CL, 0)   'pull down
     Call MsecDelay(0.5)
     'If i Mod 5 = 0 Then
   
     'Call PCI7248_bin(Channel_P1B, PCI7248_PASS)
     'Call PCI7248_bin(Channel_P1A, PCI7248_BIN2)
     'Call MsecDelay(0.01)
     'Cls
     'ElseIf (i Mod 5 = 1) Then
    ' Call PCI7248_bin(Channel_P1B, PCI7248_BIN2)
     'Call PCI7248_bin(Channel_P1A, PCI7248_BIN3)
    ' Call MsecDelay(0.01)
    ' ElseIf (i Mod 5 = 2) Then
    ' Call PCI7248_bin(Channel_P1B, PCI7248_BIN3)
    'Call PCI7248_bin(Channel_P1A, PCI7248_BIN4)
    ' Call MsecDelay(0.01)
    ' ElseIf (i Mod 5 = 3) Then
    ' Call PCI7248_bin(Channel_P1B, PCI7248_BIN4)
    ' Call PCI7248_bin(Channel_P1A, PCI7248_BIN5)
     'Call MsecDelay(0.01)
    ' ElseIf (i Mod 5 = 4) Then
    ' Call PCI7248_bin(Channel_P1B, PCI7248_BIN5)
     'Call PCI7248_bin(Channel_P1A, PCI7248_PASS)
     'Call MsecDelay(0.01)
     'End If
   MsgBox "Channel_P2CL= 0"
   ' DO_P = &HFF
   ' k = DO_WritePort(card, Channel_P1A, DO_P)
   ' DO_P = &HFF
   ' k = DO_WritePort(card, Channel_P1B, DO_P)
    'k = DO_ReadPort(card, Channel_P1CH, DI_P)
i = i + 1
    Loop Until i = 2
    k = DO_WritePort(card, Channel_P2CL, 15)  'pull high
'Const PCI7248_EOT = &H1 'for 7248 card
'Const PCI7248_PASS = &HFD 'for 7248 card  11111101
'Const PCI7248_BIN2 = &HFB 'for 7248 card  11111011
'Const PCI7248_BIN3 = &HF7 'for 7248 card  11110111
'Const PCI7248_BIN4 = &HEF 'for 7248 card  11101111
'Const PCI7248_BIN5 = &HDF 'for 7248 card  11011111
    MsgBox "Channel_P2CL= 15"

End Sub



Private Sub Command5_Click()
Dim Bin1Counter1 As Integer
Dim Bin2Counter1 As Integer
Dim Bin3Counter1 As Integer
Dim Bin4Counter1 As Integer
Dim Bin5Counter1 As Integer
Dim Bin1Counter2 As Integer
Dim Bin2Counter2 As Integer
Dim Bin3Counter2 As Integer
Dim Bin4Counter2 As Integer
Dim Bin5Counter2 As Integer

Dim PassCounter1 As Integer
Dim FailCounter1 As Integer
Dim RTPassCounter1 As Integer
Dim RTFailCounter1 As Integer

Dim PassCounter2 As Integer
Dim FailCounter2 As Integer
Dim RTPassCounter2 As Integer
Dim RTFailCounter2 As Integer
Dim T1NotGetReadyCounter As Integer
Dim T1UnknownCounter As Integer
Dim T1GponFailCounter As Integer
Dim T1SD_WFCounter As Integer
Dim T1SD_RFCounter As Integer
Dim T1CF_WFCounter As Integer
Dim T1CF_RFCounter As Integer
Dim T1XD_WFCounter As Integer
Dim T1XD_RFCounter As Integer
Dim T1SM_WFCounter As Integer
Dim T1SM_RFCounter As Integer
Dim T2NotGetReadyCounter As Integer
Dim T2UnknownCounter As Integer
Dim T2GponFailCounter As Integer
Dim T2SD_WFCounter As Integer
Dim T2SD_RFCounter As Integer
Dim T2CF_WFCounter As Integer
Dim T2CF_RFCounter As Integer
Dim T2XD_WFCounter As Integer
Dim T2XD_RFCounter As Integer
Dim T2SM_WFCounter As Integer
Dim T2SM_RFCounter As Integer
    Bin1Counter1 = 0
    Bin2Counter1 = 0
    Bin3Counter1 = 0
    Bin4Counter1 = 0
    Bin5Counter1 = 0
    Bin1Counter2 = 0
    Bin2Counter2 = 0
    Bin3Counter2 = 0
    Bin4Counter2 = 0
    Bin5Counter2 = 0
    T1NotGetReadyCounter = 0
    T1UnknownCounter = 0
    T1GponFailCounter = 0
    T1SD_WFCounter = 0
    T1SD_RFCounter = 0
    T1CF_WFCounter = 0
    T1CF_RFCounter = 0
    T1XD_WFCounter = 0
    T1XD_RFCounter = 0
    T1SM_WFCounter = 0
    T1SM_RFCounter = 0
    T2NotGetReadyCounter = 0
    T2UnknownCounter = 0
    T2GponFailCounter = 0
    T2SD_WFCounter = 0
    T2SD_RFCounter = 0
    T2CF_WFCounter = 0
    T2CF_RFCounter = 0
    T2XD_WFCounter = 0
    T2XD_RFCounter = 0
    T2SM_WFCounter = 0
    T2SM_RFCounter = 0
End Sub

Private Sub Command6_Click()
On Error Resume Next
ReportBegin = 0
' report control  begin
Call ReportActive
If ReportCheck.value = 1 Then
Do
  DoEvents
Loop While ReportBegin = 0
End If

    avgTestTime = 0
    totalTestTime = 0
    testTime = 0
    UPH = 0
    Label35.Caption = "Avarage Test Time : "
    Label36.Caption = "UPH: "

ReportBegin = 0
Call LockOption

Dim Bin1Counter1 As Long
Dim Bin2Counter1 As Long
Dim Bin3Counter1 As Long
Dim Bin4Counter1 As Long
Dim Bin5Counter1 As Long
Dim Bin1Counter2 As Long
Dim Bin2Counter2 As Long
Dim Bin3Counter2 As Long
Dim Bin4Counter2 As Long
Dim Bin5Counter2 As Long


Dim PassCounter1 As Long
Dim PassCounter2 As Long

Dim OffLPassCounter1 As Long
Dim OffLPassCounter2 As Long

Dim RTPassCounter1 As Long
Dim RTPassCounter2 As Long

Dim OffLRTPassCounter1 As Long
Dim OffLRTPassCounter2 As Long

Dim FailCounter1 As Long
Dim FailCounter2 As Long

Dim OffLFailCounter1 As Long
Dim OffLFailCounter2 As Long

Dim RTFailCounter1 As Long
Dim RTFailCounter2 As Long

Dim OffLRTFailCounter1 As Long
Dim OffLRTFailCounter2 As Long

Dim RTTestCounter1 As Long
Dim RTTestCounter2 As Long

Dim OffLRTTestCounter1 As Long
Dim OffLRTTestCounter2 As Long

Dim debug1 As String
'====================================='設定IO PORT輸入輸出

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


'=====================================

continuefail1 = 0
continuefail1_bin2 = 0
continuefail1_bin3 = 0
continuefail1_bin4 = 0
continuefail1_bin5 = 0
continuefail2 = 0
continuefail2_bin2 = 0
continuefail2_bin3 = 0
continuefail2_bin4 = 0
continuefail2_bin5 = 0
flag = 1
SendMP_Flag = False
MPChipName = ""
RealChipName = ""

Command2.SetFocus

GetGPIBStatus(0) = False
GetGPIBStatus(1) = False

Cls
'=============================================' begin state
Print "begin state"

Combo1.Enabled = False
Combo2.Enabled = False

'///////////////////////////////////////////////////////////////
'
'                     MAIN LOOP
'
'///////////////////////////////////////////////////////////////
Do 'DO1
    Cls
    If ChipName = "" Then
        Label12.BackColor = &H8000000F
        Label14.BackColor = &H8000000F
        Cls
        MsgBox "Select Chip"
        Exit Sub
    End If
   
    If Option1.value = True Then '雙機
        site1 = 1
        site2 = 1
        Label12.BackColor = &H8080FF
        Label14.BackColor = &H8080FF
    End If
    
    
    If Option2.value = True Then '1號機
        site1 = 1
        site2 = 0
        Label12.BackColor = &H8080FF
        Label14.BackColor = &H8000000F
    End If
    
    
    If Option3.value = True Then '2號機
        site1 = 0
        site2 = 1
        Label12.BackColor = &H8000000F
        Label14.BackColor = &H8080FF
    End If

   
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Display status
   
    TestResult1 = ""
    TestResult2 = ""
    
    RTTestResult1 = ""
    RTTestResult2 = ""
    
    
    NoCardTestResult1 = ""
    NoCardTestResult2 = ""
    
    'Text1.Text = ""
    'Text2.Text = ""
    Text3.Text = WaitStartCounter
    Text4.Text = WaitStartTime
    Text5.Text = WaitStartTimeOutCounter
    
    If TestMode = 0 Then
        Text6.Text = TestCounter
        Text25.Text = RTTestCounter1
        Text26.Text = RTTestCounter2
    Else
        Text6.Text = OffLTestCounter
        Text25.Text = OffLRTTestCounter1
        Text26.Text = OffLRTTestCounter2
    End If
    
    If Check1.value = 1 Then
        TestMode = 1  '離線模式
        Text3.Text = "TestMode"
        Text4.Text = "TestMode"
        Text5.Text = "TestMode"
        Text6.Text = "TestMode"
        'Text6.Text = OffLTestCounter
    Else
        TestMode = 0  '上線模式
        Text3.Text = WaitStartCounter
        Text4.Text = WaitStartTime
        Text5.Text = WaitStartTimeOutCounter
        Text6.Text = TestCounter
    
    End If

     
    Text7.Text = TestCycleTime1
    
    If TestMode = 0 Then
        Text8.Text = PassCounter1
        Text9.Text = FailCounter1
    Else
        Text8.Text = OffLPassCounter1
        Text9.Text = OffLFailCounter1
    End If
    
    Text10.Text = WaitTestTimeOutCounter1
    
    Text17.Text = RTTestCycleTime1
    
    If TestMode = 0 Then
        Text18.Text = RTPassCounter1
        Text19.Text = RTFailCounter1
    Else
        Text18.Text = OffLRTPassCounter1
        Text19.Text = OffLRTFailCounter1
    End If
    
    Text20.Text = RTWaitTestTimeOutCounter1
    
    Text11.Text = TestCycleTime2
    
    If TestMode = 0 Then
        Text12.Text = PassCounter2
        Text13.Text = FailCounter2
    Else
        Text12.Text = OffLPassCounter2
        Text13.Text = OffLFailCounter2
    End If
    
    Text14.Text = WaitTestTimeOutCounter2
   
    Text21.Text = RTTestCycleTime2
    
    If TestMode = 0 Then
        Text22.Text = RTPassCounter2
        Text23.Text = RTFailCounter2
    Else
        Text22.Text = OffLRTPassCounter2
        Text23.Text = OffLRTFailCounter2
    End If
    
    Text24.Text = RTWaitTestTimeOutCounter2
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Initial Counter and Control variable
    
    
    TestCycleTime1 = 0
    TestCycleTime2 = 0
    
    
    RTTestCycleTime1 = 0
    RTTestCycleTime2 = 0
    
    NoCardTestCycleTime1 = 0
    NoCardTestCycleTime2 = 0
    ' inital time out flag


LoopTest_Label:
    
    GetStart = 0
    WaitStartTime = 0
    WaitStartTimeOut = 0
    
    WaitTestTimeOut1 = 0
    TestStop1 = 0
    TestStop1_1 = 0
    TestStop1_2 = 0
    TestStop1_3 = 0
    NoCardTestStop1 = 0
    
    WaitTestTimeOut2 = 0
    TestStop2 = 0
    TestStop2_1 = 0
    TestStop2_2 = 0
    TestStop2_3 = 0
    NoCardTestStop2 = 0
    
    RTTestCycleTime1 = 0
    RTTestCycleTime2 = 0
    
    RTWaitTestTimeOut1 = 0
    RTTestStop1 = 0
    
    RTWaitTestTimeOut2 = 0
    RTTestStop2 = 0
    
    DoEvents
'===================================================
Select Case ChipName
    
    Case "AU6366S4", "AU66S4_F"
         Call AU6366S4TestON         ' call HOST control program
    Case "AU6330", "AU6331", "AU6331SD", "AU6331MS", "AU6331NC", "AU6333SD", "AU6333NC", "AU6331F", "AU6333F"
        Call CommomdTestON
    Case "AU6363", "AU6366SD", "AU6366CF", "AU6366XD", "AU6366MS", "AU6366C"
        Call CommomdTestON
    Case "AU6367_1", "AU6367_2", "AU6367", "AU6367F"
         Call CommomdTestON
    Case "AU6377"
         Call CommomdTestON
    Case "AU6368_S", "AU6368", "AU6368N1", "AU6368NA", "AU6368NC", "AU6368NF", "AU6368F", "AU6368N2", "AU6368N4"
        Call CommomdTestON
    Case "AU6368PR", "AU6368PF"
        Call AU6368PROTestON
    Case "AU6369S2", "AU63692F", "AU6369S3", "AU63693F", "AU6369_1", "AU6369_2", "AU6369CF", "AU69CF_F", "AU6369XD"
        Call CommomdTestON
    Case "AU6373S2", "AU6375_1", "AU6375_2", "AU6375", "AU6375F"
        Call CommomdTestON
    Case "AU6384", "AU6385_1", "AU6385_2", "AU6386", "AU6386_D", "AU6386A3", "AU6388", "AU6389", "AU6386DF", "AU6388F", "AU6389F", "AU6980", "AU6980F"
        Call CommomdTestON
    Case "AU6390", "AU9520JJ", "AU9520N", "AU9520_1"
        Call CommomdTestON
    Case "AU9368_S", "AU9368_F", "AU9368", "AU9368_1", "AU9369S3", "AU9386", "AU9368_B"
        Call CommomdTestON
    Case "AU9510", "AU9520"
        Call SmartCardTestON
         'Call CommomdTestON
    Case "AU6980MB", "AU6610", "AU661048", "AU3130", "AU9520V4", "AU9520V5", "AU6254", "AU6254A"
           Call CommomdTestON
    
     Case "AU6333HC", "AU6334HC", "AU6366HC", "AU6368HC", "AU6375HC", "AU6377HC"
        
           Call CommomdTestON
    
     Case "AU6333AS", "AU6334AS", "AU6366AS", "AU6368AS", "AU6375AS", "AU6377AS"
             
          Call CommomdTestON
          
    Case Else
          If Left(ChipName, 2) = "AU" Then
             Call CommomdTestON
          Else
             MsgBox "ChipName error"
             Exit Sub
          End If
          
End Select
'=======================================================
'////////////////////////////////////// Tester 1 PASS Bin

err:
    ' default value
       gpon1 = "PASS"
       gpon2 = "PASS"
      
      
    If Check6.value = 1 Then     'check GPON7_LED & Power_LED
    
       gpon1 = ""
       gpon2 = ""
    
        Dim DI_Power As Long
        k = DO_ReadPort(card, Channel_P2CL, DI_Power)
        Print "DI_Power="; DI_Power
        
        If TestResult1 = "PASS" Then
            If site1 = 1 And (DI_Power Mod 4 = 0) Then
                gpon1 = "PASS"
                Print "gpon1="; gpon1
            Else
                gpon1 = "FAIL"
                Print "gpon1="; gpon1
                 TestResult1 = "gponFail"
            End If
        End If
        
        If TestResult2 = "PASS" Then
            If site2 = 1 And (DI_Power <= 3) Then
                gpon2 = "PASS"
                Print "gpon2="; gpon2
            Else
                gpon2 = "FAIL"
                Print "gpon2="; gpon2
                 TestResult2 = "gponFail"
            End If
        End If
        
        
    End If
    

    

Label3 = TestResult1
Label16 = TestResult2

Print "close power"

If Check2.value = 0 Then  'Check2.Value = 1 => continuous supply power
    
        'If ChipName = "AU6366S4" Then
      '  If ChipName = "AU6366S4" Or ChipName = "AU66S4_F" Or ChipName = "AU9520" Then
      '      k = DO_ReadPort(card, Channel_P1CL, DI_S1)
      '      k = DO_ReadPort(card, Channel_P1CH, DI_S2)
      '  Else
      '      k = DO_WritePort(card, Channel_P1CL, &HF) ' send 1111 => "SetSITE1 Power OFF" & " SetSITE1 CDN HIGH"Channel_P1CL = 15
      '      k = DO_WritePort(card, Channel_P1CH, &HF) ' send 1111 => "SetSITE2 Power OFF" & " SetSITE2 CDN HIGH"Channel_P1CH = 15
      '  End If
    k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off"& " SetSITE2 CDN high" Channel_P2A = 255
    k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off"& " SetSITE1 CDN high" Channel_P2A = 255
                   
End If

If site1 = 1 Then
 


 Select Case TestResult1
        Case "PASS"
            TestResult1 = "PASS"
        Case "UNKNOW", "bin2", "Bin2"
            TestResult1 = "bin2"
            T1UnknownCounter = T1UnknownCounter + 1
            BinForm.Text1.Text = T1UnknownCounter
        Case "gponFail", "bin3", "Bin3"
            TestResult1 = "bin3"
            T1GponFailCounter = T1GponFailCounter + 1
            BinForm.Text29.Text = T1GponFailCounter
        Case "SD_WF"
            TestResult1 = "bin3"
            T1SD_WFCounter = T1SD_WFCounter + 1
            BinForm.Text2.Text = T1SD_WFCounter
        Case "SD_RF"
            TestResult1 = "bin3"
            T1SD_RFCounter = T1SD_RFCounter + 1
            BinForm.Text3.Text = T1SD_RFCounter
        Case "CF_WF"
            TestResult1 = "bin3"
            T1CF_WFCounter = T1CF_WFCounter + 1
            BinForm.Text4.Text = T1CF_WFCounter
        Case "CF_RF"
            TestResult1 = "bin3"
            T1CF_RFCounter = T1CF_RFCounter + 1
            BinForm.Text5.Text = T1CF_RFCounter
        Case "XD_WF", "bin4", "Bin4"
            TestResult1 = "bin4"
            T1XD_WFCounter = T1XD_WFCounter + 1
            BinForm.Text6.Text = T1XD_WFCounter
        Case "XD_RF"
            TestResult1 = "bin4"
            T1XD_RFCounter = T1XD_RFCounter + 1
            BinForm.Text7.Text = T1XD_RFCounter
        Case "MS_WF", "bin5", "TimeOut", "Bin5"
            TestResult1 = "bin5"
            T1SM_WFCounter = T1SM_WFCounter + 1
             BinForm.Text8.Text = T1SM_WFCounter
        Case "MS_RF"
            TestResult1 = "bin5"
            T1SM_RFCounter = T1SM_RFCounter + 1
            BinForm.Text9.Text = T1SM_RFCounter
        Case Else
            TestResult1 = "bin2"
            T1UnknownCounter = T1UnknownCounter + 1
            BinForm.Text1.Text = T1UnknownCounter
        End Select

End If
If site2 = 1 Then


Select Case TestResult2
        Case "PASS"
            TestResult2 = "PASS"
        Case "UNKNOW", "bin2", "Bin2"
            TestResult2 = "bin2"
            T2UnknownCounter = T2UnknownCounter + 1
            BinForm.Text15.Text = T2UnknownCounter
        Case "gponFail", "bin3", "Bin3"
            TestResult2 = "bin3"
            T2GponFailCounter = T2GponFailCounter + 1
            BinForm.Text30.Text = T2GponFailCounter
        Case "SD_WF"
            TestResult2 = "bin3"
            T2SD_WFCounter = T2SD_WFCounter + 1
            BinForm.Text16.Text = T2SD_WFCounter
        Case "SD_RF"
            TestResult2 = "bin3"
            T2SD_RFCounter = T2SD_RFCounter + 1
            BinForm.Text17.Text = T2SD_RFCounter
        Case "CF_WF"
            TestResult2 = "bin3"
            T2CF_WFCounter = T2CF_WFCounter + 1
            BinForm.Text18.Text = T2CF_WFCounter
        Case "CF_RF"
            TestResult2 = "bin3"
            T2CF_RFCounter = T2CF_RFCounter + 1
            BinForm.Text19.Text = T1CF_RFCounter
        Case "XD_WF", "bin4", "Bin4"
            TestResult2 = "bin4"
            T2XD_WFCounter = T2XD_WFCounter + 1
            BinForm.Text20.Text = T2XD_WFCounter
        Case "XD_RF"
            TestResult2 = "bin4"
            T2XD_RFCounter = T2XD_RFCounter + 1
            BinForm.Text21.Text = T2XD_RFCounter
        Case "MS_WF", "bin5", "TimeOut", "Bin5"
            TestResult2 = "bin5"
            T2SM_WFCounter = T2SM_WFCounter + 1
             BinForm.Text22.Text = T2SM_WFCounter
        Case "MS_RF"
            TestResult2 = "bin5"
            T2SM_RFCounter = T2SM_RFCounter + 1
            BinForm.Text23.Text = T2SM_RFCounter
        Case Else
            TestResult2 = "bin2"
            T2UnknownCounter = T2UnknownCounter + 1
            BinForm.Text15.Text = T2UnknownCounter
        End Select

End If

If GreaTekChipName = "AU6368A1" Then  ' For GTK AU6368A1 sorting case
    
     If TestResult1 = "bin2" Then
        TestResult1 = "bin4"
     End If
        
     If TestResult1 = "bin3" Then
        TestResult1 = "bin5"
     End If


    If TestResult2 = "bin2" Then
        TestResult2 = "bin4"
     End If
        
     If TestResult2 = "bin3" Then
        TestResult2 = "bin5"
     End If

End If


If testTime = 0 Or SendMP_Flag = True Then
    
    If testTime = 0 Then
        testTime = testTime + 1
    End If
    
    RealTestTime = Timer - OldRealTestTime
    If SPILFlag Then
        Label27.Caption = "Real Test Time :" & RealTestTime & " s"
    Else
        Label27.Caption = "實際測試時間(不含 load / unload)  :" & RealTestTime & " s"
    End If
Else
    RealTestTime = Timer - OldRealTestTime
    totalTestTime = totalTestTime + RealTestTime
    avgTestTime = totalTestTime / testTime
    Debug.Print avgTestTime
    Label35.Caption = "Avarage Test Time : " & avgTestTime & " s, 測試數： " & testTime
    testTime = testTime + 1
 
    If SPILFlag Then
        Label27.Caption = "Real Test Time :" & RealTestTime & " s"
    Else
        Label27.Caption = "實際測試時間(不含 load / unload)  :" & RealTestTime & " s"
    End If
End If

Label36.Caption = "UPH: " & (3600 / avgTestTime)
 
If (TestResult1 = "PASS" Or RTTestResult1 = "PASS") And GetStart = 1 And site1 = 1 And gpon1 = "PASS" Then

        Print "\\\\\\\\\\site1 = "; TestResult1
              
        If TestResult1 = "PASS" Then
            If LoopTestCycle <> 0 Then
                'HostForm.Cls
                LoopTestCounter1 = LoopTestCounter1 + 1
                HostForm.Print "Site1 LoopTest: " & LoopTestCounter1
                LoopTest1_Flag = True       'need Loop Test
                
                If LoopTestCounter1 = LoopTestCycle Then
                    LoopTestCounter1 = 0
                    LoopTest1_Flag = False
                    
                    If TestMode = 0 Then
                        PassCounter1 = PassCounter1 + 1
                    Else
                        OffLPassCounter1 = OffLPassCounter1 + 1
                    End If
                
                End If
            Else
                LoopTest1_Flag = False
                If TestMode = 0 Then
                    PassCounter1 = PassCounter1 + 1
                Else
                    OffLPassCounter1 = OffLPassCounter1 + 1
                End If
            End If
            
            Label2 = "PASS1"
            Label3.BackColor = RGB(0, 255, 0)
        Else
            If TestMode = 0 Then
                RTPassCounter1 = RTPassCounter1 + 1
            Else
                OffLRTPassCounter1 = OffLRTPassCounter1 + 1
            End If
            Print "-------------- RTPASS 1"
            Label2 = "RTPASS1"
            Label3.BackColor = RGB(0, 0, 255)
        End If
        
        If LoopTest1_Flag = False Then
            
            If TestMode = 0 Then
                Bin1Counter1 = Bin1Counter1 + 1
            End If
            
            BinForm.Text10.Text = Bin1Counter1
        
            Call PCI7248_bin(Channel_P2B, PCI7248_PASS, Check2.value)
        End If
       
       
End If
            
'////////////////////////////////////// Tester 1 FAIL Bin
If Check3.value = 0 Then
    If (((TestResult1 <> "PASS" And RTTestResult1 <> "PASS") And GetStart = 1) Or WaitStartTimeOut = 1) And site1 = 1 Then
    
    
         Print "\\\\\\\\\\site1 = "; TestResult1
        
        If WaitStartTimeOut = 0 Then
            If TestMode = 0 Then
                FailCounter1 = FailCounter1 + 1
            Else
                OffLFailCounter1 = OffLFailCounter1 + 1
            End If
        End If
        
         
        
        'Bin fail
        Select Case TestResult1
        
            Case "bin2" '(site1_bin2 = &H2)+(PowerStatus=&H80)  '10000010
                If TestMode = 0 Then
                    Bin2Counter1 = Bin2Counter1 + 1
                End If
                BinForm.Text11.Text = Bin2Counter1
                Call PCI7248_bin(Channel_P2B, PCI7248_BIN2, Check2.value)
            Case "bin3"
                If TestMode = 0 Then
                    Bin3Counter1 = Bin3Counter1 + 1
                End If
                BinForm.Text12.Text = Bin3Counter1
                Call PCI7248_bin(Channel_P2B, PCI7248_BIN3, Check2.value)
            Case "bin4"
                If TestMode = 0 Then
                    Bin4Counter1 = Bin4Counter1 + 1
                End If
                BinForm.Text13.Text = Bin4Counter1
                Call PCI7248_bin(Channel_P2B, PCI7248_BIN4, Check2.value)
            Case "bin5"
                If TestMode = 0 Then
                    Bin5Counter1 = Bin5Counter1 + 1
                End If
                BinForm.Text14.Text = Bin5Counter1
                Call PCI7248_bin(Channel_P2B, PCI7248_BIN5, Check2.value)
        End Select
        
            
            Label2 = "FAIL1"
            Label3.BackColor = RGB(255, 0, 0)
    End If

End If

 '=======================================Tester 1 FAIL Bin 'Check3.Value = 1=>不RT
        
If Check3.value = 1 And TestResult1 <> "PASS" Then
    If (((TestResult1 <> "PASS") And GetStart = 1) Or WaitStartTimeOut = 1) And site1 = 1 Then
         Print "\\\\\\\\\\site1 = "; TestResult1
            If WaitStartTimeOut = 0 Then
                If TestMode = 0 Then
                    FailCounter1 = FailCounter1 + 1
                Else
                    OffLFailCounter1 = OffLFailCounter1 + 1
                End If
            End If
        
        Select Case TestResult1
            Case "bin2"
                If TestMode = 0 Then
                    Bin2Counter1 = Bin2Counter1 + 1
                End If
                BinForm.Text11.Text = Bin2Counter1
                Call PCI7248_bin(Channel_P2B, PCI7248_BIN2, Check2.value)
            Case "bin3"
                If TestMode = 0 Then
                    Bin3Counter1 = Bin3Counter1 + 1
                End If
                BinForm.Text12.Text = Bin3Counter1
                Call PCI7248_bin(Channel_P2B, PCI7248_BIN3, Check2.value)
            Case "bin4"
                If TestMode = 0 Then
                    Bin4Counter1 = Bin4Counter1 + 1
                End If
                BinForm.Text13.Text = Bin4Counter1
                Call PCI7248_bin(Channel_P2B, PCI7248_BIN4, Check2.value)
            Case "bin5"
                If TestMode = 0 Then
                    Bin5Counter1 = Bin5Counter1 + 1
                End If
                BinForm.Text14.Text = Bin5Counter1
                Call PCI7248_bin(Channel_P2B, PCI7248_BIN5, Check2.value)
                
            Case Else
                If TestMode = 0 Then
                    Bin2Counter1 = Bin2Counter1 + 1
                End If
                BinForm.Text11.Text = Bin2Counter1
                Call PCI7248_bin(Channel_P2B, PCI7248_BIN2, Check2.value)
        End Select
            
        
        Label2 = "FAIL1"
        Label3.BackColor = RGB(255, 0, 0)
    End If

End If

   Call MsecDelay(0.2)
            

 '////////////////////////////////////// Tester 2 PASS Bin
        
If (TestResult2 = "PASS" Or RTTestResult2 = "PASS") And GetStart = 1 And site2 = 1 And gpon2 = "PASS" Then

        Print "\\\\\\\\\\site2 = "; TestResult2
        If TestResult2 = "PASS" Then
            
            If LoopTestCycle <> 0 Then
                'HostForm.Cls
                LoopTestCounter2 = LoopTestCounter2 + 1
                HostForm.Print "Site2 LoopTest: " & LoopTestCounter2
                LoopTest2_Flag = True       'need Loop Test
                
                If LoopTestCounter2 = LoopTestCycle Then
                    LoopTestCounter2 = 0
                    LoopTest2_Flag = False
                    
                    If TestMode = 0 Then
                        PassCounter2 = PassCounter2 + 1
                    Else
                        OffLPassCounter2 = OffLPassCounter2 + 1
                    End If
                
                End If
            Else
                LoopTest2_Flag = False
                If TestMode = 0 Then
                    PassCounter2 = PassCounter2 + 1
                Else
                    OffLPassCounter2 = OffLPassCounter2 + 1
                End If
            End If
            
            
            Print "\\\\\\\\\\\\\\ PASS 2"
            Label15 = "PASS2"
            Label16.BackColor = RGB(0, 255, 0)
        Else
            If TestMode = 0 Then
                RTPassCounter2 = RTPassCounter2 + 1
            Else
                OffLRTPassCounter2 = OffLRTPassCounter2 + 1
            End If
            Print "-------------- RTPASS 2"
            Label15 = "RTPASS2"
            Label16.BackColor = RGB(0, 0, 255)
        End If
        
        If LoopTest2_Flag = False Then
            If TestMode = 0 Then
                Bin1Counter2 = Bin1Counter2 + 1
            End If
            BinForm.Text24.Text = Bin1Counter2
            Call PCI7248_bin(Channel_P2A, PCI7248_PASS, Check2.value)
        End If
    
End If
        
        
'======================================= Tester 2 FAIL Bin 'Check3.Value = 0=>要RT
If Check3.value = 0 Then
     If (((TestResult2 <> "PASS" And RTTestResult2 <> "PASS") And GetStart = 1) Or WaitStartTimeOut = 1) And site2 = 1 Then
           
             Print "\\\\\\\\\\site2 = "; TestResult2
             
            If WaitStartTimeOut = 0 Then
              If TestMode = 0 Then
                  FailCounter2 = FailCounter2 + 1
              Else
                  OffLFailCounter2 = OffLFailCounter2 + 1
              End If
            End If
            
            Select Case TestResult2
            
                Case "bin2"
                    If TestMode = 0 Then
                        Bin2Counter2 = Bin2Counter2 + 1
                    End If
                    BinForm.Text25.Text = Bin2Counter2
                    Call PCI7248_bin(Channel_P2A, PCI7248_BIN2, Check2.value)
                Case "bin3"
                    If TestMode = 0 Then
                        Bin3Counter2 = Bin3Counter2 + 1
                    End If
                    BinForm.Text26.Text = Bin3Counter2
                    Call PCI7248_bin(Channel_P2A, PCI7248_BIN3, Check2.value)
                Case "bin4"
                    If TestMode = 0 Then
                        Bin4Counter2 = Bin4Counter2 + 1
                    End If
                    BinForm.Text27.Text = Bin4Counter2
                    Call PCI7248_bin(Channel_P2A, PCI7248_BIN4, Check2.value)
                Case "bin5"
                    If TestMode = 0 Then
                        Bin5Counter2 = Bin5Counter2 + 1
                    End If
                    BinForm.Text28.Text = Bin5Counter2
                    Call PCI7248_bin(Channel_P2A, PCI7248_BIN5, Check2.value)
            End Select
            
              Label15 = "FAIL!"
              Label16.BackColor = RGB(255, 0, 0)
             
                         
    End If
End If
        
'=========================================Tester 2 FAIL Bin 'Check3.Value = 1=>不RT
If Check3.value = 1 And TestResult2 <> "PASS" Then
    If (((TestResult2 <> "PASS") And GetStart = 1) Or WaitStartTimeOut = 1) And site2 = 1 Then
           
             Print "\\\\\\\\\\site2 = "; TestResult2
             
            If WaitStartTimeOut = 0 Then
                If TestMode = 0 Then
                    FailCounter2 = FailCounter2 + 1
                Else
                    OffLFailCounter2 = OffLFailCounter2 + 1
                End If
            End If
            
              
            
            Select Case TestResult2
            
                Case "bin2"
                    If TestMode = 0 Then
                        Bin2Counter2 = Bin2Counter2 + 1
                    End If
                    BinForm.Text25.Text = Bin2Counter2
                    Call PCI7248_bin(Channel_P2A, PCI7248_BIN2, Check2.value)
                Case "bin3"
                    If TestMode = 0 Then
                        Bin3Counter2 = Bin3Counter2 + 1
                    End If
                    BinForm.Text26.Text = Bin3Counter2
                    Call PCI7248_bin(Channel_P2A, PCI7248_BIN3, Check2.value)
                Case "bin4"
                    If TestMode = 0 Then
                        Bin4Counter2 = Bin4Counter2 + 1
                    End If
                    BinForm.Text27.Text = Bin4Counter2
                    Call PCI7248_bin(Channel_P2A, PCI7248_BIN4, Check2.value)
                Case "bin5"
                    If TestMode = 0 Then
                        Bin5Counter2 = Bin5Counter2 + 1
                    End If
                    BinForm.Text28.Text = Bin5Counter2
                    Call PCI7248_bin(Channel_P2A, PCI7248_BIN5, Check2.value)
                    
                  Case Else
                    If TestMode = 0 Then
                        Bin2Counter2 = Bin2Counter2 + 1
                    End If
                    BinForm.Text25.Text = Bin2Counter2
                    Call PCI7248_bin(Channel_P2A, PCI7248_BIN2, Check2.value)
                 
            End Select
           
              Label15 = "FAIL!"
              Label16.BackColor = RGB(255, 0, 0)
             
                         
    End If
End If
'===========================================================
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Reset  Latch
                  
        'atheist 2013/3/27   ContinueFail reset UPT2 power relay (P1CL:1,2) move to TestEnd (Receive Bin value)
        
        If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
            If (continuefail1 >= 3) Or (continuefail2 >= 3) Then
                Print "Reset UPT2 ......"
                k = DO_WritePort(card, Channel_P1CL, &H0)
                Call MsecDelay(0.2)
                k = DO_WritePort(card, Channel_P1CL, &HF)
                'Call MsecDelay(3#)                          'Wait UPT2 Initial
            End If
        End If
        
            
        Print "end state"
                   
              
               
       ' for check5 =============================Arch add 940529
        If (TestResult1 = "PASS" Or RTTestResult1 = "PASS") Then
            continuefail1 = 0
            continuefail1_bin2 = 0
            continuefail1_bin3 = 0
            continuefail1_bin4 = 0
            continuefail1_bin5 = 0
        ElseIf TestResult1 = "bin2" Or RTTestResult1 = "bin2" Then
            continuefail1 = continuefail1 + 1
            continuefail1_bin2 = continuefail1_bin2 + 1
        ElseIf TestResult1 = "bin3" Or RTTestResult1 = "bin3" Then
            continuefail1 = continuefail1 + 1
            continuefail1_bin3 = continuefail1_bin3 + 1
        ElseIf TestResult1 = "bin4" Or RTTestResult1 = "bin4" Then
            continuefail1 = continuefail1 + 1
            continuefail1_bin4 = continuefail1_bin4 + 1
        ElseIf TestResult1 = "bin5" Or RTTestResult1 = "bin5" Then
            continuefail1 = continuefail1 + 1
            continuefail1_bin5 = continuefail1_bin5 + 1
        End If
        
        If (TestResult2 = "PASS" Or RTTestResult2 = "PASS") Then
            continuefail2 = 0
            continuefail2_bin2 = 0
            continuefail2_bin3 = 0
            continuefail2_bin4 = 0
            continuefail2_bin5 = 0
        ElseIf TestResult2 = "bin2" Or RTTestResult2 = "bin2" Then
            continuefail2 = continuefail2 + 1
            continuefail2_bin2 = continuefail2_bin2 + 1
        ElseIf TestResult2 = "bin3" Or RTTestResult2 = "bin3" Then
            continuefail2 = continuefail2 + 1
            continuefail2_bin3 = continuefail2_bin3 + 1
        ElseIf TestResult2 = "bin4" Or RTTestResult2 = "bin4" Then
            continuefail2 = continuefail2 + 1
            continuefail2_bin4 = continuefail2_bin4 + 1
        ElseIf TestResult2 = "bin5" Or RTTestResult2 = "bin5" Then
            continuefail2 = continuefail2 + 1
            continuefail2_bin5 = continuefail2_bin5 + 1
        End If
       
        If ((continuefail1 >= 5) Or (continuefail2 >= 5)) And (InStr(ChipName, "U69") = 2) And (Mid(ChipName, 12, 1) <> "U") And (Len(ChipName) = 14) Then
            If (ChipName = "AU6981HLF30") Or (ChipName = "AU6981HLF28") Then
                'do nothing
            Else
                If (RealChipName = "") And (MPChipName = "") Then
                    SendMP_Flag = True
                    RealChipName = Trim(Combo1.Text) & Trim(Combo2.Text)
                    MPChipName = Left(ChipName, 10) & "M" & Right(ChipName, 3)
                End If
                ChipName = MPChipName
            End If
        ElseIf ((continuefail1 >= 5) Or (continuefail2 >= 5)) And (InStr(ChipName, "U87") = 2) And (Len(ChipName) = 15) Then
            If (RealChipName = "") And (MPChipName = "") Then
                    SendMP_Flag = True
                    RealChipName = Trim(Combo1.Text) & Trim(Combo2.Text)
                    MPChipName = Left(ChipName, 11) & "M" & Right(ChipName, 3)
                End If
            ChipName = MPChipName
        End If
        
        If (continuefail1 = 0) And (continuefail2 = 0) And (InStr(ChipName, "U87") = 2) And (Len(ChipName) = 15) Then
            SendMP_Flag = False
            ChipName = Trim(Combo1.Text) & Trim(Combo2.Text)
        End If
        
        If (continuefail1 = 0) And (continuefail2 = 0) And (InStr(ChipName, "U69") = 2) And (Mid(ChipName, 12, 1) <> "U") And (Len(ChipName) = 14) Then
            SendMP_Flag = False
            ChipName = Trim(Combo1.Text) & Trim(Combo2.Text)
        End If
        
'=======================================send testend to test for start loop?
   
    
   ' If Site1 = 1 Then
    
   ' MSComm1.Output = "testend"   ' trans end signal to TEST PC
   ' MSComm1.InBufferCount = 0
   ' MSComm1.InputLen = 0
   ' WaitForTest1 = Timer ' wait for timer  and test result
   ' End If
   ' buf1 = MSComm1.Input
    
   ' If Site2 = 1 Then
    
   ' MSComm2.Output = "testend"   ' trans end signal to TEST PC
   ' MSComm2.InBufferCount = 0
   ' MSComm2.InputLen = 0
   ' WaitForTest2 = Timer
   ' End If
   ' buf2 = MSComm2.Input
    
   ' Print "send end test signal to test"
    
    If LoopTestCycle <> 0 Then
        
        If ((Option1 = True) And ((TestResult1 <> "PASS") Or (TestResult2 <> "PASS"))) Then
            LoopTest1_Flag = False
            LoopTest2_Flag = False
            LoopTestCounter1 = 0
            LoopTestCounter2 = 0
            GoTo TestEnd
        ElseIf ((Option2 = True) And (TestResult1 <> "PASS")) Then
            LoopTest1_Flag = False
            LoopTestCounter1 = 0
            GoTo TestEnd
        ElseIf ((Option3 = True) And (TestResult2 <> "PASS")) Then
            LoopTest2_Flag = False
            LoopTestCounter2 = 0
            GoTo TestEnd
        End If
        
        If (Option1 = True) And (((LoopTest1_Flag = True) And (TestResult1 = "PASS")) Or ((TestResult2 = "PASS") And (LoopTest2_Flag = True))) _
        Or (Option2 = True) And (LoopTest1_Flag = True) And (TestResult1 = "PASS") _
        Or (Option3 = True) And (LoopTest2_Flag = True) And (TestResult2 = "PASS") Then
            GoTo LoopTest_Label
        End If
    
        If (Option1 = True) And ((LoopTest1_Flag = False) And (LoopTest2_Flag = False)) And (TestResult1 = "PASS") And (TestResult2 = "PASS") Then
            HostForm.Cls
            HostForm.Print "Site1 LoopTest: " & LoopTestCycle & " Cycle All PASS"
            HostForm.Print "Site2 LoopTest: " & LoopTestCycle & " Cycle All PASS"
        ElseIf (Option2 = True) And (LoopTest1_Flag = False) And (TestResult1 = "PASS") Then
            HostForm.Cls
            HostForm.Print "Site1 LoopTest: " & LoopTestCycle & " Cycle All PASS"
        ElseIf (Option3 = True) And (LoopTest2_Flag = False) And (TestResult1 = "PASS") Then
            HostForm.Cls
            HostForm.Print "Site2 LoopTest: " & LoopTestCycle & " Cycle All PASS"
        End If
   
   End If


TestEnd:
             Bin1Site1 = Bin1Counter1
             Bin2Site1 = Bin2Counter1
             Bin3Site1 = Bin3Counter1
             Bin4Site1 = Bin4Counter1
             Bin5Site1 = Bin5Counter1
             
             
             Bin1Site2 = Bin1Counter2
             Bin2Site2 = Bin2Counter2
             Bin3Site2 = Bin3Counter2
             Bin4Site2 = Bin4Counter2
             Bin5Site2 = Bin5Counter2
            
             If SendMP_Flag = True Then
                ChipName = RealChipName
             End If
             
             Call UpdateDB
            
             If SendMP_Flag = True Then
                ChipName = MPChipName
             End If
                   
               If flag = 0 Then      ' stop or one cycle
                    'ChipName = ""
                    'Label17.Caption = ""
                    'Label18.Caption = ""
                    'Label19.Caption = ""
                    'Command1.Enabled = False
                    'Command6.Enabled = False
                    'Combo1.Clear
                    'Combo2.Clear
                    'Label12.BackColor = &H8000000F
                    'Label14.BackColor = &H8000000F
                    'Label3.BackColor = &HFFFFFF
                    'Label16.BackColor = &HFFFFFF
                    Cls
                 '   Call PrintReport
                    Print "STOP TEST!!!"
                   Exit Sub
               End If
               
               
               If Check4.value = 1 Then     ' stop or one cycle
                    
                    Combo1.Enabled = True
                    Combo2.Enabled = True
                    Call UnlockOption
                    'Label12.BackColor = &H8000000F
                    'Label14.BackColor = &H8000000F
                    'Label3.BackColor = &HFFFFFF
                    'Label16.BackColor = &HFFFFFF
                    Exit Sub
                   Print "STOP TEST!!!"
                 '   MsgBox "end state"
                   
               End If
               
Loop  'do1

   ' MSComm1.PortOpen = False '關閉序列埠
End Sub


Private Sub Command9_Click()
Call GetProcessIDSub
End Sub

Private Sub Form_Load()

On Error Resume Next

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

    avgTestTime = 0
    totalTestTime = 0
    testTime = 0
    UPH = 0

    If App.EXEName <> "Host" Or CheckMe(Me) Then
        End
    End If

    ReportCheck = 1  ' default for report function
    BinForm.Show
        '設置或返回序列埠號

    MSComm1.CommPort = 1
    '===============================
        '"串列傳輸速率,同位元檢查方式,資料位元數,停止位元數"33
    MSComm1.Settings = "9600,N,8,1"
    '===============================3
        '打開或關閉序列埠
    If MSComm1.PortOpen = False Then
    MSComm1.PortOpen = True
    End If
    '===============================
        '返回接收緩衝區內等待取讀取的位元組數,屬性為0表示清空接收緩衝區的內容
    MSComm1.InBufferCount = 0
    '===============================
        '設置或返回接收緩衝區內用input 讀入位元組數,屬性為0表示input 讀取整個緩衝區的內容
    MSComm1.InputLen = 0
    '===========================333====
    Call SPIL
    If SPILFlag = 1 Then
      MSComm2.CommPort = 3
    Else
      MSComm2.CommPort = 2
    End If

    '===============================
        '"串列傳輸速率,同位元檢查方式,資料位元數,停止位元數"
    MSComm2.Settings = "9600,N,8,1"
    '===============================
        '打開或關閉序列埠
    If MSComm2.PortOpen = False Then
     MSComm2.PortOpen = True
     End If
    '===============================
        '返回接收緩衝區內等待取讀取的位元組數,屬性為0表示清空接收緩衝區的內容
    MSComm2.InBufferCount = 0
    '===============================
        '設置或返回接收緩衝區內用input 讀入位元組數,屬性為0表示input 讀取整個緩衝區的內容
    MSComm2.InputLen = 0

    Command1.Enabled = False
    Command6.Enabled = False
    
    Combo1.Clear
    Combo1.Text = "IC型號"
    Combo2.Text = "程式版別"
    Combo2.Enabled = False
    
    Set FD = FS.GetFolder(App.Path & "\PGM_ListDB\")
    
    If Not FS.FolderExists(App.Path & "\PGM_ListDB\Backup") Then
        FS.CreateFolder (App.Path & "\PGM_ListDB\Backup")
    End If
    
    MPTesterCounter = 0
    TesterCounter = 0
        
    For Each ff In FD.Files
        
        'Debug.Print ff.Name
            
        If InStr(ff.Name, "_") > 0 Then
            If Len(ff.Name) = 21 And Left(ff.Name, 8) = "MPTester" Then 'MPTester
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
        'connection to MPTester.mdb
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
        
        If Len(MPTesterRS.Fields(1)) = 15 Then
            CurrentName = Left(MPTesterRS.Fields(1), 7)
        Else
            CurrentName = Left(MPTesterRS.Fields(1), 6)
        End If
                
        If NameTmp <> CurrentName Then
            ComboExistFlag = False
            For VerifyCount = 0 To Combo1.ListCount
                If CurrentName = Combo1.List(VerifyCount) Then
                    ComboExistFlag = True
                    Exit For
                End If
            Next
                    
            If Not ComboExistFlag Then
                Combo1.AddItem CurrentName
            End If
            
            NameTmp = CurrentName
        
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
    Set TesterRS = ConnTesterDB.Execute("Select *  From [Tester] Where [Visible] = 1 Order By [PGM_Name]")

    
    TesterRS.MoveFirst
    
    Do Until TesterRS.EOF

        CurrentName = Left(TesterRS.Fields(1), 6)
                
        If NameTmp <> CurrentName Then
            ComboExistFlag = False
            For VerifyCount = 0 To Combo1.ListCount
                If CurrentName = Combo1.List(VerifyCount) Then
                    ComboExistFlag = True
                    Exit For
                End If
            Next
                    
            If Not ComboExistFlag Then
                Combo1.AddItem CurrentName
            End If
            
            NameTmp = CurrentName
        
        End If
                
        TesterRS.MoveNext
    Loop
    
    
    ConnMPTesterDB.Close
    Set ConnMPTesterDB = Nothing
    
    ConnTesterDB.Close
    Set ConnTesterDB = Nothing
    
    HostForm.Caption = HostForm.Caption & LastDateCode
    Label1.Caption = Label1.Caption & Mid(LastDateCode, 1, 4) & "/" & Mid(LastDateCode, 5, 2) & "/" & Mid(LastDateCode, 7, 2)
   
    If AllenDebug = 0 Then
        card = Register_Card(PCI_7248, 0) 'FOR PCI_7248
        Call SetTimer_1ms
        'SettingForm.Show 1 'FOR PCI_7248
        If card < 0 Then 'FOR PCI_7248
           MsgBox "Register Card Failed" 'FOR PCI_7248
        '   End 'FOR PCI_7248
        End If 'FOR PCI_7248
        Card_Initial 'FOR PCI_7248
 
    End If
    
    GPIBReady(0) = False
    GPIBReady(1) = False
    GetGPIBStatus(0) = False
    GetGPIBStatus(1) = False
    VB6_Flag = False
    
    Call Greatek  ' 20060329
    
    
    ' open I/O interface

    'a = OpenLinkDevice(0, hDevice) ' initial state
End Sub


Private Sub Command4_Click()

  k = DO_WritePort(card, Channel_P2CH, 0)
   Call MsecDelay(2)
   k = DO_WritePort(card, Channel_P2CH, 15)
End Sub


Sub PCI7248_bin(Channel As Byte, PCI7248bin As Byte, Check2Value As Byte)

Dim k As Integer
Dim DO_P As Long

        Print " CHANNEL=", Channel
        
        'Call Timer_1ms(5)
        If Check2Value = 0 Then
            DO_P = PCI7248bin
        Else
            DO_P = PCI7248bin - 128
        End If
        
            Print " send DO_P=", Hex(DO_P)
            k = DO_WritePort(card, Channel, DO_P)
            
        Call Timer_1ms(12)
        '========================================
        
        If Check2Value = 0 Then
           DO_P = PCI7248bin - PCI7248_EOT
        Else
           DO_P = PCI7248bin - PCI7248_EOT - 128
        End If
        
            Print " send DO_P=", Hex(DO_P)
            k = DO_WritePort(card, Channel, DO_P)
       
        Call Timer_1ms(7)
        '=======================================
        
        If Check2Value = 0 Then
            DO_P = PCI7248bin
        Else
            DO_P = PCI7248bin - 128
        End If
        
            Print " send DO_P=", Hex(DO_P)
            k = DO_WritePort(card, Channel, DO_P)
       
        Call Timer_1ms(7)
        '========================================
        
        If Check2Value = 0 Then
            DO_P = &HFF
        Else
            DO_P = &HFF - 128
        End If
        
            Print " send DO_P=", Hex(DO_P)
            k = DO_WritePort(card, Channel, DO_P)
        
        

'Const PCI7248_EOT = &H1 'for 7248 card
'Const PCI7248_PASS = &HFD 'for 7248 card  11111101
'Const PCI7248_BIN2 = &HFB 'for 7248 card  11111011
'Const PCI7248_BIN3 = &HF7 'for 7248 card  11110111
'Const PCI7248_BIN4 = &HEF 'for 7248 card  11101111
'Const PCI7248_BIN5 = &HDF 'for 7248 card  11011111

    
End Sub
Sub PrintReport()
If (ReportCheck.value = 0) Or (Check1.value = 1) Then
Exit Sub
End If
 
If AllenDebug = 1 Then
Bin1Site1 = 2903
Bin2Site1 = 10
Bin3Site1 = 6
Bin4Site1 = 1
Bin5Site1 = 1
Bin1Site2 = 2886
Bin2Site2 = 20
Bin3Site2 = 9
Bin4Site2 = 3
Bin5Site2 = 2
End If
'2. time control
 
  '  Dim EndDay As String
  '  Dim EndSecond As String
  '  Dim SNow As String
  '  Dim OutFileName As String
    EndSecond = Format(Now, "HH:MM:SS")
    EndDay = Format(Now, "YYYY/MM/DD")
    
    OutFileName = RunCardNO & "_" & ProcessID & "_" & Left(EndDay, 4) & Mid(EndDay, 6, 2) & Right(EndDay, 2)
    
    OutFileName = OutFileName & Left(EndSecond, 2) & Mid(EndSecond, 4, 2) & Right(EndSecond, 2) & ".txt"
    
    EndAt = EndDay & Space(1) & EndSecond
    
'3. Summary


Dim TestedSite1 As Long
Dim TestedSite2 As Long
Dim TestedTotal As Long
Dim TestedPercent As Single

Dim PassSite1 As Long
Dim PassSite2 As Long
Dim PassTotal As Long
Dim PassPercent As Single

Dim FailSite1 As Long
Dim FailSite2 As Long
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

' calculate summary





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
            
            
'=================================================================
Call UpdateDB
            
            
            
            
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
Print #1, "-------------------------------------------------------"
Print #1, Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Total  " & Space(3) & "Total"
Print #1, Space(13) & "COUNT  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percen"
Print #1, "-------------------------------------------------------"


'================ file output

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


'=============== file output

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

Close #1

'=================================================== printer section ===========================
    
'Printer.CurrentX = 300
Call GetProcessIDSub
'Call PrintReportSummary2
Exit Sub ' new code and keep the old individual data

If AllenDebug = 1 Then
     Exit Sub
End If
 Printer.FontSize = 14
 Printer.Font = "標楷體"
 Printer.Print "#####################################################"
 Printer.Print "Name of PC: " & NameofPC
 Printer.Print "Program Name: " & ProgramName
 Printer.Print "Program Rersion Code: " & ProgramRevisionCode
 Printer.Print "Device ID: " & DeviceID
 Printer.Print "Run Card NO: " & RunCardNO
 Printer.Print "Lot ID: " & LotID
 Printer.Print "Process: " & ProcessID
 Printer.Print "Start at: " & StartAt
 Printer.Print "End at: " & EndAt
 Printer.Print "HandelerID: " & HandlerID
 Printer.Print "Operator Name: " & OperatorName
 Printer.Print
 Printer.Print "-------------------------------------------------------"
 Printer.Print Space(13) & "Site 1 " & Space(3) & "Site 2 " & Space(3) & "Total  " & Space(3) & "Total"
 Printer.Print Space(13) & "COUNT  " & Space(3) & "Count  " & Space(3) & "Count  " & Space(3) & "Percen"
 Printer.Print "-------------------------------------------------------"



 
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


 Printer.Print "2 BIN2" & Space(7) & Space(7 - Len(Format(Bin2Site1, "#######"))) & Format(Bin2Site1, "#######") _
                       & Space(3) & Space(7 - Len(Format(Bin2Site2, "#######"))) & Format(Bin2Site2, "#######") _
                       & Space(3) & Space(7 - Len(Format(Bin2Total, "#######"))) & Format(Bin2Total, "#######") _
                       & Space(3) & Space(7 - Len(Format(Bin2Percent, "0.00%"))) & Format(Bin2Percent, "0.00%")


 Printer.Print "3 BIN3" & Space(7) & Space(7 - Len(Format(Bin3Site1, "#######"))) & Format(Bin3Site1, "#######") _
                       & Space(3) & Space(7 - Len(Format(Bin3Site2, "#######"))) & Format(Bin3Site2, "#######") _
                       & Space(3) & Space(7 - Len(Format(Bin3Total, "#######"))) & Format(Bin3Total, "#######") _
                       & Space(3) & Space(7 - Len(Format(Bin3Percent, "0.00%"))) & Format(Bin3Percent, "0.00%")


 Printer.Print "4 BIN4" & Space(7) & Space(7 - Len(Format(Bin4Site1, "#######"))) & Format(Bin4Site1, "#######") _
                       & Space(3) & Space(7 - Len(Format(Bin4Site2, "#######"))) & Format(Bin4Site2, "#######") _
                       & Space(3) & Space(7 - Len(Format(Bin4Total, "#######"))) & Format(Bin4Total, "#######") _
                       & Space(3) & Space(7 - Len(Format(Bin4Percent, "0.00%"))) & Format(Bin4Percent, "0.00%")


 Printer.Print "5 BIN5" & Space(7) & Space(7 - Len(Format(Bin5Site1, "#######"))) & Format(Bin5Site1, "#######") _
                       & Space(3) & Space(7 - Len(Format(Bin5Site2, "#######"))) & Format(Bin5Site2, "#######") _
                       & Space(3) & Space(7 - Len(Format(Bin5Total, "#######"))) & Format(Bin5Total, "#######") _
                       & Space(3) & Space(7 - Len(Format(Bin5Percent, "0.00%"))) & Format(Bin5Percent, "0.00%")

 Printer.EndDoc



    
End Sub
Private Sub Command2_Click()

    avgTestTime = 0
    totalTestTime = 0
    testTime = 0
    UPH = 0

'1. loop flag
On Error Resume Next

If (MsgBox("Stop Test ?", vbYesNo + vbQuestion + vbDefaultButton2, "Comform Stop") = vbNo) Then
  Exit Sub
End If

Call UnlockOption
flag = 0

Call PrintReport
    
End Sub

Private Sub Command3_Click()
 End
End Sub

Sub TimeDelay(Time As Long)
Dim i As Long
    i = 0
    For i = 0 To Time   'PwrLed.GPO 7,6,5,4,3
         
    Next
End Sub
Sub ReStartDelay1(WaitSec As Single)
Dim start As Single
Dim pause As Single
Dim i As Long
i = 0
 Label28.BackColor = RGB(255, 0, 0)
 Label29.BackColor = RGB(255, 0, 0)


Do

    start = Timer
    Do
        DoEvents
        pause = Timer
    Loop Until pause - start >= 1
    
    Label28 = WaitSec - i
    Label29 = WaitSec - i
    i = i + 1
Loop Until i > WaitSec
 Label28.BackColor = &HFFFF00
 Label29.BackColor = &HFFFF00
End Sub


Private Sub Card_Initial()
  Dim i As Integer, j As Integer
  Dim result As Integer
  
  For i = 0 To 1  'Initial status is Output for all channels
    result = DIO_PortConfig(card, i * 5 + Channel_P1A, OUTPUT_PORT)
    'Shape_a(i).FillColor = OUTPUT_COLOR
    'status_a(i) = OUTPUT_PORT
    'For j = 0 To 7
      'bit_a(i * 8 + j) = doa_1
    'Next j
    value_a(i) = &HFF
    result = DO_WritePort(card, i * 5 + Channel_P1A, value_a(i))
    '===================================================================
    result = DIO_PortConfig(card, i * 5 + Channel_P1B, OUTPUT_PORT)
    'Shape_b(i).FillColor = OUTPUT_COLOR
    'status_b(i) = OUTPUT_PORT
    'For j = 0 To 7
     ' bit_b(i * 8 + j) = dob_1
    'Next j
    value_b(i) = &HFF
    result = DO_WritePort(card, i * 5 + Channel_P1B, value_b(i))
    '===================================================================
    result = DIO_PortConfig(card, i * 5 + Channel_P1CH, OUTPUT_PORT)
    'Shape_cu(i).FillColor = OUTPUT_COLOR
    'status_cu(i) = OUTPUT_PORT
    'For j = 0 To 3
      'bit_cu(i * 4 + j) = doc_1
    'Next j
    value_cu(i) = &HF
    result = DO_WritePort(card, i * 5 + Channel_P1CH, value_cu(i))
    '===================================================================
    result = DIO_PortConfig(card, i * 5 + Channel_P1CL, OUTPUT_PORT)
    'Shape_cl(i).FillColor = OUTPUT_COLOR
    'status_cl(i) = OUTPUT_PORT
    'For j = 0 To 3
      'bit_cl(i * 4 + j) = doc_1
    'Next j
    value_cl(i) = &HF
    result = DO_WritePort(card, i * 5 + Channel_P1CL, value_cl(i))
  Next i
End Sub





Private Sub MSComm1_OnComm()

'Select Case MSComm1.CommEvent

'Case comEvReceive

'If MSComm1.InBufferCount > 4 Then

 ' P



'End Select


End Sub



Private Sub ShowBin_Click()
 BinForm.Show
End Sub

Private Sub AU6366S4TestON()


'*step1=>\\\\\\\\\\\\\\\\\\\\\ Get Start Signal From Handle
'*
'*wait Start Signal From Handle
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    WaitForStart = Timer   ' Get Vcc on from Chip
     
    If TestMode = 0 Then  'ON LINE MODE    'Check1.Value = 0 => TestMode = 0  '上線模式
    
          Print "wait Start"
            Do     'wait  (VCC PowerON) & (hander 5ms start) signal
                   DoEvents
                  WaitStartTime = Timer - WaitForStart
                   k = DO_ReadPort(card, Channel_P2CH, DI_P)
                   'Call MsecDelay(0.1)
        '   Loop Until DI_P = 14 Or DI_P = 13 Or DI_P = 12 Or WaitStartTime > WAIT_START_TIME_OUT  'Allen
           Loop Until DI_P = 14 Or DI_P = 13 Or DI_P = 12
           Label31.Caption = DI_P
            Print "DI_P=", DI_P
            TotalRealTestTime = Timer - OldTotalRealTestTime
            OldTotalRealTestTime = Timer
            OldRealTestTime = Timer
            
     Else                 'Check1.Value = 1 => TestMode = 1   '離線模式
     
            Call MsecDelay(0.8)
            WaitStartTime = 0.8
            DI_P = 14
            TotalRealTestTime = Timer - OldTotalRealTestTime
            OldTotalRealTestTime = Timer
            OldRealTestTime = Timer
            
            'k = DO_WritePort(card, Channel_P1A, 255)
            'k = DO_WritePort(card, Channel_P1B, 255)
            
     End If
     
        buf1 = MSComm1.Input
        Label26.Caption = "實際總測試時間(含 load / unload) :" & TotalRealTestTime & "s"
        WaitStartCounter = WaitStartCounter + 1
                  
    If WaitStartTime > WAIT_START_TIME_OUT Then
         WaitStartTimeOut = 1
         WaitStartTimeOutCounter = WaitStartTimeOutCounter + 1
    End If
    'If ChipName = "AU6366S4" Then ' ARCH add 9/16
        k = DO_WritePort(card, Channel_P1A, &H3F) ' select S1 send 0011,1111 => Set (IN1,IN0)=(0,0)  and Channel_P2A = 63
        Call MsecDelay(0.01)
        Print "Change to SD state"
        k = DO_ReadPort(card, Channel_P1CL, DI_S1)
        Print "DI_S1="; DI_S1
        k = DO_ReadPort(card, Channel_P1CH, DI_S2)
        Print "DI_S2="; DI_S2
    'End If
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'
'    SHOW Alarm
'
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

If Check5.value = 1 Then 'arch change 940529 連續fail 警告

    If site1 = 1 And continuefail1 >= AlarmLimit Then
    
        If continuefail1_bin2 >= 3 Then
            Call MsecDelay(3)
        End If
        
        
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        If continuefail1_bin2 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check Chip Contact and Tester Driver!"
            '  MsgBox "site1 countiue fail please check Chip Contact and Tester Driver!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 1
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 1
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin2 = 0
            continuefail1 = 0
        ElseIf continuefail1_bin3 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check  Flash & CF & SD CARD!"
            
            ' MsgBox "site1 countiue fail please check  Flash & CF & SD CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 2
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 2
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin3 = 0
            continuefail1 = 0
            
        ElseIf continuefail1_bin4 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check XD CARD!"
            
            
            'MsgBox "site1 countiue fail please check XD CARD!"
            
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 3
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
                
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 3
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin4 = 0
            continuefail1 = 0
        ElseIf continuefail1_bin5 >= AlarmLimit Then
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check MS CARD!"
            ' MsgBox "site1 countiue fail please check MS CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 4
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 4
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin5 = 0
            continuefail1 = 0
        Else
            Print "Site1 check continuefail start !"
        End If
    
    End If 'If Site1 = 1 And continuefail1 >= AlarmLimit Then
   '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    If site2 = 1 And continuefail2 >= AlarmLimit Then
    
    
        If continuefail2_bin2 >= 3 Then
            Call MsecDelay(3)
        End If
    
        If continuefail2_bin2 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check Chip Contact and Tester Driver!"
            '  MsgBox "site2 countiue fail please check Chip Contact and Tester Driver!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 5
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 5
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
            continuefail2_bin2 = 0
            continuefail2 = 0
        ElseIf continuefail2_bin3 >= AlarmLimit Then
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check  Flash & CF & SD CARD!"
            'MsgBox "site2 countiue fail please check  Flash & CF & SD CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 6
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 6
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail2_bin3 = 0
            continuefail2 = 0
        ElseIf continuefail2_bin4 >= AlarmLimit Then
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check XD CARD!"
            
            'MsgBox "site2 countiue fail please check XD CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 7
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
        
        
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 7
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
            continuefail2_bin4 = 0
            continuefail2 = 0
        ElseIf continuefail2_bin5 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check MS CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 8
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            
            Print "Alarm!!!"
            Do
            
                DoEvents
                
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            
            Loop While (1)
            
            Print "Alarm Clear"
        
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 8
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
            
            '  MsgBox "site2 countiue fail please check MS CARD!"
            continuefail2_bin5 = 0
            continuefail2 = 0
        Else
            Print "Site2  check continuefail start !!"
        End If
    
    End If 'If Site2 = 1 And continuefail2 >= AlarmLimit Then

Else
    Print "on standard test step!"
End If 'If Check5.Value = 1 Then  連續fail 警告

'*******************************************************
'*
'*   OPEN power
'*
'**********************************************************
    If DI_P < 12 And DI_P >= 15 Then   'Allen 20050607 , change DI_P > 15, to DI_P >= 15
        Print "no start"
        GoTo err
    Else
       
        Print "get start signal!"
    
        Call MsecDelay(CAPACTOR_CHARGE)
        Call MsecDelay(UNLOAD_DRIVER)
          '  If ChipName = "AU6366S4" Then
                k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                k = DO_ReadPort(card, Channel_P1CH, DI_S2)
          '  Else
          '      k = DO_WritePort(card, Channel_P1CL, &HC) ' send 1100 => SetSITE1 power" Channel_P1CL = 12
          '      k = DO_WritePort(card, Channel_P1CH, &HC) ' send 1100 => SetSITE2 power" Channel_P1CH = 12
          '  End If
        k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power" Channel_P2A = 127
        k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power" Channel_P2b = 127
        NewPowerOnTime = POWER_ON_TIME - 0.4
    
        If NewPowerOnTime > 0 Then
            Call MsecDelay(NewPowerOnTime)
        End If
        
    End If
   
   
   
'*STEP2=> wait tester send ready signal\\\\\\\\\\\\\\
'*
'*  Check Tester Ready Signal
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


   MSComm2.InBufferCount = 0
   MSComm1.InBufferCount = 0
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
            
            If site1 = 1 Then
                         If TesterReady1 = 0 Then
                         
                               buf1 = MSComm1.Input
                               
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
             
            If site2 = 1 Then
                         If TesterReady2 = 0 Then
                         
                               buf2 = MSComm2.Input
                               
                               TesterStatus2 = TesterStatus2 & buf2
                               
                               If (InStr(1, TesterStatus2, "Ready") <> 0) Then
                                         TesterReady2 = 1
                               End If
                         End If
                         
            Else
               TesterReady2 = 1
                  
            End If
                             
            '===================================
            ' Reset rountine : condsider Reset fail
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
                
                    '=========== Reset  Rountine
                    
                    If TesterReady1 = 0 And TesterDownCount1 = 0 And FirstRun = 1 Then ' reset tester1
                    '============== close module power
                    ResetCounter1 = ResetCounter1 + 1
                      '  If ChipName = "AU6366S4" Then
                            k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                            k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                      '  Else
                      '      k = DO_WritePort(card, Channel_P1CL, &HF) ' send 1111 => Set SITE1 power off" Channel_P1CL = 15
                      '      k = DO_WritePort(card, Channel_P1CH, &HF) ' send 1111 => Set SITE2 power off" Channel_P1CH = 15
                      '  End If
                    k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    '============= Reset PC
                    TesterDownCount1 = 1
                    k = DO_WritePort(card, Channel_P1B, &HF) ' send 0000,1111 => RESET PC " Channel_P1B= 15
                    Call MsecDelay(2)
                    k = DO_WritePort(card, Channel_P1B, &HFF) ' send 1111,1111 => RESET PC " Channel_P1B= 255
                    WaitForPowerOn1 = Timer
                    '============== clear comm buffer
                    MSComm1.InBufferCount = 0
                    TesterStatus1 = ""
                
                End If
            
            
                If TesterReady2 = 0 And TesterDownCount2 = 0 And FirstRun = 1 Then ' reset tester2
                    '============== close module power
                    ResetCounter2 = ResetCounter2 + 1
                      '  If ChipName = "AU6366S4" Then
                            k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                            k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                      '  Else
                      '      k = DO_WritePort(card, Channel_P1CL, &HF) ' send 1111 => Set SITE1 power off" Channel_P1CL = 15
                      '      k = DO_WritePort(card, Channel_P1CH, &HF) ' send 1111 => Set SITE2 power off" Channel_P1CH = 15
                      '  End If
                    k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    '============== Reset PC
                    TesterDownCount2 = 1
                    k = DO_WritePort(card, Channel_P1B, &HF) ' send 0000,1111 => RESET PC " Channel_P1B= 15
                    Call MsecDelay(2)
                    k = DO_WritePort(card, Channel_P1B, &HFF) ' send 1111,1111 => RESET PC " Channel_P1B= 255
                    WaitForPowerOn2 = Timer
                    '============== clear comm buffer
                    MSComm2.InBufferCount = 0
                    TesterStatus2 = ""
                
                End If
            
            End If 'If Timer - WaitForReady > 1 Then
            
                                        
            '===============================
            ' screen down count routine
            '==============================
            
             If TesterDownCount1 = 1 Then
             
                 TesterDownCountTimer1 = Timer - WaitForPowerOn1
                 Label28.Caption = CInt(TesterDownCountTimer1)
                 
                 If TesterReady1 = 1 Then
                     '====== open module power
                     '   If ChipName = "AU6366S4" Then
                            k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                            k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                     '   Else
                     '       k = DO_WritePort(card, Channel_P1CL, &HC) ' send 1100 => SetSITE1 power" Channel_P1CL = 12
                     '       k = DO_WritePort(card, Channel_P1CH, &HC) ' send 1100 => SetSITE2 power" Channel_P1CH = 12
                     '   End If
                     k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 255
                     k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 255
                     
                     Call MsecDelay(POWER_ON_TIME)
                     '=== clear flag
                     TesterDownCount1 = 0
                 End If
                 
                 If TesterDownCountTimer1 > 90 Then  'Reset fail
                     TesterDownCount1 = 0
                 End If
             
             
             End If 'If TesterDownCount1 = 1 Then
             
             '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
             
             If TesterDownCount2 = 1 Then
             
                 TesterDownCountTimer2 = Timer - WaitForPowerOn2
                 Label29.Caption = CInt(TesterDownCountTimer2)
                 
                 If TesterReady2 = 1 Then
                     '====== open module power
                     '   If ChipName = "AU6366S4" Then
                            k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                            k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                     '   Else
                     '       k = DO_WritePort(card, Channel_P1CL, &HC) ' send 1100 => SetSITE1 power" Channel_P1CL = 12
                     '       k = DO_WritePort(card, Channel_P1CH, &HC) ' send 1100 => SetSITE2 power" Channel_P1CH = 12
                     '   End If
                     k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 255
                     k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 255
                     Call MsecDelay(POWER_ON_TIME)
                     '=== clear flag
                     TesterDownCount2 = 0
                 End If
                 
                 If TesterDownCountTimer2 > 90 Then    ' Reset fail
                     TesterDownCount2 = 0
                 End If
                 
             End If 'If TesterDownCount2 = 1 Then
              '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%5%%%%%
        End If
                 
    Loop Until (TesterReady1 = 1) And (TesterReady2 = 1)
    
        FirstRun = 1
         
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'*
'*    Testing Loop
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
If (DI_P >= 12) And (DI_P < 15) Then
  
    
        ' init falg
         GetStart = 1
    Label3.BackColor = RGB(255, 255, 255)
    Label3 = ""
    Label16.BackColor = RGB(255, 255, 255)
    Label16 = ""
         
         
             
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '
    '                Site1 and Site2  begin
    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
             
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '        Testing LED function    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Print "==========================="
          

        ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '
        '  Allen 0526 begin 1 : for no card test,pull high Card detect
        '
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
           If NoCardTest.value = 1 Then
                
                   'pull High
                  '  If ChipName = "AU6366S4" Then
                        k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                        k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                  '  Else
                  '      k = DO_WritePort(card, Channel_P1CL, &HD) ' send 1101 => "SetSITE1 Power ON" & " SetSITE1 CDN High"Channel_P1CL = 13
                  '      k = DO_WritePort(card, Channel_P1CH, &HD) ' send 1101 => "SetSITE2 Power ON" & " SetSITE2 CDN High"Channel_P1CH = 13
                  '  End If
                 k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power on"& " SetSITE2 CDN High" Channel_P2A = 127
                 k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power on"& " SetSITE1 CDN High" Channel_P2A = 127
           Else
                 '   If ChipName = "AU6366S4" Then
                        k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                        k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                 '   Else
                 '       k = DO_WritePort(card, Channel_P1CL, &HC) ' send 1100 => "SetSITE1 Power ON" & " SetSITE1 CDN Low"Channel_P1CL = 12
                 '       k = DO_WritePort(card, Channel_P1CH, &HC) ' send 1100 => "SetSITE2 Power ON" & " SetSITE2 CDN Low"Channel_P1CH = 12
                 '   End If
                 k = DO_WritePort(card, Channel_P2A, &H3F) ' send 0011,1111 => Set power on"& " SetSITE2 CDN Low" Channel_P2A = 63
                 k = DO_WritePort(card, Channel_P2B, &H3F) ' send 0011,1111 => Set power on"& " SetSITE1 CDN Low" Channel_P2A = 63
           End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '
        '  Allen 0526 End  1 : for no card test,pull high Card detect
        '
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
   
    '*STEP4=> Waitting for Response from  Tester\\\\\\\\\\\\\\\\\\\\\
    '*
    '*    Wait Test Result from each Tester
    '*
    '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
   '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      '
      '  Allen 0601 Remark : no card on board test card detect and card change signal
      '
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
        NoCardTestResult1 = ""
        NoCardTestResult2 = ""
        
        If site1 = 1 And NoCardTest.value = 1 Then
    
            MSComm1.Output = ChipName   ' trans strat test signal to TEST PC
            MSComm1.InBufferCount = 0
            MSComm1.InputLen = 0
            NoCardWaitForTest1 = Timer ' wait for timer  and test result
        End If
        
        If site2 = 1 And NoCardTest.value = 1 Then
        
            MSComm2.Output = ChipName   ' trans strat test signal to TEST PC
            MSComm2.InBufferCount = 0
            MSComm2.InputLen = 0
            NoCardWaitForTest2 = Timer
        End If
    
    
        Print "send begin test signal to test"
        TesterStatus1 = ""
        TesterStatus2 = ""
        
      
         Do
            DoEvents
            '========================
            If site1 = 1 And NoCardTest.value = 1 Then
                If NoCardTestStop1 = 0 Then
                
                    If MSComm1.InBufferCount >= 4 Then
                        NoCardTestResult1 = MSComm1.Input
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
            '========================
            
            If site2 = 1 And NoCardTest.value = 1 Then
                If NoCardTestStop2 = 0 Then
                     If MSComm2.InBufferCount >= 4 Then
                            NoCardTestResult2 = MSComm2.Input
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
          '========================
          
        Loop Until (NoCardTestStop1 = 1) And (NoCardTestStop2 = 1)
    
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      '
      '  Allen 0526 Remark : no card on board test card detect and card change signal
      '
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
    '*STEP3=>Send command to PC teser\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '*
    '*    Send ChipName to PC teser
    '*
    '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    TestResult1 = ""
    TestResult2 = ""
    
    If NoCardTest.value = 1 Then
              '  If ChipName = "AU6366S4" Then
                    k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                    k = DO_ReadPort(card, Channel_P1CH, DI_S2)
              '  Else
              '      k = DO_WritePort(card, Channel_P1CL, &HC) ' send 1100 => "SetSITE1 Power ON" & " SetSITE1 CDN Low"Channel_P1CL = 12
              '      k = DO_WritePort(card, Channel_P1CH, &HC) ' send 1100 => "SetSITE2 Power ON" & " SetSITE2 CDN Low"Channel_P1CH = 12
              '  End If
              k = DO_WritePort(card, Channel_P2A, &H3F) ' send 0011,1111 => Set power on"& " SetSITE2 CDN Low" Channel_P2A = 63
              k = DO_WritePort(card, Channel_P2B, &H3F) ' send 0011,1111 => Set power on"& " SetSITE1 CDN Low" Channel_P2A = 63
'             Call MsecDelay(0.1)
            If site1 = 1 Then  '****** Continue condition lock at PC tester
            
            MSComm1.Output = NoCardTestResult1   ' only pass can continue at PC Tester
            MSComm1.InBufferCount = 0
            MSComm1.InputLen = 0
            WaitForTest1 = Timer ' wait for timer  and test result
             
              If NoCardTestResult1 <> "PASS" Then
                 TestResult1 = NoCardTestResult1
              End If
            
            End If
          
            
            If site2 = 1 Then
            
                MSComm2.Output = NoCardTestResult2   ' only pass can continue at PC Tester
                MSComm2.InBufferCount = 0
                MSComm2.InputLen = 0
                WaitForTest2 = Timer
            
                If NoCardTestResult2 <> "PASS" Then
                    TestResult2 = NoCardTestResult2
                End If
            End If
    
    
    
    Else
          '  If ChipName = "AU6366S4" Then
                k = DO_WritePort(card, Channel_P1A, &H3F) ' select S1 send 0011,1111 => Set (IN1,IN0)=(0,0)  and Channel_P2A = 63
                Call MsecDelay(0.1)
                Print "Change to SD state"
                k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                Print "DI_S1="; DI_S1
                k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                Print "DI_S2="; DI_S2
                 
           '  End If
    
    
    
            If site1 = 1 And DI_S1 = 14 Then
            ' If site1 = 1 Then
                MSComm1.Output = ChipName   ' trans strat test signal to TEST PC
                MSComm1.InBufferCount = 0
                MSComm1.InputLen = 0
                WaitForTest1 = Timer ' wait for timer  and test result
             Else
                TestResult1 = "CARD initial FAIL"
                TestStop1 = 1
             End If
            
            
           If site2 = 1 And DI_S2 = 14 Then
             'If site2 = 1 Then
                MSComm2.Output = ChipName   ' trans strat test signal to TEST PC
                MSComm2.InBufferCount = 0
                MSComm2.InputLen = 0
                WaitForTest2 = Timer
             Else
                TestResult2 = "CARD initial FAIL"
                TestStop2 = 1
             End If
    End If
'===================================================
  
DoEvents
If ChipName = "AU6366S4" Or ChipName = "AU66S4_F" Then '特定為 AU6366S4 使用
'ARCH try add AU6366s4 point


    Do
            DoEvents 'wait test slot0 result "SD"
     
                If site1 = 1 Then
                    If TestStop1 = 0 Then
                    
                        If MSComm1.InBufferCount >= 4 Then
                            TestResult1 = MSComm1.Input
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
                
                   '========================
                   
                If site2 = 1 Then
                    If TestStop2 = 0 Then
                    
                        If MSComm2.InBufferCount >= 4 Then
                            TestResult2 = MSComm2.Input
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
    Loop Until (TestStop1 = 1) And (TestStop2 = 1)
    '==================='wait test slot0 "SD" result end==========================
               Label3 = TestResult1
               Label16 = TestResult2
                If TestResult1 = "SD_PASS" Or TestResult2 = "SD_PASS" Then
                    k = DO_WritePort(card, Channel_P1A, &H7F) ' select S2 send 0111,1111 => Set (IN1,IN0)=(0,1)  and Channel_P2A = 127
                    Call MsecDelay(0.01)
                    Print "Change to CF state"
                    k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                    Print "DI_S1="; DI_S1
                    k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                    Print "DI_S2="; DI_S2
                    
                        k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                        k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off" Channel_P2b = 255
                       Call MsecDelay(0.1)
                        k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power" Channel_P2A = 127
                        k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power" Channel_P2b = 127
                       Call MsecDelay(0.8)
                End If
                    
                If site1 = 1 And TestResult1 = "SD_PASS" And DI_S1 = 13 Then
                'If site1 = 1 And TestResult1 = "SD_PASS" Then
                    OldTestResult1 = TestResult1
                    MSComm1.Output = TestResult1   ' trans strat test signal to TEST PC
                    MSComm1.InBufferCount = 0
                    MSComm1.InputLen = 0
                    WaitForTest1 = Timer ' wait for timer  and test result
                    TestStop1 = 0 'reset 初始條件設定
                    TestResult1 = ""
                End If
                
                If site2 = 1 And TestResult2 = "SD_PASS" And DI_S2 = 13 Then
                ' If site2 = 1 And TestResult2 = "SD_PASS" Then
                    OldTestResult2 = TestResult2
                    MSComm2.Output = TestResult2   ' trans strat test signal to TEST PC
                    MSComm2.InBufferCount = 0
                    MSComm2.InputLen = 0
                    WaitForTest2 = Timer
                    TestStop2 = 0 'reset 初始條件設定
                    TestResult2 = ""
                End If
        
      Do   '===============wait test slot1 "CF" result
            DoEvents
     
                If site1 = 1 And OldTestResult1 = "SD_PASS" Then
                  
                    If TestStop1 = 0 Then
                       
                        If MSComm1.InBufferCount >= 4 Then
                            TestResult1 = MSComm1.Input
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
                
                   '========================
                   
                If site2 = 1 And OldTestResult2 = "SD_PASS" Then
                    If TestStop2 = 0 Then
                    
                        TestResult2 = ""
                        If MSComm2.InBufferCount >= 4 Then
                            TestResult2 = MSComm2.Input
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
           
    Loop Until (TestStop1 = 1) And (TestStop2 = 1)
     '======================='wait test slot1 "CF" result end=======================
                Label3 = TestResult1
                Label16 = TestResult2
                If TestResult1 = "CF_PASS" Or TestResult2 = "CF_PASS" Then
                    k = DO_WritePort(card, Channel_P1A, &HBF) ' select S3 send 1011,1111 => Set (IN1,IN0)=(1,0)  and Channel_P2A = 191
                    Call MsecDelay(0.05)
                    Print "Change to XD state"
                    k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                    Print "DI_S1="; DI_S1
                    k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                    Print "DI_S2="; DI_S2
                    
                        k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                        k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off" Channel_P2b = 255
                        Call MsecDelay(0.1)
                        k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power" Channel_P2A = 127
                        k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power" Channel_P2b = 127
                        Call MsecDelay(0.6)
                        Call MsecDelay(0.2)
                End If
                    
                If site1 = 1 And TestResult1 = "CF_PASS" And DI_S1 = 11 Then
                'If site1 = 1 And TestResult1 = "CF_PASS" Then
                    OldTestResult1 = TestResult1
                    MSComm1.Output = TestResult1   ' trans strat test signal to TEST PC
                    MSComm1.InBufferCount = 0
                    MSComm1.InputLen = 0
                    WaitForTest1 = Timer ' wait for timer  and test result
                    TestStop1 = 0 'reset 初始條件設定
                End If
                
                
                If site2 = 1 And TestResult2 = "CF_PASS" And DI_S2 = 11 Then
                'If site2 = 1 And TestResult2 = "CF_PASS" Then
                    OldTestResult2 = TestResult2
                    MSComm2.Output = TestResult2   ' trans strat test signal to TEST PC
                    MSComm2.InBufferCount = 0
                    MSComm2.InputLen = 0
                    WaitForTest2 = Timer
                    TestStop2 = 0 'reset  初始條件設定
                End If
                
       Do       '=============wait test slot2 "XD" result
            DoEvents
                If site1 = 1 And OldTestResult1 = "CF_PASS" Then
                    If TestStop1 = 0 Then
                        TestResult1 = ""
                        If MSComm1.InBufferCount >= 4 Then
                            TestResult1 = MSComm1.Input
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
                
                   '========================
                   
                If site2 = 1 And OldTestResult2 = "CF_PASS" Then
                    If TestStop2 = 0 Then
                        TestResult2 = ""
                        If MSComm2.InBufferCount >= 4 Then
                            TestResult2 = MSComm2.Input
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
     
    Loop Until (TestStop1 = 1) And (TestStop2 = 1)
    
     '====================='wait test slot2 "XD" result end ==============================
               Label3 = TestResult1
               Label16 = TestResult2
                If TestResult1 = "XD_PASS" Or TestResult2 = "XD_PASS" Then
                    k = DO_WritePort(card, Channel_P1A, &HFF) ' select S3 send 1111,1111 => Set (IN1,IN0)=(1,1)  and Channel_P2A = 255
                    Call MsecDelay(0.05)
                    Print "Change to MS state"
                    k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                    Print "DI_S1="; DI_S1
                    k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                    Print "DI_S2="; DI_S2
                        k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                        k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off" Channel_P2b = 255
                       Call MsecDelay(0.1)
                        k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power" Channel_P2A = 127
                        k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power" Channel_P2b = 127
                       Call MsecDelay(0.6)
                       Call MsecDelay(0.2)
                End If
                    
                    
                If site1 = 1 And TestResult1 = "XD_PASS" And DI_S1 = 7 Then
                'If site1 = 1 And TestResult1 = "XD_PASS" Then
                    OldTestResult1 = TestResult1
                    MSComm1.Output = TestResult1   ' trans strat test signal to TEST PC
                    MSComm1.InBufferCount = 0
                    MSComm1.InputLen = 0
                    WaitForTest1 = Timer ' wait for timer  and test result
                    TestStop1 = 0 'reset 初始條件設定
                End If
                
                
                If site2 = 1 And TestResult2 = "XD_PASS" And DI_S2 = 7 Then
                'If site2 = 1 And TestResult2 = "XD_PASS" Then
                    OldTestResult2 = TestResult2
                    MSComm2.Output = TestResult2   ' trans strat test signal to TEST PC
                    MSComm2.InBufferCount = 0
                    MSComm2.InputLen = 0
                    WaitForTest2 = Timer
                    TestStop2 = 0 'reset 初始條件設定
                End If
        
      Do           '=====================wait test slot3 "MS" result
            DoEvents
                If site1 = 1 And OldTestResult1 = "XD_PASS" Then
                    If TestStop1 = 0 Then
                        TestResult1 = ""
                        If MSComm1.InBufferCount >= 4 Then
                            TestResult1 = MSComm1.Input
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
                
                   '========================
                If site2 = 1 And OldTestResult2 = "XD_PASS" Then
                    If TestStop2 = 0 Then
                        TestResult2 = ""
                        If MSComm2.InBufferCount >= 4 Then
                            TestResult2 = MSComm2.Input
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
         
    Loop Until (TestStop1 = 1) And (TestStop2 = 1)
    
     '======================'wait test slot3 "MS"  result end==================================
Else '一般多槽IC 正常測試模式 ChipName <> "au6366S4"
    DoEvents
    Do
        If site1 = 1 And NoCardTestResult1 = "PASS" Then
            If TestStop1 = 0 Then
            
                If MSComm1.InBufferCount >= 4 Then
                    TestResult1 = MSComm1.Input
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
        '========================
        If site2 = 1 And NoCardTestResult2 = "PASS" Then
            If TestStop2 = 0 Then
                
                If MSComm2.InBufferCount >= 4 Then
                    TestResult2 = MSComm2.Input
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
    
    Loop Until (TestStop1 = 1) And (TestStop2 = 1)

End If
        
    '\\\\\\\\\\\\\\\\\\\\\\\\\\}wait Tester response END
                           
         If site1 = 1 Then
             Print "TestResult1= "; TestResult1
            
         End If
         
         If site2 = 1 Then
             Print "TestResult2= "; TestResult2
                            
         End If
        
        
         TestCounter = TestCounter + 1 ' Allen Debug
     
         If TestCycleTime1 > WAIT_TEST_CYCLE_OUT And site1 = 1 Then
         
           WaitTestTimeOut1 = 1
           WaitTestTimeOutCounter1 = WaitTestTimeOutCounter1 + 1
         End If
                  
         If TestCycleTime2 > WAIT_TEST_CYCLE_OUT And site2 = 1 Then
         
           WaitTestTimeOut2 = 1
           WaitTestTimeOutCounter2 = WaitTestTimeOutCounter2 + 1
         End If
      
              
    '/////////////////////////////////////////////////////////////////////////
    '
    '   RT Condition
    '
    '//////////////////////////////////////////////////////////////////////////
              
              
     If Check3.value = 1 Then   ' 不RT=> not low yield sorting
         GoTo err
     End If
     
     
     If site1 = 1 And site2 = 0 Then
         If TestResult1 = "PASS" Then
             GoTo err
         End If
         
     End If
     
     
     If site1 = 0 And site2 = 1 Then
         If TestResult2 = "PASS" Then
             GoTo err
         End If
         
     End If
     
     If site1 = 1 And site2 = 1 Then
         If TestResult1 = "PASS" And TestResult2 = "PASS" Then
             GoTo err
         End If
         
     End If
        '////////////////////////////// initial condition
           
         Print "RT begin"
                 
         '1.close power
         '2.delay 10 s
         '3.send power
         '4.RT core
         
         Print "close power"
          
           ' k = DO_WritePort(card, Channel_P1A, &HFF)
           ' k = DO_WritePort(card, Channel_P1B, &HFF)
           
           ' If ChipName = "AU6366S4" Then
                k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                k = DO_ReadPort(card, Channel_P1CH, DI_S2)
           ' Else
           '     k = DO_WritePort(card, Channel_P1CL, &HF) ' send 1111 => "SetSITE1 Power Off" & " SetSITE1 CDN high"Channel_P1CL = 15
           '     k = DO_WritePort(card, Channel_P1CH, &HF) ' send 1111 => "SetSITE2 Power Off" & " SetSITE2 CDN high"Channel_P1CH = 15
           ' End If
            k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off"& " SetSITE2 CDN high" Channel_P2A = 255
            k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off"& " SetSITE1 CDN high" Channel_P2A = 255
         Call MsecDelay(RT_INTERVAL)  ' to let system unload driver
         
         Print "Send power"
         'k = DO_WritePort(card, Channel_P1A, &H7F)
         'k = DO_WritePort(card, Channel_P1B, &H7F)
         
            'If ChipName = "AU6366S4" Then
                k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                k = DO_ReadPort(card, Channel_P1CH, DI_S2)
           ' Else
           '     k = DO_WritePort(card, Channel_P1CL, &HC) ' send 1100 => "SetSITE1 Power ON" & " SetSITE1 CDN Low"Channel_P1CL = 13
           '     k = DO_WritePort(card, Channel_P1CH, &HC) ' send 1100 => "SetSITE2 Power ON" & " SetSITE2 CDN Low"Channel_P1CH = 13
           ' End If
            k = DO_WritePort(card, Channel_P2A, &H3F) ' send 0011,1111 => Set power on"& " SetSITE2 CDN Low" Channel_P2A = 63
            k = DO_WritePort(card, Channel_P2B, &H3F) ' send 0011,1111 => Set power on"& " SetSITE1 CDN Low" Channel_P2A = 63
         Call MsecDelay(POWER_ON_TIME)
         
                 
                 
                 
        If site1 = 1 And TestResult1 <> "PASS" Then
            
            MSComm1.Output = ChipName ' trans strat test signal to TEST PC
            MSComm1.InBufferCount = 0
            MSComm1.InputLen = 0
            RTWaitForTest1 = Timer ' wait for timer  and test result
         End If
            
         If site2 = 1 And TestResult2 <> "PASS" Then
             
            MSComm2.Output = ChipName    ' trans strat test signal to TEST PC
            MSComm2.InBufferCount = 0
            MSComm2.InputLen = 0
            RTWaitForTest2 = Timer
         End If
         
        
         Print "send begin test signal to test"
         
         '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Wait for Response from PC Tester
                
            Do
                DoEvents
                
                If site1 = 1 And TestResult1 <> "PASS" Then
                    If RTTestStop1 = 0 Then
                        RTTestResult1 = MSComm1.Input
                        
                        RTTestCycleTime1 = Timer - RTWaitForTest1
                        
                        If (RTTestResult1 <> "" Or RTTestCycleTime1 > WAIT_TEST_CYCLE_OUT) Then
                        
                            RTTestCounter1 = RTTestCounter1 + 1 ' Allen Debug
                            RTTestStop1 = 1
                        End If
                    End If
                
                Else
                    RTTestStop1 = 1
                
                End If
                '========================
                
                If site2 = 1 And TestResult2 <> "PASS" Then
                    If RTTestStop2 = 0 Then
                        RTTestResult2 = MSComm2.Input
                        
                        RTTestCycleTime2 = Timer - RTWaitForTest2
                        
                        If (RTTestResult2 <> "" Or RTTestCycleTime2 > WAIT_TEST_CYCLE_OUT) Then
                            RTTestStop2 = 1
                            RTTestCounter2 = RTTestCounter2 + 1
                        End If
                    End If
                
                Else
                    RTTestStop2 = 1
                
                
                End If
            
            
            Loop Until (RTTestStop1 = 1) And (RTTestStop2 = 1)
                           
                If site1 = 1 Then
                    Print "RTTestResult1= "; RTTestResult1
                End If
                
                If site2 = 1 Then
                    Print "RTTestResult2= "; RTTestResult2
                End If
                
               
                
                If RTTestCycleTime1 > WAIT_TEST_CYCLE_OUT And site1 = 1 Then
                
                    RTWaitTestTimeOut1 = 1
                    RTWaitTestTimeOutCounter1 = RTWaitTestTimeOutCounter1 + 1
                End If
                
                If RTTestCycleTime2 > WAIT_TEST_CYCLE_OUT And site2 = 1 Then
                
                    RTWaitTestTimeOut2 = 1
                    RTWaitTestTimeOutCounter2 = RTWaitTestTimeOutCounter2 + 1
                End If
                 
                 
                 
                     
                 
End If  '////////////////////// Test end
err:              '  Testing Loop end
End Sub
Private Sub CommomdTestON()

Dim HUBEnaOn As Byte
Dim GPIBInquiryTime


If Left(ChipName, 10) = "AU6254XLS4" Then
FirstRun = 1
End If

If LoopTest1_Flag = True Or LoopTest2_Flag = True Then
    GoTo LoopTest_Start
End If


'*step1=>\\\\\\\\\\\\\\\\\\\\\ Get Start Signal From Handle
'*
'*wait Start Signal From Handle
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    WaitForStart = Timer   ' Get Vcc on from Chip
     
    If TestMode = 0 Then  'ON LINE MODE    'Check1.Value = 0 => TestMode = 0  '上線模式
    
          Print "wait Start"
            Do     'wait  (VCC PowerON) & (hander 5ms start) signal
                   DoEvents
                  WaitStartTime = Timer - WaitForStart
                   k = DO_ReadPort(card, Channel_P2CH, DI_P)
                   'Call MsecDelay(0.1)
        '   Loop Until DI_P = 14 Or DI_P = 13 Or DI_P = 12 Or WaitStartTime > WAIT_START_TIME_OUT  'Allen
           Loop Until DI_P = 14 Or DI_P = 13 Or DI_P = 12
           Label31.Caption = DI_P
            Print "DI_P=", DI_P
            TotalRealTestTime = Timer - OldTotalRealTestTime
            OldTotalRealTestTime = Timer
            OldRealTestTime = Timer
            
     Else                 'Check1.Value = 1 => TestMode = 1   '離線模式
     
            Call MsecDelay(0.2)
            WaitStartTime = 0.2
            DI_P = 14
            TotalRealTestTime = Timer - OldTotalRealTestTime
            OldTotalRealTestTime = Timer
            OldRealTestTime = Timer
            
            'k = DO_WritePort(card, Channel_P1A, 255)
            'k = DO_WritePort(card, Channel_P1B, 255)
            
     End If
     
        buf1 = MSComm1.Input
        Label26.Caption = "實際總測試時間(含 load / unload) :" & TotalRealTestTime & "s"
        WaitStartCounter = WaitStartCounter + 1
                  
    If WaitStartTime > WAIT_START_TIME_OUT Then
         WaitStartTimeOut = 1
         WaitStartTimeOutCounter = WaitStartTimeOutCounter + 1
    End If
   
' ======================= continue fail
 If Check7.value = 1 Then
 
      If continuefail1 >= AlarmLimit Or continuefail2 >= AlarmLimit Then
       
          continuefail1_bin2 = 0
          continuefail1_bin3 = 0
          continuefail1_bin4 = 0
          continuefail1_bin5 = 0
       
          continuefail2_bin2 = 0
          continuefail2_bin3 = 0
          continuefail2_bin4 = 0
          continuefail2_bin5 = 0
          
          continuefail1 = 0
          
          continuefail2 = 0
          
          
          If Left(ChipName, 10) = "AU6254XLS4" Then
          k = DO_WritePort(card, Channel_P1CH, &H0)
          
          Call MsecDelay(6)
          
           k = DO_WritePort(card, Channel_P1CH, &HF)
          Call MsecDelay(2)
          
          End If
         
          
          
             k = DO_WritePort(card, Channel_P1B, &HF) ' send 0000,1111 => RESET PC " Channel_P1B= 15
                    Call MsecDelay(2)
                    k = DO_WritePort(card, Channel_P1B, &HFF) ' send 1111,1111 => RESET PC " Channel_P1B= 255
        
       '   MsgBox "BIn2 fail"

    End If
    
    
    End If
    
    
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'
'    SHOW Alarm
'
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

If Check5.value = 1 Then 'arch change 940529 連續fail 警告

    If site1 = 1 And continuefail1 >= AlarmLimit Then
    
        If continuefail1_bin2 >= 3 Then
            Call MsecDelay(3)
        End If
        
        
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        If continuefail1_bin2 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check Chip Contact and Tester Driver!"
            '  MsgBox "site1 countiue fail please check Chip Contact and Tester Driver!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 1
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 1
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin2 = 0
            continuefail1 = 0
        ElseIf continuefail1_bin3 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check  Flash & CF & SD CARD!"
            
            ' MsgBox "site1 countiue fail please check  Flash & CF & SD CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 2
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 2
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin3 = 0
            continuefail1 = 0
            
        ElseIf continuefail1_bin4 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check XD CARD!"
            
            
            'MsgBox "site1 countiue fail please check XD CARD!"
            
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 3
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
                
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 3
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin4 = 0
            continuefail1 = 0
        ElseIf continuefail1_bin5 >= AlarmLimit Then
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check MS CARD!"
            ' MsgBox "site1 countiue fail please check MS CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 4
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 4
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin5 = 0
            continuefail1 = 0
        Else
            Print "Site1 check continuefail start !"
        End If
    
    End If 'If Site1 = 1 And continuefail1 >= AlarmLimit Then
   '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    If site2 = 1 And continuefail2 >= AlarmLimit Then
    
    
        If continuefail2_bin2 >= 3 Then
            Call MsecDelay(3)
        End If
    
        If continuefail2_bin2 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check Chip Contact and Tester Driver!"
            '  MsgBox "site2 countiue fail please check Chip Contact and Tester Driver!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 5
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 5
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
            continuefail2_bin2 = 0
            continuefail2 = 0
        ElseIf continuefail2_bin3 >= AlarmLimit Then
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check  Flash & CF & SD CARD!"
            'MsgBox "site2 countiue fail please check  Flash & CF & SD CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 6
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 6
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail2_bin3 = 0
            continuefail2 = 0
        ElseIf continuefail2_bin4 >= AlarmLimit Then
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check XD CARD!"
            
            'MsgBox "site2 countiue fail please check XD CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 7
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
        
        
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 7
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
            continuefail2_bin4 = 0
            continuefail2 = 0
        ElseIf continuefail2_bin5 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check MS CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 8
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            
            Print "Alarm!!!"
            Do
            
                DoEvents
                
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            
            Loop While (1)
            
            Print "Alarm Clear"
        
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 8
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
            
            '  MsgBox "site2 countiue fail please check MS CARD!"
            continuefail2_bin5 = 0
            continuefail2 = 0
        Else
            Print "Site2  check continuefail start !!"
        End If
    
    End If 'If Site2 = 1 And continuefail2 >= AlarmLimit Then

Else
    Print "on standard test step!"
End If 'If Check5.Value = 1 Then  連續fail 警告

'*******************************************************
'*
'*   OPEN power
'*
'**********************************************************
    If DI_P < 12 And DI_P >= 15 Then   'Allen 20050607 , change DI_P > 15, to DI_P >= 15
        Print "no start"
        GoTo err
    Else
          
        Print "get start signal!"
    
        Call MsecDelay(CAPACTOR_CHARGE)
        Call MsecDelay(UNLOAD_DRIVER)
        
      '  k = DO_WritePort(card, Channel_P1CL, &HC) ' send 1100 => SetSITE1 power" Channel_P1CL = 12
       ' k = DO_WritePort(card, Channel_P1CH, &HC) ' send 1100 => SetSITE2 power" Channel_P1CH = 12
        
        k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power" Channel_P2A = 127
        k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power" Channel_P2b = 127
        NewPowerOnTime = POWER_ON_TIME - 0.4
    
        If NewPowerOnTime > 0 Then
            Call MsecDelay(NewPowerOnTime)
        End If
        
    End If
   
   
   
'*STEP2=> wait tester send ready signal\\\\\\\\\\\\\\
'*
'*  Check Tester Ready Signal
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\



LoopTest_Start:

   MSComm2.InBufferCount = 0
   MSComm1.InBufferCount = 0
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
            
            If site1 = 1 Then
                If TesterReady1 = 0 Then
                                
                    If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                        MSComm1.Output = "~"
                                
                        If HUBEnaOn = 0 Then
                            k = DO_WritePort(card, Channel_P1CH, &H0) ' set HUB module Ena on
                            HUBEnaOn = 1
                        End If
                                
                    End If
                                
                    buf1 = MSComm1.Input
                    TesterStatus1 = TesterStatus1 & buf1
                               
                    If (InStr(1, TesterStatus1, "Ready") <> 0) Then
                        TesterReady1 = 1
                    End If
                
                    If (GetGPIBStatus(0) = False) And (VB6_Flag = True) And (Mid(ChipName, 12, 1) <> "U") Then
                        MSComm1.Output = "AUGPIBGPIBGPIB"   ' trans strat test signal to TEST PC
                        MSComm1.InBufferCount = 0
                        MSComm1.InputLen = 0
                        GPIBInquiryTime = Timer
                        'Call MsecDelay(0.02)
                        
                        Do
                            buf1 = MSComm1.Input
                            TesterStatus1 = TesterStatus1 & buf1
                            Call MsecDelay(0.05)
                        Loop Until (InStr(1, TesterStatus1, "GPIBReady") <> 0) _
                                    Or (InStr(1, TesterStatus1, "GPIBUNReady") <> 0) _
                                    Or (Timer - GPIBInquiryTime > 1)
                                                        
                        If (Timer - GPIBInquiryTime > 1) Then
                            GoTo LoopTest_Start
                        End If
                                                        
                        Call MsecDelay(0.05)
                        MSComm1.Output = "AUGPIBACK"
                        GetGPIBStatus(0) = True
                                            
                        If InStr(1, TesterStatus1, "GPIBReady") <> 0 Then
                            GPIBReady(0) = True
                            Site1GPIB_Label.BackColor = &H80FFFF
                            Site1GPIB_Label.ForeColor = &HFF0000
                            Site1GPIB_Label.Caption = " GPIB Rdy"
                        Else
                            GPIBReady(0) = False
                            Site1GPIB_Label.BackColor = &H80FFFF
                            Site1GPIB_Label.ForeColor = &H80FF&
                            Site1GPIB_Label.Caption = " No GPIB"
                        End If
                        Call MsecDelay(0.05)
                        GoTo LoopTest_Start
                    End If
                
                End If
                         
            Else
               TesterReady1 = 1
                  
            End If
            
            '========================
            ' wait for tester2 ready
            '========================
             
            If site2 = 1 Then
                If TesterReady2 = 0 Then
                         
                    If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                        MSComm2.Output = "~"
                                
                        If HUBEnaOn = 0 Then
                            k = DO_WritePort(card, Channel_P1CH, &H0) ' set HUB module Ena on
                            HUBEnaOn = 1
                        End If
                                
                    End If
                                
                    buf2 = MSComm2.Input
                    TesterStatus2 = TesterStatus2 & buf2
                               
                    If (InStr(1, TesterStatus2, "Ready") <> 0) Then
                        TesterReady2 = 1
                    End If
                
                    If (GetGPIBStatus(1) = False) And (VB6_Flag = True) And (Mid(ChipName, 12, 1) <> "U") Then
                        MSComm2.Output = "AUGPIBGPIBGPIB"   ' trans strat test signal to TEST PC
                        MSComm2.InBufferCount = 0
                        MSComm2.InputLen = 0
                        GPIBInquiryTime = Timer
                        'Call MsecDelay(0.02)
                                        
                        Do
                            buf2 = MSComm2.Input
                            TesterStatus2 = TesterStatus2 & buf2
                            Call MsecDelay(0.05)
                        Loop Until (InStr(1, TesterStatus2, "GPIBReady") <> 0) _
                                    Or (InStr(1, TesterStatus2, "GPIBUNReady") <> 0) _
                                    Or (Timer - GPIBInquiryTime > 1)
                        
                        If (Timer - GPIBInquiryTime > 1) Then
                            GoTo LoopTest_Start
                        End If
                        
                        Call MsecDelay(0.05)
                        MSComm2.Output = "AUGPIBACK"
                        GetGPIBStatus(1) = True
                                            
                        If InStr(1, TesterStatus2, "GPIBReady") <> 0 Then
                            GPIBReady(1) = True
                            Site2GPIB_Label.BackColor = &H80FFFF
                            Site2GPIB_Label.ForeColor = &HFF0000
                            Site2GPIB_Label.Caption = " GPIB Rdy"
                        Else
                            GPIBReady(1) = False
                            Site2GPIB_Label.BackColor = &H80FFFF
                            Site2GPIB_Label.ForeColor = &H80FF&
                            Site2GPIB_Label.Caption = " No GPIB"
                        End If
                        Call MsecDelay(0.05)
                        GoTo LoopTest_Start
                    End If
                    
                End If
                         
            Else
               TesterReady2 = 1
                  
            End If
                             
            '===================================
            ' Reset rountine : condsider Reset fail
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
                
                    '=========== Reset  Rountine
                    
                    If TesterReady1 = 0 And TesterDownCount1 = 0 And FirstRun = 1 Then ' reset tester1
                    '============== close module power
                    ResetCounter1 = ResetCounter1 + 1
                       
                '    k = DO_WritePort(card, Channel_P1CL, &HF) ' send 1111 => Set SITE1 power off" Channel_P1CL = 15
                '    k = DO_WritePort(card, Channel_P1CH, &HF) ' send 1111 => Set SITE2 power off" Channel_P1CH = 15
                        
                    k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    '============= Reset PC
                    TesterDownCount1 = 1
                    
                               
                    If Left(ChipName, 10) = "AU6254XLS4" Then
                    k = DO_WritePort(card, Channel_P1CH, &H0)
                    
                    Call MsecDelay(6)
                    
                     k = DO_WritePort(card, Channel_P1CH, &HF)
                    Call MsecDelay(2)
                    
                    End If
         
                    
                    k = DO_WritePort(card, Channel_P1B, &HF) ' send 0000,1111 => RESET PC " Channel_P1B= 15
                    Call MsecDelay(2)
                    k = DO_WritePort(card, Channel_P1B, &HFF) ' send 1111,1111 => RESET PC " Channel_P1B= 255
                    WaitForPowerOn1 = Timer
                    '============== clear comm buffer
                    MSComm1.InBufferCount = 0
                    TesterStatus1 = ""
                
                End If
            
            
                If TesterReady2 = 0 And TesterDownCount2 = 0 And FirstRun = 1 Then ' reset tester2
                    '============== close module power
                    ResetCounter2 = ResetCounter2 + 1
                        
                 '   k = DO_WritePort(card, Channel_P1CL, &HF) ' send 1111 => Set SITE1 power off" Channel_P1CL = 15
                 '   k = DO_WritePort(card, Channel_P1CH, &HF) ' send 1111 => Set SITE2 power off" Channel_P1CH = 15
                
                    k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    '============== Reset PC
                    TesterDownCount2 = 1
                               
                    If Left(ChipName, 10) = "AU6254XLS4" Then
                    k = DO_WritePort(card, Channel_P1CH, &H0)
                    
                    Call MsecDelay(6)
                    
                     k = DO_WritePort(card, Channel_P1CH, &HF)
                    Call MsecDelay(2)
                    
                    End If
         
                    
                    k = DO_WritePort(card, Channel_P1B, &HF) ' send 0000,1111 => RESET PC " Channel_P1B= 15
                    Call MsecDelay(2)
                    k = DO_WritePort(card, Channel_P1B, &HFF) ' send 1111,1111 => RESET PC " Channel_P1B= 255
                    WaitForPowerOn2 = Timer
                    '============== clear comm buffer
                    MSComm2.InBufferCount = 0
                    TesterStatus2 = ""
                
                End If
            
            End If 'If Timer - WaitForReady > 1 Then
            
                                        
            '===============================
            ' screen down count routine
            '==============================
            
             If TesterDownCount1 = 1 Then
             
                 TesterDownCountTimer1 = Timer - WaitForPowerOn1
                 Label28.Caption = CInt(TesterDownCountTimer1)
                 
                 If TesterReady1 = 1 Then
                     '====== open module power
                       
                 '   k = DO_WritePort(card, Channel_P1CL, &HC) ' send 1100 => SetSITE1 power" Channel_P1CL = 12
                 '   k = DO_WritePort(card, Channel_P1CH, &HC) ' send 1100 => SetSITE2 power" Channel_P1CH = 12
                
                    k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 255
                     
                     Call MsecDelay(POWER_ON_TIME)
                     '=== clear flag
                     TesterDownCount1 = 0
                 End If
                 
                 If TesterDownCountTimer1 > 90 Then  'Reset fail
                     TesterDownCount1 = 0
                 End If
             
             
             End If 'If TesterDownCount1 = 1 Then
             
             '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
             
             If TesterDownCount2 = 1 Then
             
                 TesterDownCountTimer2 = Timer - WaitForPowerOn2
                 Label29.Caption = CInt(TesterDownCountTimer2)
                 
                 If TesterReady2 = 1 Then
                     '====== open module power
                        
                 '   k = DO_WritePort(card, Channel_P1CL, &HC) ' send 1100 => SetSITE1 power" Channel_P1CL = 12
                 '   k = DO_WritePort(card, Channel_P1CH, &HC) ' send 1100 => SetSITE2 power" Channel_P1CH = 12
                
                    k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 255
                     Call MsecDelay(POWER_ON_TIME)
                     '=== clear flag
                     TesterDownCount2 = 0
                 End If
                 
                 If TesterDownCountTimer2 > 90 Then    ' Reset fail
                     TesterDownCount2 = 0
                 End If
                 
             End If 'If TesterDownCount2 = 1 Then
              '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%5%%%%%
        End If
                 
        If (Need_GPIB = 1) And (VB6_Flag = True) And (Mid(ChipName, 12, 1) <> "U") Then
            If (TesterReady1 = 1) And (TesterReady2 = 1) Then
                If (site1 = 1) And (site2 = 1) Then 'dual site
                    If (GPIBReady(0) = False) And (GPIBReady(1) = False) Then
                        MsgBox ("Check GPIB CARD Connection Error !!")
                        End
                    End If
                ElseIf (site1 = 1) And (site2 = 0) Then
                    If (GPIBReady(0) = False) Then
                        MsgBox ("Check Site1 GPIB CARD Connection Error !!")
                        End
                    End If
                ElseIf (site1 = 0) And (site2 = 1) Then
                    If (GPIBReady(1) = False) Then
                        MsgBox ("Check Site2 GPIB CARD Connection Error !!")
                        End
                    End If
                End If
            End If
        End If
    
    Loop Until (TesterReady1 = 1) And (TesterReady2 = 1)
    
        FirstRun = 1
         
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'*
'*    Testing Loop
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
If (DI_P >= 12) And (DI_P < 15) Then
  
    
        ' init falg
         GetStart = 1
    Label3.BackColor = RGB(255, 255, 255)
    Label3 = ""
    Label16.BackColor = RGB(255, 255, 255)
    Label16 = ""
         
         
             
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '
    '                Site1 and Site2  begin
    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
             
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '        Testing LED function    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Print "==========================="
          

        ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '
        '  Allen 0526 begin 1 : for no card test,pull high Card detect
        '
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
           If NoCardTest.value = 1 Then
                
                   'pull High
                   
             '   k = DO_WritePort(card, Channel_P1CL, &HD) ' send 1101 => "SetSITE1 Power ON" & " SetSITE1 CDN High"Channel_P1CL = 13
             '   k = DO_WritePort(card, Channel_P1CH, &HD) ' send 1101 => "SetSITE2 Power ON" & " SetSITE2 CDN High"Channel_P1CH = 13
                
                k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power on"& " SetSITE2 CDN High" Channel_P2A = 127
                k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power on"& " SetSITE1 CDN High" Channel_P2A = 127
           Else
                   
              '  k = DO_WritePort(card, Channel_P1CL, &HC) ' send 1100 => "SetSITE1 Power ON" & " SetSITE1 CDN Low"Channel_P1CL = 12
              '  k = DO_WritePort(card, Channel_P1CH, &HC) ' send 1100 => "SetSITE2 Power ON" & " SetSITE2 CDN Low"Channel_P1CH = 12
                
                k = DO_WritePort(card, Channel_P2A, &H3F) ' send 0011,1111 => Set power on"& " SetSITE2 CDN Low" Channel_P2A = 63
                k = DO_WritePort(card, Channel_P2B, &H3F) ' send 0011,1111 => Set power on"& " SetSITE1 CDN Low" Channel_P2A = 63
           End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '
        '  Allen 0526 End  1 : for no card test,pull high Card detect
        '
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
   
    '*STEP4=> Waitting for Response from  Tester\\\\\\\\\\\\\\\\\\\\\
    '*
    '*    Wait Test Result from each Tester
    '*
    '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
   '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      '
      '  Allen 0601 Remark : no card on board test card detect and card change signal
      '
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
        NoCardTestResult1 = ""
        NoCardTestResult2 = ""
        
        If site1 = 1 And NoCardTest.value = 1 Then
    
            MSComm1.Output = ChipName   ' trans strat test signal to TEST PC
            MSComm1.InBufferCount = 0
            MSComm1.InputLen = 0
            NoCardWaitForTest1 = Timer ' wait for timer  and test result
        End If
        
        If site2 = 1 And NoCardTest.value = 1 Then
        
            MSComm2.Output = ChipName   ' trans strat test signal to TEST PC
            MSComm2.InBufferCount = 0
            MSComm2.InputLen = 0
            NoCardWaitForTest2 = Timer
        End If
    
    
        Print "send begin test signal to test"
        TesterStatus1 = ""
        TesterStatus2 = ""
        
      
         Do
            DoEvents
            '========================
            If site1 = 1 And NoCardTest.value = 1 Then
                If NoCardTestStop1 = 0 Then
                
                    If MSComm1.InBufferCount >= 4 Then
                        NoCardTestResult1 = MSComm1.Input
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
            '========================
            
            If site2 = 1 And NoCardTest.value = 1 Then
                If NoCardTestStop2 = 0 Then
                     If MSComm2.InBufferCount >= 4 Then
                            NoCardTestResult2 = MSComm2.Input
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
          '========================
          
        Loop Until (NoCardTestStop1 = 1) And (NoCardTestStop2 = 1)
        
        

    
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      '
      '  Allen 0526 Remark : no card on board test card detect and card change signal
      '
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
    '*STEP3=>Send command to PC teser\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '*
    '*    Send ChipName to PC teser
    '*
    '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    TestResult1 = ""
    TestResult2 = ""
    
    If NoCardTest.value = 1 Then
              '  k = DO_WritePort(card, Channel_P1CL, &HC) ' send 1100 => "SetSITE1 Power ON" & " SetSITE1 CDN Low"Channel_P1CL = 12
              '  k = DO_WritePort(card, Channel_P1CH, &HC) ' send 1100 => "SetSITE2 Power ON" & " SetSITE2 CDN Low"Channel_P1CH = 12
               
                k = DO_WritePort(card, Channel_P2A, &H3F) ' send 0011,1111 => Set power on"& " SetSITE2 CDN Low" Channel_P2A = 63
                k = DO_WritePort(card, Channel_P2B, &H3F) ' send 0011,1111 => Set power on"& " SetSITE1 CDN Low" Channel_P2A = 63
'             Call MsecDelay(0.1)
            If site1 = 1 Then  '****** Continue condition lock at PC tester
            
                MSComm1.Output = NoCardTestResult1   ' only pass can continue at PC Tester
                MSComm1.InBufferCount = 0
                MSComm1.InputLen = 0
                WaitForTest1 = Timer ' wait for timer  and test result
                
                If NoCardTestResult1 <> "PASS" Then
                    TestResult1 = NoCardTestResult1
                End If
            
            End If
          
            
            If site2 = 1 Then
            
                MSComm2.Output = NoCardTestResult2   ' only pass can continue at PC Tester
                MSComm2.InBufferCount = 0
                MSComm2.InputLen = 0
                WaitForTest2 = Timer
            
                If NoCardTestResult2 <> "PASS" Then
                    TestResult2 = NoCardTestResult2
                End If
            End If
    
    
    
    Else
           If site1 = 1 Then
            NoCardTestResult1 = "PASS"
            
                If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                    For i = 1 To Len(ChipName)
                        MSComm1.Output = Mid(ChipName, i, 1)
                        Call MsecDelay(0.02)
                        'Debug.Print Mid(ChipName, i, 1)
                    Next
                Else
                    MSComm1.Output = ChipName   ' trans strat test signal to TEST PC
                End If
                
            MSComm1.InBufferCount = 0
            MSComm1.InputLen = 0
            WaitForTest1 = Timer ' wait for timer  and test result
            End If
            
            
            If site2 = 1 Then
            NoCardTestResult2 = "PASS"
            
                If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                    For i = 1 To Len(ChipName)
                        MSComm2.Output = Mid(ChipName, i, 1)
                        Call MsecDelay(0.02)
                        'Debug.Print Mid(ChipName, i, 1)
                    Next
                Else
                    MSComm2.Output = ChipName   ' trans strat test signal to TEST PC
                End If
                
            MSComm2.InBufferCount = 0
            MSComm2.InputLen = 0
            WaitForTest2 = Timer
            End If
    
           
    End If
'===================================================
  
DoEvents

    DoEvents
    Do
        If site1 = 1 And NoCardTestResult1 = "PASS" Then
            If TestStop1 = 0 Then
            
                If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                    MSComm1.Output = "~"
                End If
                
                If MSComm1.InBufferCount >= 4 Then
                    TestResult1 = MSComm1.Input
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
        '========================
        If site2 = 1 And NoCardTestResult2 = "PASS" Then
            If TestStop2 = 0 Then
                
                If (ChipName = "AU6350BL_1Port") Or (ChipName = "AU6350GL_2Port") Or (ChipName = "AU6350CF_3Port") Or (ChipName = "AU6350AL_4Port") Then
                    MSComm2.Output = "~"
                End If
                
                If MSComm2.InBufferCount >= 4 Then
                    TestResult2 = MSComm2.Input
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
    
    Loop Until (TestStop1 = 1) And (TestStop2 = 1)



    '\\\\\\\\\\\\\\\\\\\\\\\\\\}wait Tester response END
                           
        If site1 = 1 Then
            Print "TestResult1= "; TestResult1
        End If
         
        If site2 = 1 Then
            Print "TestResult2= "; TestResult2
        End If
        
        If HUBEnaOn = 1 Then
            k = DO_WritePort(card, Channel_P1CH, &HF) ' set HUB module Ena off
            HUBEnaOn = 0
        End If
        
        TestCounter = TestCounter + 1 ' Allen Debug
     
        If TestCycleTime1 > WAIT_TEST_CYCLE_OUT And site1 = 1 Then
         
            WaitTestTimeOut1 = 1
            WaitTestTimeOutCounter1 = WaitTestTimeOutCounter1 + 1
        End If
                  
        If TestCycleTime2 > WAIT_TEST_CYCLE_OUT And site2 = 1 Then
         
            WaitTestTimeOut2 = 1
            WaitTestTimeOutCounter2 = WaitTestTimeOutCounter2 + 1
        End If
      
              
    '/////////////////////////////////////////////////////////////////////////
    '
    '   RT Condition
    '
    '//////////////////////////////////////////////////////////////////////////
              
              
     If Check3.value = 1 Then   ' 不RT=> not low yield sorting
         GoTo err
     End If
     
     
     If site1 = 1 And site2 = 0 Then
         If TestResult1 = "PASS" Then
             GoTo err
         End If
         
     End If
     
     
     If site1 = 0 And site2 = 1 Then
         If TestResult2 = "PASS" Then
             GoTo err
         End If
         
     End If
     
     If site1 = 1 And site2 = 1 Then
         If TestResult1 = "PASS" And TestResult2 = "PASS" Then
             GoTo err
         End If
         
     End If
        '////////////////////////////// initial condition
           
         Print "RT begin"
                 
         '1.close power
         '2.delay 10 s
         '3.send power
         '4.RT core
         
         Print "close power"
          
           ' k = DO_WritePort(card, Channel_P1A, &HFF)
           ' k = DO_WritePort(card, Channel_P1B, &HFF)
           
           
        '    k = DO_WritePort(card, Channel_P1CL, &HF) ' send 1111 => "SetSITE1 Power Off" & " SetSITE1 CDN high"Channel_P1CL = 15
        '    k = DO_WritePort(card, Channel_P1CH, &HF) ' send 1111 => "SetSITE2 Power Off" & " SetSITE2 CDN high"Channel_P1CH = 15
           
            k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off"& " SetSITE2 CDN high" Channel_P2A = 255
            k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off"& " SetSITE1 CDN high" Channel_P2A = 255
         Call MsecDelay(RT_INTERVAL)  ' to let system unload driver
         
         Print "Send power"
         'k = DO_WritePort(card, Channel_P1A, &H7F)
         'k = DO_WritePort(card, Channel_P1B, &H7F)
         
           
         '   k = DO_WritePort(card, Channel_P1CL, &HC) ' send 1100 => "SetSITE1 Power ON" & " SetSITE1 CDN Low"Channel_P1CL = 13
         '   k = DO_WritePort(card, Channel_P1CH, &HC) ' send 1100 => "SetSITE2 Power ON" & " SetSITE2 CDN Low"Channel_P1CH = 13
        
            k = DO_WritePort(card, Channel_P2A, &H3F) ' send 0011,1111 => Set power on"& " SetSITE2 CDN Low" Channel_P2A = 63
            k = DO_WritePort(card, Channel_P2B, &H3F) ' send 0011,1111 => Set power on"& " SetSITE1 CDN Low" Channel_P2A = 63
         Call MsecDelay(POWER_ON_TIME)
         
                 
                 
                 
        If site1 = 1 And TestResult1 <> "PASS" Then
            
            MSComm1.Output = ChipName ' trans strat test signal to TEST PC
            MSComm1.InBufferCount = 0
            MSComm1.InputLen = 0
            RTWaitForTest1 = Timer ' wait for timer  and test result
         End If
            
         If site2 = 1 And TestResult2 <> "PASS" Then
             
            MSComm2.Output = ChipName    ' trans strat test signal to TEST PC
            MSComm2.InBufferCount = 0
            MSComm2.InputLen = 0
            RTWaitForTest2 = Timer
         End If
         
        
         Print "send begin test signal to test"
         
         '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Wait for Response from PC Tester
                
            Do
                DoEvents
                
                If site1 = 1 And TestResult1 <> "PASS" Then
                    If RTTestStop1 = 0 Then
                        RTTestResult1 = MSComm1.Input
                        
                        RTTestCycleTime1 = Timer - RTWaitForTest1
                        
                        If (RTTestResult1 <> "" Or RTTestCycleTime1 > WAIT_TEST_CYCLE_OUT) Then
                        
                            RTTestCounter1 = RTTestCounter1 + 1 ' Allen Debug
                            RTTestStop1 = 1
                        End If
                    End If
                
                Else
                    RTTestStop1 = 1
                
                End If
                '========================
                
                If site2 = 1 And TestResult2 <> "PASS" Then
                    If RTTestStop2 = 0 Then
                        RTTestResult2 = MSComm2.Input
                        
                        RTTestCycleTime2 = Timer - RTWaitForTest2
                        
                        If (RTTestResult2 <> "" Or RTTestCycleTime2 > WAIT_TEST_CYCLE_OUT) Then
                            RTTestStop2 = 1
                            RTTestCounter2 = RTTestCounter2 + 1
                        End If
                    End If
                
                Else
                    RTTestStop2 = 1
                
                
                End If
            
            
            Loop Until (RTTestStop1 = 1) And (RTTestStop2 = 1)
                           
                If site1 = 1 Then
                    Print "RTTestResult1= "; RTTestResult1
                End If
                
                If site2 = 1 Then
                    Print "RTTestResult2= "; RTTestResult2
                End If
                
               
                
                If RTTestCycleTime1 > WAIT_TEST_CYCLE_OUT And site1 = 1 Then
                
                    RTWaitTestTimeOut1 = 1
                    RTWaitTestTimeOutCounter1 = RTWaitTestTimeOutCounter1 + 1
                End If
                
                If RTTestCycleTime2 > WAIT_TEST_CYCLE_OUT And site2 = 1 Then
                
                    RTWaitTestTimeOut2 = 1
                    RTWaitTestTimeOutCounter2 = RTWaitTestTimeOutCounter2 + 1
                End If
                 
                 
                 
                     
                 
End If  '////////////////////// Test end
err:              '  Testing Loop end
End Sub

Private Sub AU6368PROTestON()


'*step1=>\\\\\\\\\\\\\\\\\\\\\ Get Start Signal From Handle
'*
'*wait Start Signal From Handle
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    WaitForStart = Timer   ' Get Vcc on from Chip
     
    If TestMode = 0 Then  'ON LINE MODE    'Check1.Value = 0 => TestMode = 0  '上線模式
    
          Print "wait Start"
            Do     'wait  (VCC PowerON) & (hander 5ms start) signal
                   DoEvents
                  WaitStartTime = Timer - WaitForStart
                   k = DO_ReadPort(card, Channel_P2CH, DI_P)
                   'Call MsecDelay(0.1)
        '   Loop Until DI_P = 14 Or DI_P = 13 Or DI_P = 12 Or WaitStartTime > WAIT_START_TIME_OUT  'Allen
           Loop Until DI_P = 14 Or DI_P = 13 Or DI_P = 12
           Label31.Caption = DI_P
            Print "DI_P=", DI_P
            TotalRealTestTime = Timer - OldTotalRealTestTime
            OldTotalRealTestTime = Timer
            OldRealTestTime = Timer
            
     Else                 'Check1.Value = 1 => TestMode = 1   '離線模式
     
            Call MsecDelay(0.8)
            WaitStartTime = 0.8
            DI_P = 14
            TotalRealTestTime = Timer - OldTotalRealTestTime
            OldTotalRealTestTime = Timer
            OldRealTestTime = Timer
            
            'k = DO_WritePort(card, Channel_P1A, 255)
            'k = DO_WritePort(card, Channel_P1B, 255)
            
     End If
     
        buf1 = MSComm1.Input
        Label26.Caption = "實際總測試時間(含 load / unload) :" & TotalRealTestTime & "s"
        WaitStartCounter = WaitStartCounter + 1
                  
    If WaitStartTime > WAIT_START_TIME_OUT Then
         WaitStartTimeOut = 1
         WaitStartTimeOutCounter = WaitStartTimeOutCounter + 1
    End If
    'If ChipName = "AU6368PR" Then ' ARCH add 10/27
        k = DO_WritePort(card, Channel_P1A, &HBF) ' select S1 send 1011,1111 => Set (CDN,IN0)=(1,0)
        Call MsecDelay(0.01)
        Print "Change to CDN=>HIGH state"
       ' k = DO_ReadPort(card, Channel_P1CL, DI_S1)
       ' Print "DI_S1="; DI_S1
       ' k = DO_ReadPort(card, Channel_P1CH, DI_S2)
       ' Print "DI_S2="; DI_S2
    'End If
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'
'    SHOW Alarm
'
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

If Check5.value = 1 Then 'arch change 940529 連續fail 警告

    If site1 = 1 And continuefail1 >= AlarmLimit Then
    
        If continuefail1_bin2 >= 3 Then
            Call MsecDelay(3)
        End If
        
        
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        If continuefail1_bin2 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check Chip Contact and Tester Driver!"
            '  MsgBox "site1 countiue fail please check Chip Contact and Tester Driver!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 1
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 1
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin2 = 0
            continuefail1 = 0
        ElseIf continuefail1_bin3 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check  Flash & CF & SD CARD!"
            
            ' MsgBox "site1 countiue fail please check  Flash & CF & SD CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 2
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 2
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin3 = 0
            continuefail1 = 0
            
        ElseIf continuefail1_bin4 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check XD CARD!"
            
            
            'MsgBox "site1 countiue fail please check XD CARD!"
            
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 3
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
                
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 3
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin4 = 0
            continuefail1 = 0
        ElseIf continuefail1_bin5 >= AlarmLimit Then
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check MS CARD!"
            ' MsgBox "site1 countiue fail please check MS CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 4
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 4
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin5 = 0
            continuefail1 = 0
        Else
            Print "Site1 check continuefail start !"
        End If
    
    End If 'If Site1 = 1 And continuefail1 >= AlarmLimit Then
   '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    If site2 = 1 And continuefail2 >= AlarmLimit Then
    
    
        If continuefail2_bin2 >= 3 Then
            Call MsecDelay(3)
        End If
    
        If continuefail2_bin2 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check Chip Contact and Tester Driver!"
            '  MsgBox "site2 countiue fail please check Chip Contact and Tester Driver!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 5
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 5
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
            continuefail2_bin2 = 0
            continuefail2 = 0
        ElseIf continuefail2_bin3 >= AlarmLimit Then
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check  Flash & CF & SD CARD!"
            'MsgBox "site2 countiue fail please check  Flash & CF & SD CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 6
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 6
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail2_bin3 = 0
            continuefail2 = 0
        ElseIf continuefail2_bin4 >= AlarmLimit Then
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check XD CARD!"
            
            'MsgBox "site2 countiue fail please check XD CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 7
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
        
        
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 7
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
            continuefail2_bin4 = 0
            continuefail2 = 0
        ElseIf continuefail2_bin5 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check MS CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 8
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            
            Print "Alarm!!!"
            Do
            
                DoEvents
                
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            
            Loop While (1)
            
            Print "Alarm Clear"
        
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 8
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
            
            '  MsgBox "site2 countiue fail please check MS CARD!"
            continuefail2_bin5 = 0
            continuefail2 = 0
        Else
            Print "Site2  check continuefail start !!"
        End If
    
    End If 'If Site2 = 1 And continuefail2 >= AlarmLimit Then

Else
    Print "on standard test step!"
End If 'If Check5.Value = 1 Then  連續fail 警告

'*******************************************************
'*
'*   OPEN power
'*
'**********************************************************
    If DI_P < 12 And DI_P >= 15 Then   'Allen 20050607 , change DI_P > 15, to DI_P >= 15
        Print "no start"
        GoTo err
    Else
    
        Print "get start signal!"
    
        Call MsecDelay(CAPACTOR_CHARGE)
        Call MsecDelay(UNLOAD_DRIVER)
        
        k = DO_ReadPort(card, Channel_P1CL, DI_S1)
        k = DO_ReadPort(card, Channel_P1CH, DI_S2)
        k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power" Channel_P2A = 127
        k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power" Channel_P2b = 127
        
        NewPowerOnTime = POWER_ON_TIME - 0.4
    
        If NewPowerOnTime > 0 Then
            Call MsecDelay(NewPowerOnTime)
        End If
        
    End If
   
   
   
'*STEP2=> wait tester send ready signal\\\\\\\\\\\\\\
'*
'*  Check Tester Ready Signal
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


   MSComm2.InBufferCount = 0
   MSComm1.InBufferCount = 0
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
            
            If site1 = 1 Then
                         If TesterReady1 = 0 Then
                         
                               buf1 = MSComm1.Input
                               
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
             
            If site2 = 1 Then
                         If TesterReady2 = 0 Then
                         
                               buf2 = MSComm2.Input
                               
                               TesterStatus2 = TesterStatus2 & buf2
                               
                               If (InStr(1, TesterStatus2, "Ready") <> 0) Then
                                         TesterReady2 = 1
                               End If
                         End If
                         
            Else
               TesterReady2 = 1
                  
            End If
                             
            '===================================
            ' Reset rountine : condsider Reset fail
            '===================================
        
            If Timer - WaitForReady > 1 Then
            
                If ResetCounter1 > 2 Or ResetCounter2 > 2 Then  ' Alarm for reset fail
                
                    Print "Alarm : Reset PC fail"
                    MsgBox "Alarm : Reset PC fail "
                    ResetCounter1 = 0
                    ResetCounter2 = 0
                    Exit Sub
                Else
                
                    '=========== Reset  Rountine
                    
                    If TesterReady1 = 0 And TesterDownCount1 = 0 And FirstRun = 1 Then ' reset tester1
                    '============== close module power
                    ResetCounter1 = ResetCounter1 + 1
                      
                    k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                    k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                    k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    
                    '============= Reset PC
                    
                    TesterDownCount1 = 1
                    k = DO_WritePort(card, Channel_P1B, &HF) ' send 0000,1111 => RESET PC " Channel_P1B= 15
                    Call MsecDelay(2)
                    k = DO_WritePort(card, Channel_P1B, &HFF) ' send 1111,1111 => RESET PC " Channel_P1B= 255
                    WaitForPowerOn1 = Timer
                    '============== clear comm buffer
                    MSComm1.InBufferCount = 0
                    TesterStatus1 = ""
                
                End If
            
            
                If TesterReady2 = 0 And TesterDownCount2 = 0 And FirstRun = 1 Then ' reset tester2
                    '============== close module power
                    ResetCounter2 = ResetCounter2 + 1
                     
                    k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                    k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                    
                    k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    '============== Reset PC
                    TesterDownCount2 = 1
                    k = DO_WritePort(card, Channel_P1B, &HF) ' send 0000,1111 => RESET PC " Channel_P1B= 15
                    Call MsecDelay(2)
                    k = DO_WritePort(card, Channel_P1B, &HFF) ' send 1111,1111 => RESET PC " Channel_P1B= 255
                    WaitForPowerOn2 = Timer
                    '============== clear comm buffer
                    MSComm2.InBufferCount = 0
                    TesterStatus2 = ""
                
                End If
            
            End If 'If Timer - WaitForReady > 1 Then
            
                                        
            '===============================
            ' screen down count routine
            '==============================
            
             If TesterDownCount1 = 1 Then
             
                 TesterDownCountTimer1 = Timer - WaitForPowerOn1
                 Label28.Caption = CInt(TesterDownCountTimer1)
                 
                 If TesterReady1 = 1 Then
                     '====== open module power
                    
                    k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                    k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                    
                    k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 255
                     
                     Call MsecDelay(POWER_ON_TIME)
                     '=== clear flag
                     TesterDownCount1 = 0
                 End If
                 
                 If TesterDownCountTimer1 > 90 Then  'Reset fail
                     TesterDownCount1 = 0
                 End If
             
             
             End If 'If TesterDownCount1 = 1 Then
             
             '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
             
             If TesterDownCount2 = 1 Then
             
                 TesterDownCountTimer2 = Timer - WaitForPowerOn2
                 Label29.Caption = CInt(TesterDownCountTimer2)
                 
                 If TesterReady2 = 1 Then
                     '====== open module power
                    
                    k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                    k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                    
                    k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 255
                     Call MsecDelay(POWER_ON_TIME)
                     '=== clear flag
                     TesterDownCount2 = 0
                 End If
                 
                 If TesterDownCountTimer2 > 90 Then    ' Reset fail
                     TesterDownCount2 = 0
                 End If
                 
             End If 'If TesterDownCount2 = 1 Then
              '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%5%%%%%
        End If
                 
    Loop Until (TesterReady1 = 1) And (TesterReady2 = 1)
    
        FirstRun = 1
         
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'*
'*    Testing Loop
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
If (DI_P >= 12) And (DI_P < 15) Then
  
    
        ' init falg
         GetStart = 1
    Label3.BackColor = RGB(255, 255, 255)
    Label3 = ""
    Label16.BackColor = RGB(255, 255, 255)
    Label16 = ""
         
         
             
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '
    '                Site1 and Site2  begin
    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
             
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '        Testing LED function    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Print "==========================="
          

        ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '
        '  Allen 0526 begin 1 : for no card test,pull high Card detect
        '
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
           If NoCardTest.value = 1 Then
                
                
                k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                
                k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power on"& " SetSITE2 CDN High" Channel_P2A = 127
                k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power on"& " SetSITE1 CDN High" Channel_P2A = 127
                k = DO_WritePort(card, Channel_P1A, &HBF) ' select NOCARD state send 1011,1111 => Set (CDN,IN0)=(1,0)
           Else
                 '   If ChipName = "AU6366S4" Then
                k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                
                k = DO_WritePort(card, Channel_P2A, &H3F) ' send 0011,1111 => Set power on"& " SetSITE2 CDN Low" Channel_P2A = 63
                k = DO_WritePort(card, Channel_P2B, &H3F) ' send 0011,1111 => Set power on"& " SetSITE1 CDN Low" Channel_P2A = 63
           End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '
        '  Allen 0526 End  1 : for no card test,pull high Card detect
        '
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
   
    '*STEP4=> Waitting for Response from  Tester\\\\\\\\\\\\\\\\\\\\\
    '*
    '*    Wait Test Result from each Tester
    '*
    '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
   '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      '
      '  Allen 0601 Remark : no card on board test card detect and card change signal
      '
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
        NoCardTestResult1 = ""
        NoCardTestResult2 = ""
        
        If site1 = 1 And NoCardTest.value = 1 Then
    
            MSComm1.Output = ChipName   ' trans strat test signal to TEST PC
            MSComm1.InBufferCount = 0
            MSComm1.InputLen = 0
            NoCardWaitForTest1 = Timer ' wait for timer  and test result
        End If
        
        If site2 = 1 And NoCardTest.value = 1 Then
        
            MSComm2.Output = ChipName   ' trans strat test signal to TEST PC
            MSComm2.InBufferCount = 0
            MSComm2.InputLen = 0
            NoCardWaitForTest2 = Timer
        End If
    
    
        Print "send begin test signal to test"
        TesterStatus1 = ""
        TesterStatus2 = ""
        
      
         Do
            DoEvents
            '========================
            If site1 = 1 And NoCardTest.value = 1 Then
                If NoCardTestStop1 = 0 Then
                
                    If MSComm1.InBufferCount >= 4 Then
                        NoCardTestResult1 = MSComm1.Input
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
            '========================
            
            If site2 = 1 And NoCardTest.value = 1 Then
                If NoCardTestStop2 = 0 Then
                     If MSComm2.InBufferCount >= 4 Then
                            NoCardTestResult2 = MSComm2.Input
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
          '========================
          
        Loop Until (NoCardTestStop1 = 1) And (NoCardTestStop2 = 1)
    
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      '
      '  Allen 0526 Remark : no card on board test card detect and card change signal
      '
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
    '*STEP3=>Send command to PC teser\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '*
    '*    Send ChipName to PC teser
    '*
    '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    TestResult1 = ""
    TestResult2 = ""
    
    If NoCardTest.value = 1 Then
                k = DO_WritePort(card, Channel_P1A, &H3F) ' select S1 send 0011,1111 => Set (CDN,IN0)=(0,0)
                k = DO_WritePort(card, Channel_P2A, &H3F) ' send 0011,1111 => Set power on"& " SetSITE2 CDN Low" Channel_P2A = 63
                k = DO_WritePort(card, Channel_P2B, &H3F) ' send 0011,1111 => Set power on"& " SetSITE1 CDN Low" Channel_P2A = 63
'
            If site1 = 1 Then  '****** Continue condition lock at PC tester
            
            MSComm1.Output = NoCardTestResult1   ' only pass can continue at PC Tester
            MSComm1.InBufferCount = 0
            MSComm1.InputLen = 0
            WaitForTest1 = Timer ' wait for timer  and test result
             
              If NoCardTestResult1 <> "PASS" Then
                 TestResult1 = NoCardTestResult1
              End If
            
            End If
          
            
            If site2 = 1 Then
            
                MSComm2.Output = NoCardTestResult2   ' only pass can continue at PC Tester
                MSComm2.InBufferCount = 0
                MSComm2.InputLen = 0
                WaitForTest2 = Timer
            
                If NoCardTestResult2 <> "PASS" Then
                    TestResult2 = NoCardTestResult2
                End If
            End If
    
    
    
    Else
          '  If ChipName = "AU6366S4" Then
                k = DO_WritePort(card, Channel_P1A, &H3F) ' select S1 send 0011,1111 => Set (CDN,IN0)=(0,0)  and Channel_P2A = 63
                Call MsecDelay(0.1)
                Print "Change to SD state"
                k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                Print "DI_S1="; DI_S1
                k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                Print "DI_S2="; DI_S2
                 
           '  End If
    
    
    
            If site1 = 1 And DI_S1 = 14 Then
            ' If site1 = 1 Then
                MSComm1.Output = ChipName   ' trans strat test signal to TEST PC
                MSComm1.InBufferCount = 0
                MSComm1.InputLen = 0
                WaitForTest1 = Timer ' wait for timer  and test result
             Else
                TestResult1 = "CARD initial FAIL"
                TestStop1 = 1
             End If
            
            
           If site2 = 1 And DI_S2 = 14 Then
             'If site2 = 1 Then
                MSComm2.Output = ChipName   ' trans strat test signal to TEST PC
                MSComm2.InBufferCount = 0
                MSComm2.InputLen = 0
                WaitForTest2 = Timer
             Else
                TestResult2 = "CARD initial FAIL"
                TestStop2 = 1
             End If
    End If
'===================================================
  
DoEvents
    Do
            DoEvents 'wait test slot0,1,2,3 result "68_PASS"
     
                If site1 = 1 Then
                    If TestStop1 = 0 Then
                    
                        If MSComm1.InBufferCount >= 4 Then
                            TestResult1 = MSComm1.Input
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
                
                   '========================
                   
                If site2 = 1 Then
                    If TestStop2 = 0 Then
                    
                        If MSComm2.InBufferCount >= 4 Then
                            TestResult2 = MSComm2.Input
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
    Loop Until (TestStop1 = 1) And (TestStop2 = 1)
    '==================='wait test "69_PASS" result end==========================
               Label3 = TestResult1
               Label16 = TestResult2
                If TestResult1 = "68_PASS" Or TestResult2 = "68_PASS" Then
                    k = DO_WritePort(card, Channel_P1A, &HFF) ' select S2 send 1111,1111 => Set (CDN,IN0)=(1,1)  and Channel_P2A = 255
                    Call MsecDelay(0.1)
                    k = DO_WritePort(card, Channel_P1A, &H7F) ' select S2 send 0111,1111 => Set (CDN,IN0)=(0,1)  and Channel_P2A = 127
                    Print "Change to MS_PRO CARD state"
                    Call MsecDelay(0.2)
                End If
                    
                If site1 = 1 And TestResult1 = "68_PASS" Then
                
                    OldTestResult1 = TestResult1
                    MSComm1.Output = TestResult1   ' trans strat test signal to TEST PC
                    MSComm1.InBufferCount = 0
                    MSComm1.InputLen = 0
                    WaitForTest1 = Timer ' wait for timer  and test result
                    TestStop1 = 0 'reset 初始條件設定
                    TestResult1 = ""
                End If
                
                If site2 = 1 And TestResult2 = "68_PASS" Then
                
                    OldTestResult2 = TestResult2
                    MSComm2.Output = TestResult2   ' trans strat test signal to TEST PC
                    MSComm2.InBufferCount = 0
                    MSComm2.InputLen = 0
                    WaitForTest2 = Timer
                    TestStop2 = 0 'reset 初始條件設定
                    TestResult2 = ""
                End If
        
      Do   '===============wait test slot3 "MS_PRO" result
            DoEvents
     
                If site1 = 1 And OldTestResult1 = "68_PASS" Then
                  
                    If TestStop1 = 0 Then
                       
                        If MSComm1.InBufferCount >= 4 Then
                            TestResult1 = MSComm1.Input
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
                
                   '========================
                   
                If site2 = 1 And OldTestResult2 = "68_PASS" Then
                    If TestStop2 = 0 Then
                    
                        TestResult2 = ""
                        If MSComm2.InBufferCount >= 4 Then
                            TestResult2 = MSComm2.Input
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
           
    Loop Until (TestStop1 = 1) And (TestStop2 = 1)
     
    
     '======================'wait test slot3 "MS PRO"  result end==================================

        
    '\\\\\\\\\\\\\\\\\\\\\\\\\\}wait Tester response END
                           
         If site1 = 1 Then
             Print "TestResult1= "; TestResult1
            
         End If
         
         If site2 = 1 Then
             Print "TestResult2= "; TestResult2
                            
         End If
        
        
         TestCounter = TestCounter + 1 ' Allen Debug
     
         If TestCycleTime1 > WAIT_TEST_CYCLE_OUT And site1 = 1 Then
         
           WaitTestTimeOut1 = 1
           WaitTestTimeOutCounter1 = WaitTestTimeOutCounter1 + 1
         End If
                  
         If TestCycleTime2 > WAIT_TEST_CYCLE_OUT And site2 = 1 Then
         
           WaitTestTimeOut2 = 1
           WaitTestTimeOutCounter2 = WaitTestTimeOutCounter2 + 1
         End If
      
              
    '/////////////////////////////////////////////////////////////////////////
    '
    '   RT Condition
    '
    '//////////////////////////////////////////////////////////////////////////
              
              
     If Check3.value = 1 Then   ' 不RT=> not low yield sorting
         GoTo err
     End If
     
     
     If site1 = 1 And site2 = 0 Then
         If TestResult1 = "PASS" Then
             GoTo err
         End If
         
     End If
     
     
     If site1 = 0 And site2 = 1 Then
         If TestResult2 = "PASS" Then
             GoTo err
         End If
         
     End If
     
     If site1 = 1 And site2 = 1 Then
         If TestResult1 = "PASS" And TestResult2 = "PASS" Then
             GoTo err
         End If
         
     End If
        '////////////////////////////// initial condition
           
         Print "RT begin"
                 
         '1.close power
         '2.delay 10 s
         '3.send power
         '4.RT core
         
         Print "close power"
            
            k = DO_ReadPort(card, Channel_P1CL, DI_S1)
            k = DO_ReadPort(card, Channel_P1CH, DI_S2)
            
            k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off"& " SetSITE2 CDN high" Channel_P2A = 255
            k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off"& " SetSITE1 CDN high" Channel_P2A = 255
         Call MsecDelay(RT_INTERVAL)  ' to let system unload driver
         
         Print "Send power"
         
            k = DO_ReadPort(card, Channel_P1CL, DI_S1)
            k = DO_ReadPort(card, Channel_P1CH, DI_S2)
            
            k = DO_WritePort(card, Channel_P2A, &H3F) ' send 0011,1111 => Set power on"& " SetSITE2 CDN Low" Channel_P2A = 63
            k = DO_WritePort(card, Channel_P2B, &H3F) ' send 0011,1111 => Set power on"& " SetSITE1 CDN Low" Channel_P2A = 63
         Call MsecDelay(POWER_ON_TIME)
         
                 
                 
                 
        If site1 = 1 And TestResult1 <> "PASS" Then
            
            MSComm1.Output = ChipName ' trans strat test signal to TEST PC
            MSComm1.InBufferCount = 0
            MSComm1.InputLen = 0
            RTWaitForTest1 = Timer ' wait for timer  and test result
         End If
            
         If site2 = 1 And TestResult2 <> "PASS" Then
             
            MSComm2.Output = ChipName    ' trans strat test signal to TEST PC
            MSComm2.InBufferCount = 0
            MSComm2.InputLen = 0
            RTWaitForTest2 = Timer
         End If
         
        
         Print "send begin test signal to test"
         
         '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Wait for Response from PC Tester
                
            Do
                DoEvents
                
                If site1 = 1 And TestResult1 <> "PASS" Then
                    If RTTestStop1 = 0 Then
                        RTTestResult1 = MSComm1.Input
                        
                        RTTestCycleTime1 = Timer - RTWaitForTest1
                        
                        If (RTTestResult1 <> "" Or RTTestCycleTime1 > WAIT_TEST_CYCLE_OUT) Then
                        
                            RTTestCounter1 = RTTestCounter1 + 1 ' Allen Debug
                            RTTestStop1 = 1
                        End If
                    End If
                
                Else
                    RTTestStop1 = 1
                
                End If
                '========================
                
                If site2 = 1 And TestResult2 <> "PASS" Then
                    If RTTestStop2 = 0 Then
                        RTTestResult2 = MSComm2.Input
                        
                        RTTestCycleTime2 = Timer - RTWaitForTest2
                        
                        If (RTTestResult2 <> "" Or RTTestCycleTime2 > WAIT_TEST_CYCLE_OUT) Then
                            RTTestStop2 = 1
                            RTTestCounter2 = RTTestCounter2 + 1
                        End If
                    End If
                
                Else
                    RTTestStop2 = 1
                
                
                End If
            
            
            Loop Until (RTTestStop1 = 1) And (RTTestStop2 = 1)
                           
                If site1 = 1 Then
                    Print "RTTestResult1= "; RTTestResult1
                End If
                
                If site2 = 1 Then
                    Print "RTTestResult2= "; RTTestResult2
                End If
                
               
                
                If RTTestCycleTime1 > WAIT_TEST_CYCLE_OUT And site1 = 1 Then
                
                    RTWaitTestTimeOut1 = 1
                    RTWaitTestTimeOutCounter1 = RTWaitTestTimeOutCounter1 + 1
                End If
                
                If RTTestCycleTime2 > WAIT_TEST_CYCLE_OUT And site2 = 1 Then
                
                    RTWaitTestTimeOut2 = 1
                    RTWaitTestTimeOutCounter2 = RTWaitTestTimeOutCounter2 + 1
                End If
                 
                 
                 
                     
                 
End If  '////////////////////// Test end
err:              '  Testing Loop end
End Sub
Private Sub SmartCardTestON()


'*step1=>\\\\\\\\\\\\\\\\\\\\\ Get Start Signal From Handle
'*
'*wait Start Signal From Handle
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    WaitForStart = Timer   ' Get Vcc on from Chip
     
    If TestMode = 0 Then  'ON LINE MODE    'Check1.Value = 0 => TestMode = 0  '上線模式
    
          Print "wait Start"
            Do     'wait  (VCC PowerON) & (hander 5ms start) signal
                   DoEvents
                  WaitStartTime = Timer - WaitForStart
                   k = DO_ReadPort(card, Channel_P2CH, DI_P)
                   'Call MsecDelay(0.1)
        '   Loop Until DI_P = 14 Or DI_P = 13 Or DI_P = 12 Or WaitStartTime > WAIT_START_TIME_OUT  'Allen
           Loop Until DI_P = 14 Or DI_P = 13 Or DI_P = 12
           Label31.Caption = DI_P
            Print "DI_P=", DI_P
            TotalRealTestTime = Timer - OldTotalRealTestTime
            OldTotalRealTestTime = Timer
            OldRealTestTime = Timer
            
     Else                 'Check1.Value = 1 => TestMode = 1   '離線模式
     
            Call MsecDelay(0.8)
            WaitStartTime = 0.8
            DI_P = 14
            TotalRealTestTime = Timer - OldTotalRealTestTime
            OldTotalRealTestTime = Timer
            OldRealTestTime = Timer
            
            'k = DO_WritePort(card, Channel_P1A, 255)
            'k = DO_WritePort(card, Channel_P1B, 255)
            
     End If
     
        buf1 = MSComm1.Input
        Label26.Caption = "實際總測試時間(含 load / unload) :" & TotalRealTestTime & "s"
        WaitStartCounter = WaitStartCounter + 1
                  
    If WaitStartTime > WAIT_START_TIME_OUT Then
         WaitStartTimeOut = 1
         WaitStartTimeOutCounter = WaitStartTimeOutCounter + 1
    End If
    'If ChipName = "AU9520"or"AU9510" Then ' ARCH add 10/27
        k = DO_WritePort(card, Channel_P1A, &HFF) ' select S1 send 1111,1111 => Set (xs2ena,xs1ena)=(1,0)
        Call MsecDelay(0.01)
        Print "take smart card out=>HIGH state"
       ' k = DO_ReadPort(card, Channel_P1CL, DI_S1)
       ' Print "DI_S1="; DI_S1
       ' k = DO_ReadPort(card, Channel_P1CH, DI_S2)
       ' Print "DI_S2="; DI_S2
    'End If
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'
'    SHOW Alarm
'
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

If Check5.value = 1 Then 'arch change 940529 連續fail 警告

    If site1 = 1 And continuefail1 >= AlarmLimit Then
    
        If continuefail1_bin2 >= 3 Then
            Call MsecDelay(3)
        End If
        
        
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        If continuefail1_bin2 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check Chip Contact and Tester Driver!"
            '  MsgBox "site1 countiue fail please check Chip Contact and Tester Driver!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 1
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 1
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin2 = 0
            continuefail1 = 0
        ElseIf continuefail1_bin3 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check  Flash & CF & SD CARD!"
            
            ' MsgBox "site1 countiue fail please check  Flash & CF & SD CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 2
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 2
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin3 = 0
            continuefail1 = 0
            
        ElseIf continuefail1_bin4 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check XD CARD!"
            
            
            'MsgBox "site1 countiue fail please check XD CARD!"
            
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 3
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
                
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 3
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin4 = 0
            continuefail1 = 0
        ElseIf continuefail1_bin5 >= AlarmLimit Then
            Alarm.Show
            Alarm.Label1 = "site1 countiue fail please check MS CARD!"
            ' MsgBox "site1 countiue fail please check MS CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 4
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 4
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail1_bin5 = 0
            continuefail1 = 0
        Else
            Print "Site1 check continuefail start !"
        End If
    
    End If 'If Site1 = 1 And continuefail1 >= AlarmLimit Then
   '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    If site2 = 1 And continuefail2 >= AlarmLimit Then
    
    
        If continuefail2_bin2 >= 3 Then
            Call MsecDelay(3)
        End If
    
        If continuefail2_bin2 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check Chip Contact and Tester Driver!"
            '  MsgBox "site2 countiue fail please check Chip Contact and Tester Driver!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 5
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 5
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
            continuefail2_bin2 = 0
            continuefail2 = 0
        ElseIf continuefail2_bin3 >= AlarmLimit Then
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check  Flash & CF & SD CARD!"
            'MsgBox "site2 countiue fail please check  Flash & CF & SD CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 6
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
            
            
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 6
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            continuefail2_bin3 = 0
            continuefail2 = 0
        ElseIf continuefail2_bin4 >= AlarmLimit Then
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check XD CARD!"
            
            'MsgBox "site2 countiue fail please check XD CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 7
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            Print "Alarm!!!"
            Do
                DoEvents
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            Loop While (1)
        
        
            Print "Alarm Clear"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 7
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
            continuefail2_bin4 = 0
            continuefail2 = 0
        ElseIf continuefail2_bin5 >= AlarmLimit Then
        
            Alarm.Show
            Alarm.Label1 = "site2 countiue fail please check MS CARD!"
            
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 begin 8
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            AlarmCtrl = 1
            Cls
            
            Print "Alarm!!!"
            Do
            
                DoEvents
                
                If AlarmCtrl = 0 Then
                    Exit Do
                End If
            
            Loop While (1)
            
            Print "Alarm Clear"
        
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '
            ' Allen 20050606 end 8
            '
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
            
            '  MsgBox "site2 countiue fail please check MS CARD!"
            continuefail2_bin5 = 0
            continuefail2 = 0
        Else
            Print "Site2  check continuefail start !!"
        End If
    
    End If 'If Site2 = 1 And continuefail2 >= AlarmLimit Then

Else
    Print "on standard test step!"
End If 'If Check5.Value = 1 Then  連續fail 警告

'*******************************************************
'*
'*   OPEN power
'*
'**********************************************************
    If DI_P < 12 And DI_P >= 15 Then   'Allen 20050607 , change DI_P > 15, to DI_P >= 15
        Print "no start"
        GoTo err
    Else
    
        Print "get start signal!"
    
        Call MsecDelay(CAPACTOR_CHARGE)
        Call MsecDelay(UNLOAD_DRIVER)
        
        k = DO_ReadPort(card, Channel_P1CL, DI_S1)
        k = DO_ReadPort(card, Channel_P1CH, DI_S2)
        k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power" Channel_P2A = 127
        k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power" Channel_P2b = 127
        
        NewPowerOnTime = POWER_ON_TIME - 0.4
    
        If NewPowerOnTime > 0 Then
            Call MsecDelay(NewPowerOnTime)
        End If
        
    End If
   
   
   
'*STEP2=> wait tester send ready signal\\\\\\\\\\\\\\
'*
'*  Check Tester Ready Signal
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


   MSComm2.InBufferCount = 0
   MSComm1.InBufferCount = 0
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
            
            If site1 = 1 Then
                         If TesterReady1 = 0 Then
                         
                               buf1 = MSComm1.Input
                               
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
             
            If site2 = 1 Then
                         If TesterReady2 = 0 Then
                         
                               buf2 = MSComm2.Input
                               
                               TesterStatus2 = TesterStatus2 & buf2
                               
                               If (InStr(1, TesterStatus2, "Ready") <> 0) Then
                                         TesterReady2 = 1
                               End If
                         End If
                         
            Else
               TesterReady2 = 1
                  
            End If
                             
            '===================================
            ' Reset rountine : condsider Reset fail
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
                
                    '=========== Reset  Rountine
                    
                    If TesterReady1 = 0 And TesterDownCount1 = 0 And FirstRun = 1 Then ' reset tester1
                    '============== close module power
                    ResetCounter1 = ResetCounter1 + 1
                      
                    k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                    k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                    k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    
                    '============= Reset PC
                    
                    TesterDownCount1 = 1
                    k = DO_WritePort(card, Channel_P1B, &HF) ' send 0000,1111 => RESET PC " Channel_P1B= 15
                    Call MsecDelay(2)
                    k = DO_WritePort(card, Channel_P1B, &HFF) ' send 1111,1111 => RESET PC " Channel_P1B= 255
                    WaitForPowerOn1 = Timer
                    '============== clear comm buffer
                    MSComm1.InBufferCount = 0
                    TesterStatus1 = ""
                
                End If
            
            
                If TesterReady2 = 0 And TesterDownCount2 = 0 And FirstRun = 1 Then ' reset tester2
                    '============== close module power
                    ResetCounter2 = ResetCounter2 + 1
                     
                    k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                    k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                    
                    k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off" Channel_P2A = 255
                    '============== Reset PC
                    TesterDownCount2 = 1
                    k = DO_WritePort(card, Channel_P1B, &HF) ' send 0000,1111 => RESET PC " Channel_P1B= 15
                    Call MsecDelay(2)
                    k = DO_WritePort(card, Channel_P1B, &HFF) ' send 1111,1111 => RESET PC " Channel_P1B= 255
                    WaitForPowerOn2 = Timer
                    '============== clear comm buffer
                    MSComm2.InBufferCount = 0
                    TesterStatus2 = ""
                
                End If
            
            End If 'If Timer - WaitForReady > 1 Then
            
                                        
            '===============================
            ' screen down count routine
            '==============================
            
             If TesterDownCount1 = 1 Then
             
                 TesterDownCountTimer1 = Timer - WaitForPowerOn1
                 Label28.Caption = CInt(TesterDownCountTimer1)
                 
                 If TesterReady1 = 1 Then
                     '====== open module power
                    
                    k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                    k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                    
                    k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 127
                    k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 127
                    k = DO_WritePort(card, Channel_P1A, &HFF) ' send 1111,1111 => Set bus swith off" Channel_P2A = 255
                     Call MsecDelay(POWER_ON_TIME)
                     '=== clear flag
                     TesterDownCount1 = 0
                 End If
                 
                 If TesterDownCountTimer1 > 90 Then  'Reset fail
                     TesterDownCount1 = 0
                 End If
             
             
             End If 'If TesterDownCount1 = 1 Then
             
             '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
             
             If TesterDownCount2 = 1 Then
             
                 TesterDownCountTimer2 = Timer - WaitForPowerOn2
                 Label29.Caption = CInt(TesterDownCountTimer2)
                 
                 If TesterReady2 = 1 Then
                     '====== open module power
                    
                    k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                    k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                    
                    k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power on" Channel_P2A = 255
                    k = DO_WritePort(card, Channel_P1A, &HFF) ' send 1111,1111 => Set bus swith off" Channel_P2A = 255
                    
                     Call MsecDelay(POWER_ON_TIME)
                     '=== clear flag
                     TesterDownCount2 = 0
                 End If
                 
                 If TesterDownCountTimer2 > 90 Then    ' Reset fail
                     TesterDownCount2 = 0
                 End If
                 
             End If 'If TesterDownCount2 = 1 Then
              '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%5%%%%%
        End If
                 
    Loop Until (TesterReady1 = 1) And (TesterReady2 = 1)
    
        FirstRun = 1
         
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'*
'*    Testing Loop
'*
'*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
If (DI_P >= 12) And (DI_P < 15) Then
  
    
        ' init falg
         GetStart = 1
    Label3.BackColor = RGB(255, 255, 255)
    Label3 = ""
    Label16.BackColor = RGB(255, 255, 255)
    Label16 = ""
         
         
             
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '
    '                Site1 and Site2  begin
    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
             
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '        Testing LED function    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Print "==========================="
          

        ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '
        '  Allen 0526 begin 1 : for no card test,pull high Card detect
        '
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
           If NoCardTest.value = 1 Then
                
                
                k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                
                k = DO_WritePort(card, Channel_P2A, &H7F) ' send 0111,1111 => Set power on"& " SetSITE2 CDN High" Channel_P2A = 127
                k = DO_WritePort(card, Channel_P2B, &H7F) ' send 0111,1111 => Set power on"& " SetSITE1 CDN High" Channel_P2A = 127
                'k = DO_WritePort(card, Channel_P1A, &HBF) ' select NOCARD state send 1011,1111 => Set (CDN,IN0)=(1,0)
                k = DO_WritePort(card, Channel_P1A, &HFF) ' send 1111,1111 => Set bus swith off" Channel_P2A = 255
                    
           Else
                 '   If ChipName = "AU6366S4" Then
                k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                
                k = DO_WritePort(card, Channel_P2A, &H3F) ' send 0011,1111 => Set power on"& " SetSITE2 CDN Low" Channel_P2A = 63
                k = DO_WritePort(card, Channel_P2B, &H3F) ' send 0011,1111 => Set power on"& " SetSITE1 CDN Low" Channel_P2A = 63
                k = DO_WritePort(card, Channel_P1A, &H3F) ' send 1111,1111 => Set bus swith on" Channel_P2A = 255
                    
           End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '
        '  Allen 0526 End  1 : for no card test,pull high Card detect
        '
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            
   
    '*STEP4=> Waitting for Response from  Tester\\\\\\\\\\\\\\\\\\\\\
    '*
    '*    Wait Test Result from each Tester
    '*
    '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
   '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      '
      '  Allen 0601 Remark : no card on board test card detect and card change signal
      '
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
        NoCardTestResult1 = ""
        NoCardTestResult2 = ""
        
     '   If site1 = 1 And NoCardTest.Value = 1 Then
    
     '       MSComm1.Output = ChipName   ' trans strat test signal to TEST PC
      '      MSComm1.InBufferCount = 0
     '       MSComm1.InputLen = 0
     '       NoCardWaitForTest1 = Timer ' wait for timer  and test result
     '   End If
        
     '   If site2 = 1 And NoCardTest.Value = 1 Then
        
    '       MSComm2.Output = ChipName   ' trans strat test signal to TEST PC
    '        MSComm2.InBufferCount = 0
    '        MSComm2.InputLen = 0
    '        NoCardWaitForTest2 = Timer
    '    End If
    
    
    '    Print "send begin test signal to test"
        TesterStatus1 = ""
        TesterStatus2 = ""
        
      
         Do
            DoEvents
            '========================
            If site1 = 1 And NoCardTest.value = 1 Then
                If NoCardTestStop1 = 0 Then
                
                    'If MSComm1.InBufferCount >= 4 Then
                   '     NoCardTestResult1 = MSComm1.Input
                   ' End If
                                       
                   ' NoCardTestResult1 = Parser(NoCardTestResult1)
                   k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                   k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                   If DI_S1 = 11 Then
                       NoCardTestResult1 = ChipName
                    Else
                       NoCardTestResult1 = "BIN2"
                    End If
                    
                    NoCardTestCycleTime1 = Timer - NoCardWaitForTest1
                    
                    If (NoCardTestResult1 <> "" Or NoCardTestCycleTime1 > NO_CARD_TEST_TIME) Then
                       NoCardTestStop1 = 1
                    End If
                End If
            
            Else
                NoCardTestStop1 = 1
            End If
            '========================
            
            If site2 = 1 And NoCardTest.value = 1 Then
                If NoCardTestStop2 = 0 Then
                    ' If MSComm2.InBufferCount >= 4 Then
                     '       NoCardTestResult2 = MSComm2.Input
                    ' End If
                    
                    ' NoCardTestResult2 = Parser(NoCardTestResult2)
                    k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                    k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                    If DI_S2 = 11 Then
                         NoCardTestResult2 = ChipName
                    Else
                         NoCardTestResult2 = "BIN2"
                    End If
                    
                     NoCardTestCycleTime2 = Timer - NoCardWaitForTest2
                    
                    If (NoCardTestResult2 <> "" Or NoCardTestCycleTime2 > NO_CARD_TEST_TIME) Then
                        NoCardTestStop2 = 1
                        
                    End If
                End If
            
            Else
                NoCardTestStop2 = 1
            End If
          '========================
          
        Loop Until (NoCardTestStop1 = 1) And (NoCardTestStop2 = 1)
    
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      '
      '  Allen 0526 Remark : no card on board test card detect and card change signal
      '
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
    '*STEP3=>Send command to PC teser\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '*
    '*    Send ChipName to PC teser
    '*
    '*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    TestResult1 = ""
    TestResult2 = ""
    
    If NoCardTest.value = 1 Then
                k = DO_WritePort(card, Channel_P1A, &H3F) ' select S1 send 0011,1111 => Set (CDN,IN0)=(0,0)
                k = DO_WritePort(card, Channel_P2A, &H3F) ' send 0011,1111 => Set power on"& " SetSITE2 CDN Low" Channel_P2A = 63
                k = DO_WritePort(card, Channel_P2B, &H3F) ' send 0011,1111 => Set power on"& " SetSITE1 CDN Low" Channel_P2A = 63
                Call MsecDelay(0.5) '''arch add 95/3/3
            If site1 = 1 Then  '****** Continue condition lock at PC tester
            
            MSComm1.Output = NoCardTestResult1   ' only pass can continue at PC Tester
            MSComm1.InBufferCount = 0
            MSComm1.InputLen = 0
            WaitForTest1 = Timer ' wait for timer  and test result
             
              If NoCardTestResult1 <> "PASS" Then
                 TestResult1 = NoCardTestResult1
              End If
            
            End If
          
            
            If site2 = 1 Then
            
                MSComm2.Output = NoCardTestResult2   ' only pass can continue at PC Tester
                MSComm2.InBufferCount = 0
                MSComm2.InputLen = 0
                WaitForTest2 = Timer
            
                If NoCardTestResult2 <> "PASS" Then
                    TestResult2 = NoCardTestResult2
                End If
            End If
    
    
    
    Else
          '  If ChipName = "AU6366S4" Then
                k = DO_WritePort(card, Channel_P1A, &H3F) ' select S1 send 0011,1111 => Set (CDN,IN0)=(0,0)  and Channel_P2A = 63
                Call MsecDelay(0.1)
                Print "Smart Card on test state"
                k = DO_ReadPort(card, Channel_P1CL, DI_S1)
                Print "DI_S1="; DI_S1
                k = DO_ReadPort(card, Channel_P1CH, DI_S2)
                Print "DI_S2="; DI_S2
                 
           '  End If
    
    
    
            If site1 = 1 And DI_S1 = 8 Then
            ' If site1 = 1 Then
                MSComm1.Output = ChipName   ' trans strat test signal to TEST PC
                MSComm1.InBufferCount = 0
                MSComm1.InputLen = 0
                WaitForTest1 = Timer ' wait for timer  and test result
             Else
                TestResult1 = "CARD initial FAIL"
                TestStop1 = 1
             End If
            
            
           If site2 = 1 And DI_S2 = 8 Then
             'If site2 = 1 Then
                MSComm2.Output = ChipName   ' trans strat test signal to TEST PC
                MSComm2.InBufferCount = 0
                MSComm2.InputLen = 0
                WaitForTest2 = Timer
             Else
                TestResult2 = "CARD initial FAIL"
                TestStop2 = 1
             End If
    End If
'===================================================
  
DoEvents
    Do
            DoEvents 'wait test slot0,1,2,3 result "68_PASS"
     
                If site1 = 1 Then
                    If TestStop1 = 0 Then
                    
                        If MSComm1.InBufferCount >= 4 Then
                            TestResult1 = MSComm1.Input
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
                
                   '========================
                   
                If site2 = 1 Then
                    If TestStop2 = 0 Then
                    
                        If MSComm2.InBufferCount >= 4 Then
                            TestResult2 = MSComm2.Input
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
    Loop Until (TestStop1 = 1) And (TestStop2 = 1)
   
         If site1 = 1 Then
             Print "TestResult1= "; TestResult1
            
         End If
         
         If site2 = 1 Then
             Print "TestResult2= "; TestResult2
                            
         End If
        
        
         TestCounter = TestCounter + 1 ' Allen Debug
     
         If TestCycleTime1 > WAIT_TEST_CYCLE_OUT And site1 = 1 Then
         
           WaitTestTimeOut1 = 1
           WaitTestTimeOutCounter1 = WaitTestTimeOutCounter1 + 1
         End If
                  
         If TestCycleTime2 > WAIT_TEST_CYCLE_OUT And site2 = 1 Then
         
           WaitTestTimeOut2 = 1
           WaitTestTimeOutCounter2 = WaitTestTimeOutCounter2 + 1
         End If
      
              
    '/////////////////////////////////////////////////////////////////////////
    '
    '   RT Condition
    '
    '//////////////////////////////////////////////////////////////////////////
              
              
     If Check3.value = 1 Then   ' 不RT=> not low yield sorting
         GoTo err
     End If
     
     
     If site1 = 1 And site2 = 0 Then
         If TestResult1 = "PASS" Then
             GoTo err
         End If
         
     End If
     
     
     If site1 = 0 And site2 = 1 Then
         If TestResult2 = "PASS" Then
             GoTo err
         End If
         
     End If
     
     If site1 = 1 And site2 = 1 Then
         If TestResult1 = "PASS" And TestResult2 = "PASS" Then
             GoTo err
         End If
         
     End If
        '////////////////////////////// initial condition
           
         Print "RT begin"
                 
         '1.close power
         '2.delay 10 s
         '3.send power
         '4.RT core
         
         Print "close power"
            
            k = DO_ReadPort(card, Channel_P1CL, DI_S1)
            k = DO_ReadPort(card, Channel_P1CH, DI_S2)
            
            k = DO_WritePort(card, Channel_P1A, &HFF) ' send 1111,1111 => Set bus swith off" Channel_P2A = 255
                    
            
            
            k = DO_WritePort(card, Channel_P2A, &HFF) ' send 1111,1111 => Set power off"& " SetSITE2 CDN high" Channel_P2A = 255
            k = DO_WritePort(card, Channel_P2B, &HFF) ' send 1111,1111 => Set power off"& " SetSITE1 CDN high" Channel_P2A = 255
         Call MsecDelay(RT_INTERVAL)  ' to let system unload driver
         
         Print "Send power"
         
            k = DO_ReadPort(card, Channel_P1CL, DI_S1)
            k = DO_ReadPort(card, Channel_P1CH, DI_S2)
            
            k = DO_WritePort(card, Channel_P1A, &H3F) ' send 0011,1111 => Set bus swith on" Channel_P2A = 255
                    
            
            k = DO_WritePort(card, Channel_P2A, &H3F) ' send 0011,1111 => Set power on"& " SetSITE2 CDN Low" Channel_P2A = 63
            k = DO_WritePort(card, Channel_P2B, &H3F) ' send 0011,1111 => Set power on"& " SetSITE1 CDN Low" Channel_P2A = 63
         Call MsecDelay(POWER_ON_TIME)
         
                 
                 
                 
        If site1 = 1 And TestResult1 <> "PASS" Then
            
            MSComm1.Output = ChipName ' trans strat test signal to TEST PC
            MSComm1.InBufferCount = 0
            MSComm1.InputLen = 0
            RTWaitForTest1 = Timer ' wait for timer  and test result
         End If
            
         If site2 = 1 And TestResult2 <> "PASS" Then
             
            MSComm2.Output = ChipName    ' trans strat test signal to TEST PC
            MSComm2.InBufferCount = 0
            MSComm2.InputLen = 0
            RTWaitForTest2 = Timer
         End If
         
        
         Print "send begin test signal to test"
         
         '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Wait for Response from PC Tester
                
            Do
                DoEvents
                
                If site1 = 1 And TestResult1 <> "PASS" Then
                    If RTTestStop1 = 0 Then
                        RTTestResult1 = MSComm1.Input
                        
                        RTTestCycleTime1 = Timer - RTWaitForTest1
                        
                        If (RTTestResult1 <> "" Or RTTestCycleTime1 > WAIT_TEST_CYCLE_OUT) Then
                        
                            RTTestCounter1 = RTTestCounter1 + 1 ' Allen Debug
                            RTTestStop1 = 1
                        End If
                    End If
                
                Else
                    RTTestStop1 = 1
                
                End If
                '========================
                
                If site2 = 1 And TestResult2 <> "PASS" Then
                    If RTTestStop2 = 0 Then
                        RTTestResult2 = MSComm2.Input
                        
                        RTTestCycleTime2 = Timer - RTWaitForTest2
                        
                        If (RTTestResult2 <> "" Or RTTestCycleTime2 > WAIT_TEST_CYCLE_OUT) Then
                            RTTestStop2 = 1
                            RTTestCounter2 = RTTestCounter2 + 1
                        End If
                    End If
                
                Else
                    RTTestStop2 = 1
                
                
                End If
            
            
            Loop Until (RTTestStop1 = 1) And (RTTestStop2 = 1)
                           
                If site1 = 1 Then
                    Print "RTTestResult1= "; RTTestResult1
                End If
                
                If site2 = 1 Then
                    Print "RTTestResult2= "; RTTestResult2
                End If
                
               
                
                If RTTestCycleTime1 > WAIT_TEST_CYCLE_OUT And site1 = 1 Then
                
                    RTWaitTestTimeOut1 = 1
                    RTWaitTestTimeOutCounter1 = RTWaitTestTimeOutCounter1 + 1
                End If
                
                If RTTestCycleTime2 > WAIT_TEST_CYCLE_OUT And site2 = 1 Then
                
                    RTWaitTestTimeOut2 = 1
                    RTWaitTestTimeOutCounter2 = RTWaitTestTimeOutCounter2 + 1
                End If
                 
                 
                 
                     
                 
End If  '////////////////////// Test end
err:              '  Testing Loop end
End Sub

Private Sub LockOption()
  Combo1.Enabled = False
  Combo2.Enabled = False
  Option1.Enabled = False
  Option2.Enabled = False
  Option3.Enabled = False
  Check1.Enabled = False
  Check2.Enabled = False
  Check3.Enabled = False
  Check4.Enabled = False
  Check5.Enabled = False
  Check6.Enabled = False
  NoCardTest.Enabled = False
  Check7.Enabled = False
  ReportCheck.Enabled = False
  LoopTestCheck.Enabled = False
  Command1.Enabled = False
  Command3.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
End Sub
Private Sub UnlockOption()
  Combo1.Enabled = True
  Combo2.Enabled = True
  Option1.Enabled = True
  Option2.Enabled = True
  Option3.Enabled = True
  Check1.Enabled = True
  Check2.Enabled = True
  Check3.Enabled = True
  Check4.Enabled = True
  Check5.Enabled = True
  Check6.Enabled = True
  NoCardTest.Enabled = True
  Check7.Enabled = True
  ReportCheck.Enabled = True
  LoopTestCheck.Enabled = True
  Command1.Enabled = True
  Command3.Enabled = True
  Command4.Enabled = True
  Command5.Enabled = True
  Command6.Enabled = True
  Command7.Enabled = True
End Sub
